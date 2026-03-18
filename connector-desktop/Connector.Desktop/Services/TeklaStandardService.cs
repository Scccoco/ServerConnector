using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using Connector.Desktop.Models;

namespace Connector.Desktop.Services;

public sealed class TeklaStandardService
{
    private readonly HttpClient _http;

    public TeklaStandardService(HttpClient http)
    {
        _http = http;
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ConnectorAgentDesktop");
        Directory.CreateDirectory(root);
        LogFilePath = Path.Combine(root, "tekla-standard.log");
    }

    public string LogFilePath { get; }

    public string ResolveGitExecutable()
    {
        var baseDir = AppContext.BaseDirectory;
        var candidates = new[]
        {
            Path.Combine(baseDir, "tools", "git", "bin", "git.exe"),
            Path.Combine(baseDir, "tools", "git", "cmd", "git.exe")
        };

        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return string.Empty;
    }

    public bool CheckGitAvailability(out string gitPath, out string details)
    {
        gitPath = ResolveGitExecutable();
        details = string.Empty;

        if (string.IsNullOrWhiteSpace(gitPath))
        {
            details = "Встроенный git не найден. Ожидается tools\\git\\bin\\git.exe или tools\\git\\cmd\\git.exe";
            return false;
        }

        var probeWorkDir = Path.GetTempPath();
        var ok = TryRunGit(gitPath, "--version", probeWorkDir, out var stdout, out var stderr);
        if (ok)
        {
            details = string.IsNullOrWhiteSpace(stdout) ? "git доступен" : stdout.Trim();
            return true;
        }

        details = string.IsNullOrWhiteSpace(stderr) ? stdout.Trim() : stderr.Trim();
        return false;
    }

    public async Task<TeklaStandardManifest?> TryGetManifestAsync(string manifestUrl, CancellationToken ct)
    {
        if (!Uri.TryCreate(manifestUrl, UriKind.Absolute, out _))
        {
            return null;
        }

        using var req = new HttpRequestMessage(HttpMethod.Get, manifestUrl);
        using var res = await _http.SendAsync(req, ct);
        if (!res.IsSuccessStatusCode)
        {
            return null;
        }

        var body = await res.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(body);
        var root = doc.RootElement;

        var revision =
            GetString(root, "revision") ??
            GetString(root, "target_revision") ??
            GetString(root, "targetRevision") ??
            GetString(root, "version");

        if (string.IsNullOrWhiteSpace(revision))
        {
            return null;
        }

        return new TeklaStandardManifest
        {
            Version = GetString(root, "version")?.Trim() ?? string.Empty,
            Revision = revision.Trim(),
            Notes = GetString(root, "notes") ?? string.Empty,
            RepoUrl = GetString(root, "repo_url") ?? string.Empty,
            RepoRef =
                GetString(root, "repo_ref") ??
                GetString(root, "git_ref") ??
                GetString(root, "revision") ??
                string.Empty
        };
    }

    public bool IsUpdateAvailable(AppSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TeklaStandardTargetRevision))
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(settings.TeklaStandardInstalledRevision))
        {
            return true;
        }

        return !string.Equals(
            settings.TeklaStandardInstalledRevision.Trim(),
            settings.TeklaStandardTargetRevision.Trim(),
            StringComparison.OrdinalIgnoreCase);
    }

    public TeklaApplyResult ApplyPendingGitUpdate(AppSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.TeklaStandardTargetRevision))
        {
            return TeklaApplyResult.Fail("Нет целевой ревизии для применения.");
        }

        var targetRevision = settings.TeklaStandardTargetRevision.Trim();
        var repoUrl = settings.TeklaStandardRepoUrl.Trim();
        var repoRef = settings.TeklaStandardRepoRef.Trim();
        var localPath = settings.TeklaStandardLocalPath.Trim();

        if (string.IsNullOrWhiteSpace(localPath))
        {
            return TeklaApplyResult.Fail("Не задан локальный путь папки фирмы Tekla.");
        }

        if (string.IsNullOrWhiteSpace(repoUrl))
        {
            return TeklaApplyResult.Fail("В manifest отсутствует repo_url для обновления через git.");
        }

        if (string.IsNullOrWhiteSpace(repoRef))
        {
            return TeklaApplyResult.Fail("В manifest отсутствует repo_ref для обновления через git.");
        }

        try
        {
            Directory.CreateDirectory(localPath);
        }
        catch (Exception ex)
        {
            return TeklaApplyResult.Fail("Не удалось создать локальную папку: " + ex.Message);
        }

        var gitExe = ResolveGitExecutable();
        if (string.IsNullOrWhiteSpace(gitExe))
        {
            return TeklaApplyResult.Fail("Встроенный git не найден в каталоге приложения (tools\\git). Обновите Connector.");
        }
        if (!TryRunGit(gitExe, "--version", localPath, out var versionOut, out var versionErr))
        {
            var reason = string.IsNullOrWhiteSpace(versionErr) ? versionOut : versionErr;
            return TeklaApplyResult.Fail("Git недоступен: " + reason);
        }

        var gitDir = Path.Combine(localPath, ".git");
        var isFreshClone = false;
        if (!Directory.Exists(gitDir))
        {
            var parent = Directory.GetParent(localPath)?.FullName;
            if (string.IsNullOrWhiteSpace(parent))
            {
                return TeklaApplyResult.Fail("Некорректный локальный путь папки фирмы.");
            }

            Directory.CreateDirectory(parent);

            if (!TryRunGit(gitExe, $"clone \"{repoUrl}\" \"{localPath}\"", parent, out var cloneOut, out var cloneErr))
            {
                var reason = string.IsNullOrWhiteSpace(cloneErr) ? cloneOut : cloneErr;
                return TeklaApplyResult.Fail("Не удалось клонировать репозиторий: " + reason);
            }

            isFreshClone = true;
        }

        var hadHead = TryGetHeadRevision(gitExe, localPath, out var previousHead);

        if (!TryRunGit(gitExe, $"fetch origin {repoRef} --depth 1", localPath, out var fetchOut, out var fetchErr))
        {
            var reason = string.IsNullOrWhiteSpace(fetchErr) ? fetchOut : fetchErr;
            return TeklaApplyResult.Fail("Не удалось получить обновление: " + reason);
        }

        if (!TryRunGit(gitExe, "checkout -f FETCH_HEAD", localPath, out var checkoutOut, out var checkoutErr))
        {
            var reason = string.IsNullOrWhiteSpace(checkoutErr) ? checkoutOut : checkoutErr;

            if (!isFreshClone && hadHead && !string.IsNullOrWhiteSpace(previousHead))
            {
                TryRunGit(gitExe, $"checkout -f {previousHead}", localPath, out _, out _);
                TryRunGit(gitExe, "clean -fd", localPath, out _, out _);
            }

            return TeklaApplyResult.Fail("Не удалось применить обновление: " + reason);
        }

        TryRunGit(gitExe, "clean -fd", localPath, out _, out _);

        settings.TeklaStandardInstalledRevision = targetRevision;
        settings.TeklaStandardLastSuccessUtc = DateTimeOffset.UtcNow;
        settings.TeklaStandardPendingAfterClose = false;

        var message = "Применено обновление Стандарт Tekla. Установлена ревизия " + targetRevision + ".";
        return TeklaApplyResult.Success(message);
    }

    public void AppendLog(string message)
    {
        var line = $"[{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
        File.AppendAllText(LogFilePath, line);
    }

    private static string? GetString(JsonElement root, string propertyName)
    {
        foreach (var property in root.EnumerateObject())
        {
            if (!string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (property.Value.ValueKind == JsonValueKind.String)
            {
                return property.Value.GetString();
            }

            if (property.Value.ValueKind == JsonValueKind.Number)
            {
                return property.Value.GetRawText();
            }
        }

        return null;
    }

    private static bool TryRunGit(string gitExe, string arguments, string workDir, out string stdout, out string stderr)
    {
        stdout = string.Empty;
        stderr = string.Empty;

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = gitExe,
                Arguments = arguments,
                WorkingDirectory = workDir,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = Process.Start(startInfo);
            if (process is null)
            {
                stderr = "Не удалось запустить процесс git.";
                return false;
            }

            stdout = process.StandardOutput.ReadToEnd();
            stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            return process.ExitCode == 0;
        }
        catch (Exception ex)
        {
            stderr = ex.Message;
            return false;
        }
    }

    private static bool TryGetHeadRevision(string gitExe, string workDir, out string revision)
    {
        revision = string.Empty;
        if (!TryRunGit(gitExe, "rev-parse --verify HEAD", workDir, out var outText, out var errText))
        {
            return false;
        }

        var value = (outText ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            value = (errText ?? string.Empty).Trim();
        }

        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        revision = value;
        return true;
    }

}

public sealed class TeklaStandardManifest
{
    public string Version { get; set; } = "";
    public string Revision { get; set; } = "";
    public string Notes { get; set; } = "";
    public string RepoUrl { get; set; } = "";
    public string RepoRef { get; set; } = "";
}

public sealed class TeklaApplyResult
{
    public bool IsSuccess { get; init; }
    public string Message { get; init; } = "";

    public static TeklaApplyResult Success(string message) => new() { IsSuccess = true, Message = message };
    public static TeklaApplyResult Fail(string message) => new() { IsSuccess = false, Message = message };
}
