using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Connector.Desktop.Services;

public sealed class UpdateService
{
    private readonly HttpClient _http;

    public UpdateService(HttpClient http)
    {
        _http = http;
    }

    public Version CurrentVersion => Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0, 0);

    public async Task<UpdateManifest?> TryGetUpdateAsync(string manifestUrl, CancellationToken ct)
    {
        if (!Uri.TryCreate(manifestUrl, UriKind.Absolute, out var manifestUri) ||
            !string.Equals(manifestUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        using var req = new HttpRequestMessage(HttpMethod.Get, manifestUri);
        using var res = await _http.SendAsync(req, ct);
        if (!res.IsSuccessStatusCode)
        {
            return null;
        }

        var body = await res.Content.ReadAsStringAsync(ct);
        var manifest = JsonSerializer.Deserialize<UpdateManifest>(body, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        if (manifest is null ||
            string.IsNullOrWhiteSpace(manifest.Version) ||
            string.IsNullOrWhiteSpace(manifest.MsiUrl) ||
            !IsValidSha256(manifest.Sha256))
        {
            return null;
        }

        return manifest;
    }

    public bool IsUpdateAvailable(UpdateManifest manifest)
    {
        return Version.TryParse(manifest.Version, out var remote) && remote > CurrentVersion;
    }

    public async Task<string> DownloadInstallerAsync(UpdateManifest manifest, CancellationToken ct)
    {
        if (!Uri.TryCreate(manifest.MsiUrl, UriKind.Absolute, out var downloadUri) ||
            !string.Equals(downloadUri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Адрес установщика должен использовать HTTPS.");
        }

        var expectedSha256 = NormalizeSha256(manifest.Sha256);
        if (!IsValidSha256(expectedSha256))
        {
            throw new InvalidOperationException("Manifest обновления не содержит корректную SHA-256 сумму.");
        }

        var dir = Path.Combine(Path.GetTempPath(), "StructuraConnectorUpdates");
        Directory.CreateDirectory(dir);

        var fileName = $"StructuraConnector_{manifest.Version}.msi";
        var filePath = Path.Combine(dir, fileName);

        using var req = new HttpRequestMessage(HttpMethod.Get, downloadUri);
        using var res = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        res.EnsureSuccessStatusCode();

        try
        {
            await using (var stream = await res.Content.ReadAsStreamAsync(ct))
            await using (var fs = File.Create(filePath))
            {
                await stream.CopyToAsync(fs, ct);
            }

            var actualSha256 = ComputeSha256(filePath);
            if (!string.Equals(actualSha256, expectedSha256, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException(
                    $"Проверка установщика не пройдена: ожидалась SHA-256 {expectedSha256}, получена {actualSha256}.");
            }
        }
        catch
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch
            {
                // Best-effort cleanup; the original verification/download error is more important.
            }
            throw;
        }

        return filePath;
    }

    public static string ComputeSha256(string filePath)
    {
        using var stream = File.OpenRead(filePath);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static string NormalizeSha256(string value)
    {
        var normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
        return normalized.StartsWith("sha256:", StringComparison.Ordinal)
            ? normalized["sha256:".Length..]
            : normalized;
    }

    private static bool IsValidSha256(string value)
    {
        var normalized = NormalizeSha256(value);
        return normalized.Length == 64 && normalized.All(Uri.IsHexDigit);
    }

    public static void RunInstaller(string msiPath)
    {
        var currentProcessId = Environment.ProcessId;
        var escapedMsiPath = msiPath.Replace("'", "''");
        var launcherScriptPath = Path.Combine(
            Path.GetTempPath(),
            "StructuraConnectorInstall_" + Guid.NewGuid().ToString("N") + ".ps1");

        var launcherScript = string.Join(Environment.NewLine, new[]
        {
            "$pidToWait = " + currentProcessId,
            "$msiPath = '" + escapedMsiPath + "'",
            "$scriptPath = $MyInvocation.MyCommand.Path",
            "",
            "try {",
            "    Wait-Process -Id $pidToWait -Timeout 60 -ErrorAction SilentlyContinue",
            "} catch {",
            "}",
            "",
            "Start-Sleep -Milliseconds 700",
            "Start-Process -FilePath 'msiexec.exe' -Verb RunAs -ArgumentList ('/i \"' + $msiPath + '\"')",
            "Start-Sleep -Seconds 2",
            "Remove-Item -LiteralPath $scriptPath -Force -ErrorAction SilentlyContinue",
            ""
        });

        File.WriteAllText(launcherScriptPath, launcherScript, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        var psi = new ProcessStartInfo("powershell.exe")
        {
            UseShellExecute = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{launcherScriptPath}\""
        };
        Process.Start(psi);
    }
}

public sealed class UpdateManifest
{
    public string Version { get; set; } = "";
    public string MsiUrl { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public string Notes { get; set; } = "";
}
