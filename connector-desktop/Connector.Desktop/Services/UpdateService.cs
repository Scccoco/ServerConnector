using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
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
        var manifest = JsonSerializer.Deserialize<UpdateManifest>(body, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        if (manifest is null || string.IsNullOrWhiteSpace(manifest.Version) || string.IsNullOrWhiteSpace(manifest.MsiUrl))
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
        var dir = Path.Combine(Path.GetTempPath(), "StructuraConnectorUpdates");
        Directory.CreateDirectory(dir);

        var fileName = $"StructuraConnector_{manifest.Version}.msi";
        var filePath = Path.Combine(dir, fileName);

        using var req = new HttpRequestMessage(HttpMethod.Get, manifest.MsiUrl);
        using var res = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, ct);
        res.EnsureSuccessStatusCode();

        await using (var stream = await res.Content.ReadAsStreamAsync(ct))
        await using (var fs = File.Create(filePath))
        {
            await stream.CopyToAsync(fs, ct);
        }

        return filePath;
    }

    public static void RunInstaller(string msiPath)
    {
        var psi = new ProcessStartInfo("msiexec.exe")
        {
            UseShellExecute = true,
            Verb = "runas",
            Arguments = $"/i \"{msiPath}\""
        };
        Process.Start(psi);
    }
}

public sealed class UpdateManifest
{
    public string Version { get; set; } = "";
    public string MsiUrl { get; set; } = "";
    public string Notes { get; set; } = "";
}
