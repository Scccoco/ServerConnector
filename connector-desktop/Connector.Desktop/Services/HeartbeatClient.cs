using System.Net.Http;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Connector.Desktop.Services;

public sealed class HeartbeatClient
{
    private readonly HttpClient _http;
    private static readonly TimeSpan HealthTimeout = TimeSpan.FromSeconds(15);
    private static readonly TimeSpan HeartbeatTimeout = TimeSpan.FromSeconds(20);
    private static readonly TimeSpan BootstrapTimeout = TimeSpan.FromSeconds(90);
    private static readonly string[] PublicIpProviders =
    {
        "https://api.ipify.org",
        "https://ifconfig.me/ip",
        "https://checkip.amazonaws.com"
    };

    public HeartbeatClient(HttpClient http)
    {
        _http = http;
    }

    public async Task<string> ResolvePublicIpAsync(CancellationToken ct)
    {
        var errors = new List<string>();

        foreach (var provider in PublicIpProviders)
        {
            using var providerCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
            providerCts.CancelAfter(TimeSpan.FromSeconds(8));

            try
            {
                using var req = new HttpRequestMessage(HttpMethod.Get, provider);
                using var res = await _http.SendAsync(req, providerCts.Token);
                res.EnsureSuccessStatusCode();

                var ipText = (await res.Content.ReadAsStringAsync(providerCts.Token)).Trim();
                if (IPAddress.TryParse(ipText, out _))
                {
                    return ipText;
                }

                errors.Add($"{provider}: invalid response '{ipText}'");
            }
            catch (OperationCanceledException)
            {
                errors.Add($"{provider}: timeout");
            }
            catch (Exception ex)
            {
                errors.Add($"{provider}: {ex.Message}");
            }
        }

        throw new InvalidOperationException(
            "Не удалось определить внешний IP. Проверьте доступ к интернету/прокси для HTTPS. " +
            "Детали: " + string.Join("; ", errors));
    }

    public async Task CheckServerHealthAsync(string serverUrl, CancellationToken ct)
    {
        var url = serverUrl.TrimEnd('/') + "/health";
        using var reqCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        reqCts.CancelAfter(HealthTimeout);
        using var req = new HttpRequestMessage(HttpMethod.Get, url);
        using var res = await _http.SendAsync(req, reqCts.Token);
        var body = await res.Content.ReadAsStringAsync(reqCts.Token);
        if (!res.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"HTTP {(int)res.StatusCode}: {body}");
        }

        if (!body.Contains("\"ok\":true", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Сервер ответил без ожидаемого признака здоровья (ok=true).");
        }
    }

    public async Task SendHeartbeatAsync(string serverUrl, string deviceId, string token, string sessionId, CancellationToken ct)
    {
        var payload = new
        {
            device_id = deviceId,
            hostname = Environment.MachineName,
            agent_version = "1.0.0"
        };

        var url = serverUrl.TrimEnd('/') + "/heartbeat";
        var json = JsonSerializer.Serialize(payload);
        using var reqCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        reqCts.CancelAfter(HeartbeatTimeout);
        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Add("X-Device-Token", token);
        if (!string.IsNullOrWhiteSpace(sessionId))
        {
            req.Headers.Add("X-Device-Session", sessionId);
        }
        req.Content = new StringContent(json, Encoding.UTF8, "application/json");

        using var res = await _http.SendAsync(req, reqCts.Token);
        var body = await res.Content.ReadAsStringAsync(reqCts.Token);
        if (!res.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"HTTP {(int)res.StatusCode}: {body}");
        }
    }

    public async Task<BootstrapResponse> BootstrapAsync(string serverUrl, string token, CancellationToken ct)
    {
        var url = serverUrl.TrimEnd('/') + "/connect/bootstrap";
        var payload = new
        {
            hostname = Environment.MachineName,
            agent_version = "desktop-1.0.0"
        };

        var json = JsonSerializer.Serialize(payload);
        using var reqCts = CancellationTokenSource.CreateLinkedTokenSource(ct);
        reqCts.CancelAfter(BootstrapTimeout);
        using var req = new HttpRequestMessage(HttpMethod.Post, url);
        req.Headers.Add("X-Device-Token", token);
        req.Content = new StringContent(json, Encoding.UTF8, "application/json");

        using var res = await _http.SendAsync(req, reqCts.Token);
        var body = await res.Content.ReadAsStringAsync(reqCts.Token);
        if (!res.IsSuccessStatusCode)
        {
            throw new InvalidOperationException($"HTTP {(int)res.StatusCode}: {body}");
        }

        var data = JsonSerializer.Deserialize<BootstrapResponse>(body, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });
        if (data is null || !data.Ok)
        {
            throw new InvalidOperationException("Некорректный ответ bootstrap от сервера.");
        }

        return data;
    }
}

public sealed class BootstrapResponse
{
    [JsonPropertyName("ok")]
    public bool Ok { get; set; }

    [JsonPropertyName("session_id")]
    public string SessionId { get; set; } = "";

    [JsonPropertyName("device_id")]
    public string DeviceId { get; set; } = "";

    [JsonPropertyName("issued_to")]
    public string IssuedTo { get; set; } = "";

    [JsonPropertyName("public_ip")]
    public string PublicIp { get; set; } = "";

    [JsonPropertyName("heartbeat_seconds")]
    public int HeartbeatSeconds { get; set; } = 60;

    [JsonPropertyName("update_manifest_url")]
    public string UpdateManifestUrl { get; set; } = "";

    [JsonPropertyName("smb_access")]
    public BootstrapSmbAccess SmbAccess { get; set; } = new();
}

public sealed class BootstrapSmbAccess
{
    [JsonPropertyName("login")]
    public string Login { get; set; } = "";

    [JsonPropertyName("username")]
    public string Username { get; set; } = "";

    [JsonPropertyName("password")]
    public string Password { get; set; } = "";

    [JsonPropertyName("share_unc")]
    public string ShareUnc { get; set; } = "";

    [JsonPropertyName("share_path")]
    public string SharePath { get; set; } = "";
}
