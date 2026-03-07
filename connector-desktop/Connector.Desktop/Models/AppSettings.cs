namespace Connector.Desktop.Models;

public sealed class AppSettings
{
    public string ServerUrl { get; set; } = "https://server.structura-most.ru";
    public string UpdateManifestUrl { get; set; } = "https://server.structura-most.ru/updates/latest.json";
    public string DeviceId { get; set; } = "";
    public string TokenCipherBase64 { get; set; } = "";
    public string SmbLogin { get; set; } = "";
    public string SmbPasswordCipherBase64 { get; set; } = "";
    public string SmbSharePath { get; set; } = @"\\62.113.36.107\BIM_Models";
    public int HeartbeatSeconds { get; set; } = 60;
    public bool AutoStart { get; set; } = true;
}
