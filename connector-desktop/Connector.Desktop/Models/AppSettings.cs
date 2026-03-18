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
    public string TeklaStandardManifestUrl { get; set; } = "https://server.structura-most.ru/updates/tekla/firm/latest.json";
    public string TeklaStandardLocalPath { get; set; } = @"C:\Company\TeklaFirm";
    public string TeklaStandardInstalledVersion { get; set; } = "";
    public string TeklaStandardTargetVersion { get; set; } = "";
    public string TeklaStandardInstalledRevision { get; set; } = "";
    public string TeklaStandardTargetRevision { get; set; } = "";
    public DateTimeOffset? TeklaStandardLastCheckUtc { get; set; }
    public DateTimeOffset? TeklaStandardLastSuccessUtc { get; set; }
    public bool TeklaStandardPendingAfterClose { get; set; }
    public string TeklaStandardLastError { get; set; } = "";
    public string TeklaStandardRepoUrl { get; set; } = "";
    public string TeklaStandardRepoRef { get; set; } = "";
    public string TeklaPublishSourcePath { get; set; } = @"\\62.113.36.107\BIM_Models\Tekla\XS_FIRM";
    public bool IsSystemAdmin { get; set; }
    public bool IsFirmAdmin { get; set; }
}
