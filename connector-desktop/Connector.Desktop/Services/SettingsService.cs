using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Connector.Desktop.Models;

namespace Connector.Desktop.Services;

public sealed class SettingsService
{
    private readonly string _settingsPath;

    public SettingsService()
    {
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ConnectorAgentDesktop");
        Directory.CreateDirectory(root);
        _settingsPath = Path.Combine(root, "settings.json");
    }

    public AppSettings Load()
    {
        if (!File.Exists(_settingsPath))
        {
            return new AppSettings
            {
                DeviceId = "pc-" + Environment.MachineName.ToLowerInvariant()
            };
        }

        var raw = File.ReadAllText(_settingsPath, Encoding.UTF8);
        var settings = JsonSerializer.Deserialize<AppSettings>(raw) ?? new AppSettings();
        if (string.IsNullOrWhiteSpace(settings.DeviceId))
        {
            settings.DeviceId = "pc-" + Environment.MachineName.ToLowerInvariant();
        }

        if (settings.HeartbeatSeconds < 10)
        {
            settings.HeartbeatSeconds = 60;
        }

        if (string.IsNullOrWhiteSpace(settings.ServerUrl))
        {
            settings.ServerUrl = "https://server.structura-most.ru";
        }

        if (string.IsNullOrWhiteSpace(settings.UpdateManifestUrl))
        {
            settings.UpdateManifestUrl = "https://server.structura-most.ru/updates/latest.json";
        }

        if (string.IsNullOrWhiteSpace(settings.SmbSharePath))
        {
            settings.SmbSharePath = @"\\62.113.36.107\BIM_Models";
        }

        return settings;
    }

    public void Save(AppSettings settings)
    {
        var json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_settingsPath, json, Encoding.UTF8);
    }

    public static string EncryptToken(string token)
    {
        var data = Encoding.UTF8.GetBytes(token);
        var encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(encrypted);
    }

    public static string DecryptToken(string cipherBase64)
    {
        if (string.IsNullOrWhiteSpace(cipherBase64))
        {
            return string.Empty;
        }

        var encrypted = Convert.FromBase64String(cipherBase64);
        var plain = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(plain);
    }

    public string SettingsPath => _settingsPath;
}
