using System.Windows;
using Connector.Desktop.Mvvm;

namespace Connector.Desktop.Features.Connector;

// "Коннектор" top-level domain module: the login/heartbeat FRONT-END (token connect + heartbeat-state line + app
// self-update card + the Tekla-sync mirror). It is a top-level domain (its own tab), like Structura/VPN — NOT a
// Tekla sub-module. The connect/heartbeat ENGINE stays in MainWindow; this module owns ONLY the View, which calls
// the engine through the injected IConnectorHost seam. The shell drives the View's controls through the public
// pass-throughs below (replacing the engine's former direct control writes — the engine LOGIC is unchanged).
public sealed class ConnectorModule : IFeatureModule
{
    private readonly ConnectorView _view;

    // The IConnectorHost is REQUIRED (the view's handlers cannot function without the shell seam). MainWindow
    // implements it and passes `this` when registering the module in ShellViewModel.
    public ConnectorModule(IConnectorHost host)
    {
        _view = new ConnectorView(host);
    }

    public string Title => "Коннектор";
    public FrameworkElement View => _view;

    // ===== Public pass-throughs the shell uses (replace the engine's old direct control writes) =========

    // Replaces MainWindow.UpdateRunStateUi's control writes. The shell still owns _isRunning and computes the
    // text/brush/enabled flags exactly as before, then pushes them here.
    public void SetRunState(string text, System.Windows.Media.Brush brush, bool startEnabled, bool stopEnabled) =>
        _view.SetRunState(text, brush, startEnabled, stopEnabled);

    // Replaces every `UpdateStateTextBlock.Text = ...` in the shell's update flow.
    public void SetUpdateState(string text) => _view.SetUpdateState(text);

    // Replaces MainWindow.UpdateActionButtonUi.
    public void SetUpdateAction(string content, bool isPrimaryStyle) => _view.SetUpdateAction(content, isPrimaryStyle);

    // Replaces the engine's `UpdateActionButton.IsEnabled = ...` writes.
    public void SetUpdateActionEnabled(bool enabled) => _view.SetUpdateActionEnabled(enabled);

    // Replaces MainWindow.OnTeklaStatusChanged's two cross-tab control writes (the header write stays in the shell).
    // MainWindow.OnTeklaStatusChanged calls this; the header write (HeaderFirmStatusTextBlock) stays inline.
    public void SetTeklaMirror(string text, System.Windows.Media.Brush brush, string buttonContent, bool isSyncStyle) =>
        _view.SetTeklaMirror(text, brush, buttonContent, isSyncStyle);

    // Replaces MainWindow.LoadSettingsToUi's writes to the moved (Collapsed) controls + the token PasswordBox.
    public void LoadFromSettings(string decryptedToken) => _view.LoadFromSettings(decryptedToken);

    // Replaces the SACRED ConnectByTokenInternalAsync's six post-connect control writes.
    public void ApplyConnectResult(string deviceId, int heartbeatSeconds, string updateManifestUrl, string smbLogin, string smbSharePath) =>
        _view.ApplyConnectResult(deviceId, heartbeatSeconds, updateManifestUrl, smbLogin, smbSharePath);

    // Like StandardModule.FlushPathEdits: capture the typed token from the moved PasswordBox so ReadSettingsFromUi can
    // source it from _settings. The shell calls this at the top of ApplyAndPersist. Returns the trimmed token.
    public string FlushEditsToSettings() => _view.GetTokenText();

    // Mirror CheckUpdatesAsync's normalised-URL echo write + ConnectSmbInternalAsync's chosen-login write.
    public void SetUpdateManifestUrl(string url) => _view.SetUpdateManifestUrl(url);
    public void SetSmbLogin(string login) => _view.SetSmbLogin(login);

    // Activated => no engine refresh needed (the engine pushes state on connect/heartbeat/update ticks). Kept to
    // satisfy IFeatureModule; mirrors the other top-level modules' OnActivatedAsync no-op-style refresh.
    public Task OnActivatedAsync() => Task.CompletedTask;
}
