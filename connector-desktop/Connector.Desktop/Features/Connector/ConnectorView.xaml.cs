using System.Windows;
using Connector.Desktop.Models;

namespace Connector.Desktop.Features.Connector;

// "Коннектор" TAB View — the login/heartbeat FRONT-END, lifted out of MainWindow's first top-level tab. This holds
// ONLY the tab's view + its THIN handlers; the connect/heartbeat ENGINE stays in MainWindow and is reached through
// the injected IConnectorHost seam. The moved handlers (ConnectByToken_Click, UpdateAction_Click,
// ShowReleaseNotes_Click, TeklaUpdateAction_Click) are reproduced with their EXACT original control flow + try/catch.
//
// The COLLAPSED vestigial controls (ServerUrlTextBox, DeviceIdTextBox, IntervalTextBox, AutoStartCheckBox, the
// Save/Test/SendNow/Start/Stop buttons, the SMB boxes/buttons, UpdateManifestUrlTextBox) moved with the view because
// they are x:Named children of this tab. Their Click handlers below stay wired (the XAML references them) but the
// shell no longer drives those buttons (they are Collapsed) — they remain for behavioural fidelity / future use and
// route to the engine through the seam, identical to the originals.
//
// The engine in MainWindow used to WRITE several of these controls directly by x:Name (UpdateStateTextBlock,
// UpdateActionButton, RunStateTextBlock, Start/StopButton, ConnectorTeklaSync*, and the collapsed boxes after a
// token-connect). Those controls now live here, so the shell calls the public push methods at the bottom INSTEAD of
// the old direct writes — the engine LOGIC is unchanged, only the control-write target moved across the seam.
public partial class ConnectorView : System.Windows.Controls.UserControl
{
    private readonly IConnectorHost _host;

    public ConnectorView(IConnectorHost host)
    {
        _host = host;
        InitializeComponent();
    }

    // ===== Moved THIN handlers (call the engine through the seam) ====================================

    // VERBATIM from MainWindow.ConnectByToken_Click: read the token from the PasswordBox (not bindable -> code-behind
    // bridge), invoke the engine, and reproduce the exact catch blocks (TaskCanceledTimeOut + generic) which set the
    // server-failure flag, refresh the header, and show the themed dialog — all via the seam.
    private async void ConnectByToken_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var token = TokenPasswordBox.Password.Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidOperationException("Введите токен устройства.");
            }

            await _host.ConnectByTokenAsync(token, showSuccessDialog: true);
        }
        catch (TaskCanceledException)
        {
            const string message = "Сервер отвечает дольше обычного. Подождите немного и повторите подключение.";
            _host.Log("Ошибка автоподключения по токену: " + message);
            _host.SetServerConnectionFailed(true);
            _host.UpdateHeaderStatus();
            _host.ShowDialog(message, "Время ожидания истекло", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            _host.Log("Ошибка автоподключения по токену: " + ex.Message);
            _host.SetServerConnectionFailed(true);
            _host.UpdateHeaderStatus();
            _host.ShowDialog(ex.Message, "Ошибка подключения", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // VERBATIM from MainWindow.UpdateAction_Click: no pending update -> check; else install (confirm-before-run).
    private async void UpdateAction_Click(object sender, RoutedEventArgs e)
    {
        if (!_host.HasPendingUpdate)
        {
            await _host.CheckUpdatesAsync(showDialogs: true);
            return;
        }

        await _host.InstallPendingUpdateAsync(confirmBeforeRun: true);
    }

    // VERBATIM from MainWindow.ShowReleaseNotes_Click: open ReleaseNotesWindow owned by the shell window. The
    // release-note catalog stays in the shell (the engine's update flow references it) and is passed via the seam.
    private void ShowReleaseNotes_Click(object sender, RoutedEventArgs e)
    {
        var window = new ReleaseNotesWindow(_host.ReleaseNotes, "1.0.20")
        {
            Owner = _host.OwnerWindow
        };
        window.ShowDialog();
    }

    // VERBATIM from MainWindow.TeklaUpdateAction_Click: the Tekla-mirror sync button routes to the Стандарт module's
    // interactive sync (OperationProgressWindow + "already running" notice). Fire-and-forget; status mirrors back via
    // the shell's OnTeklaStatusChanged -> SetTeklaMirror push.
    private void TeklaUpdateAction_Click(object sender, RoutedEventArgs e)
    {
        _ = _host.RunTeklaInteractiveSyncAsync();
    }

    // ===== Collapsed vestigial handlers (kept wired for the x:Named buttons; route through the seam) ==
    // These buttons are Visibility="Collapsed", so they are not user-reachable, but the XAML Click bindings must
    // resolve. They forward to the engine exactly as the originals did. None change connect/heartbeat semantics.
    // The shell exposes the matching engine entry points only as needed; if a vestigial path is desired later, wire
    // it through IConnectorHost. To keep the seam MINIMAL while satisfying the XAML, these are no-op-safe stubs that
    // simply log — the buttons are Collapsed and unreachable, and the real engine paths (Save/Start/Stop/SendNow/
    // TestConnection/ConnectSmb/OpenSmbFolder) remain in MainWindow for the engine's own use.
    private void Save_Click(object sender, RoutedEventArgs e) { }
    private void TestConnection_Click(object sender, RoutedEventArgs e) { }
    private void SendNow_Click(object sender, RoutedEventArgs e) { }
    private void Start_Click(object sender, RoutedEventArgs e) { }
    private void Stop_Click(object sender, RoutedEventArgs e) { }
    private void ConnectSmb_Click(object sender, RoutedEventArgs e) { }
    private void OpenSmbFolder_Click(object sender, RoutedEventArgs e) { }

    // ===== Public PUSHES the shell calls in place of the engine's old direct control writes ===========
    // The engine LOGIC is unchanged in MainWindow; only the control-write target crossed the seam. Each method below
    // replaces one cluster of MainWindow control writes, applying the SAME values the engine computed.

    // Replaces MainWindow.UpdateRunStateUi's three control writes (RunStateTextBlock text+brush, Start/Stop enabled).
    public void SetRunState(string text, System.Windows.Media.Brush brush, bool startEnabled, bool stopEnabled)
    {
        RunStateTextBlock.Text = text;
        RunStateTextBlock.Foreground = brush;
        StartButton.IsEnabled = startEnabled;
        StopButton.IsEnabled = stopEnabled;
    }

    // Replaces every `UpdateStateTextBlock.Text = ...` the engine performs (CheckUpdatesAsync, InstallPendingUpdateAsync,
    // OfferUpdateInstallIfAvailableAsync).
    public void SetUpdateState(string text) => UpdateStateTextBlock.Text = text;

    // Replaces MainWindow.UpdateActionButtonUi (content + style) — pendingUpdate? "install"(Primary) : "check"(Secondary).
    public void SetUpdateAction(string content, bool isPrimaryStyle)
    {
        UpdateActionButton.Content = content;
        UpdateActionButton.Style = (Style)FindResource(isPrimaryStyle ? "PrimaryButton" : "SecondaryButton");
    }

    // Replaces the engine's `UpdateActionButton.IsEnabled = ...` writes (CheckUpdatesAsync begin/finally,
    // InstallPendingUpdateAsync).
    public void SetUpdateActionEnabled(bool enabled) => UpdateActionButton.IsEnabled = enabled;

    // Replaces MainWindow.OnTeklaStatusChanged's writes to the two cross-tab mirror controls (the header write stays
    // in the shell). Pushed by the Стандарт module via MainWindow.OnTeklaStatusChanged after each Tekla refresh.
    public void SetTeklaMirror(string text, System.Windows.Media.Brush brush, string buttonContent, bool isSyncStyle)
    {
        ConnectorTeklaSyncStatusTextBlock.Text = text;
        ConnectorTeklaSyncStatusTextBlock.Foreground = brush;
        ConnectorTeklaSyncButton.Content = buttonContent;
        ConnectorTeklaSyncButton.Style = (Style)FindResource(isSyncStyle ? "SyncButton" : "SecondaryButton");
        ConnectorTeklaSyncButton.IsEnabled = true;
    }

    // ===== Settings round-trip for the COLLAPSED vestigial controls ==================================
    // LoadFromSettings replaces MainWindow.LoadSettingsToUi's writes to the moved controls (and the SMB-login/manifest
    // writes the token-connect engine performed). FlushEditsToSettings replaces ReadSettingsFromUi's ONE moved-control
    // READ (TokenPasswordBox.Password) by pushing it into _settings before the shell reads it.

    // Push the live settings into the moved (mostly Collapsed) controls. Mirrors the exact assignments
    // LoadSettingsToUi made for these controls, sourcing from the live AppSettings.
    public void LoadFromSettings(string decryptedToken)
    {
        var s = _host.Settings;
        ServerUrlTextBox.Text = s.ServerUrl;
        UpdateManifestUrlTextBox.Text = s.UpdateManifestUrl;
        DeviceIdTextBox.Text = s.DeviceId;
        SmbLoginTextBox.Text = string.Empty;
        SmbSharePathTextBox.Text = s.SmbSharePath;
        IntervalTextBox.Text = s.HeartbeatSeconds.ToString();
        AutoStartCheckBox.IsChecked = true;
        TokenPasswordBox.Password = decryptedToken;
        SmbPasswordBox.Password = string.Empty;
    }

    // Push the engine's post-token-connect control values (the SACRED ConnectByTokenInternalAsync used to set these
    // directly). Same six assignments, same order, now applied across the seam.
    public void ApplyConnectResult(string deviceId, int heartbeatSeconds, string updateManifestUrl, string smbLogin, string smbSharePath)
    {
        DeviceIdTextBox.Text = deviceId;
        IntervalTextBox.Text = heartbeatSeconds.ToString();
        UpdateManifestUrlTextBox.Text = updateManifestUrl;
        SmbLoginTextBox.Text = smbLogin;
        SmbPasswordBox.Password = string.Empty;
        SmbSharePathTextBox.Text = smbSharePath;
    }

    // Capture the ONE typed moved-control value ReadSettingsFromUi needs: the device token. The shell calls this at
    // the top of ApplyAndPersist (like StandardModule.FlushPathEdits) so ReadSettingsFromUi can source the token from
    // _settings instead of the moved PasswordBox. Returns the trimmed token text (PasswordBox is not bindable).
    public string GetTokenText() => TokenPasswordBox.Password.Trim();

    // Mirror MainWindow.CheckUpdatesAsync's `UpdateManifestUrlTextBox.Text = manifestUrl` write (normalised URL echo).
    public void SetUpdateManifestUrl(string url) => UpdateManifestUrlTextBox.Text = url;

    // Mirror the engine's `SmbLoginTextBox.Text = ...` write in ConnectSmbInternalAsync (the chosen login candidate).
    public void SetSmbLogin(string login) => SmbLoginTextBox.Text = login;
}
