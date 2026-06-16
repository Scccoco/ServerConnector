using System.Windows;
using Connector.Desktop.Models;

namespace Connector.Desktop.Features.Connector;

// The seam between the lifted "Коннектор" TAB (ConnectorView — the login/heartbeat front-end) and the shell
// (MainWindow). UNLIKE the Стандарт migration, the ENGINE STAYS in MainWindow: ConnectByTokenInternalAsync,
// SendHeartbeatSafeAsync, Timer_Tick, the heartbeat/update timers, CheckUpdatesAsync, InstallPendingUpdateAsync,
// LoadSettingsToUi/ReadSettingsFromUi/ApplyAndPersist, ConnectSmbInternalAsync — all unchanged. ONLY the tab's
// VIEW + its thin handlers moved into the module; those handlers call the engine through THIS interface.
// MainWindow implements it. Keep this surface MINIMAL — exactly what the moved handlers touch.
public interface IConnectorHost
{
    // ===== Engine entry points the moved handlers invoke (the engine itself is unchanged) =================

    // The login/heartbeat SPINE. ConnectByToken_Click calls this. Satisfied by: MainWindow's
    // ConnectByTokenInternalAsync(token, showSuccessDialog) — the SACRED engine, called VERBATIM.
    Task ConnectByTokenAsync(string token, bool showSuccessDialog);

    // App self-update "check" path. UpdateAction_Click calls this when there is no pending update.
    // Satisfied by: MainWindow CheckUpdatesAsync(showDialogs: true).
    Task CheckUpdatesAsync(bool showDialogs);

    // App self-update "install" path. UpdateAction_Click calls this when an update is pending.
    // Satisfied by: MainWindow InstallPendingUpdateAsync(confirmBeforeRun: true).
    Task InstallPendingUpdateAsync(bool confirmBeforeRun);

    // Is a self-update currently pending (drives UpdateAction_Click's check-vs-install branch, mirroring the old
    // `if (_pendingUpdate is null)`). Satisfied by: MainWindow `_pendingUpdate is not null`.
    bool HasPendingUpdate { get; }

    // The Tekla-mirror sync button (ConnectorTeklaSyncButton). Satisfied by: MainWindow's
    // `_standard?.RunInteractiveSyncAsync()` (the existing Стандарт interactive sync, identical to the Стандарт-tab button).
    Task RunTeklaInteractiveSyncAsync();

    // ===== Primitives the moved handlers' try/catch blocks use ===========================================

    // Journal line. Satisfied by: MainWindow `AppendLog(string)`.
    void Log(string message);

    // Window-owned themed dialog. Satisfied by: MainWindow `ThemedDialogs.Show(this, message, title, buttons, image)`.
    MessageBoxResult ShowDialog(string message, string title, MessageBoxButton buttons, MessageBoxImage image);

    // Owner window for ReleaseNotesWindow (ShowReleaseNotes_Click sets `Owner = host.OwnerWindow`).
    // Satisfied by: `this` (the MainWindow instance).
    Window OwnerWindow { get; }

    // ConnectByToken_Click's catch blocks set the shell's server-failure flag then redraw the header. These two
    // members let the moved handler reproduce that EXACTLY without owning the flag/header (which stay in the shell).
    // Satisfied by: MainWindow `_serverConnectionFailed = value` and `UpdateHeaderStatusUi()` respectively.
    void SetServerConnectionFailed(bool failed);
    void UpdateHeaderStatus();

    // The shared, mutable settings model — used ONLY for the round-trip of the COLLAPSED vestigial controls
    // (DeviceId/Interval/AutoStart/SmbSharePath/UpdateManifestUrl) so the view can render them from the live
    // settings instead of from shell-pushed strings. LIVE getter — MainWindow reassigns _settings, never cache.
    // Satisfied by: MainWindow's `_settings` field.
    AppSettings Settings { get; }

    // Release-note catalog shown by ShowReleaseNotes_Click. Satisfied by: MainWindow's static ReleaseNotes list
    // (kept in the shell because the engine's update flow also references it).
    IReadOnlyList<ReleaseNoteItem> ReleaseNotes { get; }
}
