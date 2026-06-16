using System.Windows;
using Connector.Desktop.Models;
using Connector.Desktop.Services;

namespace Connector.Desktop.Features.Tekla.Standard;

// The seam between the lifted "Стандарт" feature (StandardView) and the shell (MainWindow). The Стандарт logic
// was lifted VERBATIM out of MainWindow; every reference it used to make to a shell field/method now goes through
// this interface. MainWindow implements it (the orchestrator wires it), so the lifted code stays behaviour-identical
// while the View is self-contained. Keep this surface MINIMAL — it is exactly what the lifted code touches.
public interface IShellHost
{
    // The shared, mutable settings model. MainWindow reassigns its _settings field (e.g. ReadSettingsFromUi()),
    // so this MUST be a live getter that returns the CURRENT instance every access — never cache it in the View.
    // Satisfied by: MainWindow's `_settings` field (return it from the getter).
    AppSettings Settings { get; }

    // Persist the current settings. Satisfied by: MainWindow `_settingsService.Save(_settings)`.
    // (SettingsService.DecryptToken is static — the lifted code calls it directly, it is not on this seam.)
    void SaveSettings();

    // The publish + managed-server-restart transport. The lifted TeklaPublish_Click uses Heartbeat.PublishTeklaManifestAsync.
    // Satisfied by: MainWindow's `_heartbeatClient` field.
    HeartbeatClient Heartbeat { get; }

    // Journal line. Satisfied by: MainWindow `AppendLog(string)`.
    void Log(string message);

    // Window-owned themed dialog (replaces ThemedDialogs.Show(this, ...)).
    // Satisfied by: MainWindow `ThemedDialogs.Show(this, message, title, buttons, image)`.
    MessageBoxResult ShowDialog(string message, string title, MessageBoxButton buttons, MessageBoxImage image);

    // The owner window for OperationProgressWindow (so progress dialogs are modal-parented to the shell).
    // The View constructs `new OperationProgressWindow(...) { Owner = host.OwnerWindow }`.
    // Satisfied by: `this` (the MainWindow instance).
    Window OwnerWindow { get; }

    // The sync gate, READ by the shell's connect flow (ConnectByTokenInternalAsync) and other shell call sites,
    // and read/written by the lifted sync engine. It must remain shell-visible, so it is BACKED in the shell.
    // Satisfied by: MainWindow's `_teklaCheckInProgress` field (expose as a get/set property).
    bool TeklaCheckInProgress { get; set; }

    // Shared with the Revit-server restart, so it STAYS in the shell. RestartTeklaServer_Click forwards here.
    // Satisfied by: MainWindow's existing `RestartManagedServerAsync(serviceKey, button, displayName)`.
    Task RestartManagedServerAsync(string serviceKey, System.Windows.Controls.Button button, string displayName);

    // CROSS-TAB MIRROR + HEADER. The original UpdateTeklaUi also wrote ConnectorTeklaSyncStatusTextBlock/Button
    // (on the 'Коннектор' tab) and UpdateHeaderStatusUi wrote HeaderFirmStatusTextBlock (window header) — none of
    // those controls live in this View. The View computes the overall status (it owns BuildTeklaOverallStatus +
    // the action-button logic) and pushes it out here at the end of every refresh; the shell applies it.
    // Satisfied by: MainWindow sets ConnectorTeklaSyncStatusTextBlock + ConnectorTeklaSyncButton (Коннектор mirror)
    // and HeaderFirmStatusTextBlock (header) from these values, mirroring the old UpdateTeklaUi/UpdateHeaderStatusUi.
    void OnTeklaStatusChanged(string overallText, System.Windows.Media.Brush overallBrush, string actionButtonContent, bool actionIsSyncStyle, bool inProgress);

    // Tray balloon used by the lifted ShowTeklaSyncFailedBalloon. Satisfied by: MainWindow
    // `_trayIcon.ShowBalloonTip(durationMs, title, message, isWarning ? Forms.ToolTipIcon.Warning : Forms.ToolTipIcon.Info)`.
    void ShowTrayBalloon(int durationMs, string title, string message, bool isWarning);

    // Is the shell window currently visible & not minimised? Lifted ShowTeklaSyncFailedBalloon gates a dialog on it.
    // Satisfied by: MainWindow `IsVisible && WindowState != WindowState.Minimized`.
    bool IsWindowVisible { get; }

    // Reset the "pending update" tray-balloon-once flag. The shell owns `_teklaBalloonShown` because the residual
    // ShowTeklaPendingBalloon (NOT lifted) sets it; the lifted sync engine resets it on success/up-to-date for "firm".
    // Satisfied by: MainWindow `_teklaBalloonShown = false`.
    void ResetTeklaPendingBalloon();
}
