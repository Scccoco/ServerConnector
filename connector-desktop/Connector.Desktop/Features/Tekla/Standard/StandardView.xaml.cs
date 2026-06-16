using System.Diagnostics;
using System.IO;
using System.Windows;
using Connector.Desktop.Services;
using Forms = System.Windows.Forms;

namespace Connector.Desktop.Features.Tekla.Standard;

// "Стандарт" feature module View. The Tekla firm-folder / extensions / libraries git-sync engine + admin publish +
// server-restart UI, LIFTED VERBATIM out of MainWindow code-behind. Every former shell reference (_settings,
// _settingsService, _heartbeatClient, AppendLog, ThemedDialogs.Show(this,...), _teklaCheckInProgress, _trayIcon,
// the cross-tab/header status controls, RestartManagedServerAsync) now goes through the injected IShellHost seam.
// The logic is intentionally NOT rewritten into idiomatic MVVM — it is the same code with the seam substitutions,
// to keep runtime risk low (the app cannot be run-tested here).
public partial class StandardView : System.Windows.Controls.UserControl
{
    private readonly IShellHost _host;

    // The module owns the Стандарт-only Tekla service (was a MainWindow field: HttpClient with a 25s timeout).
    private readonly TeklaStandardService _teklaStandardService;

    // Tekla-sync error-notice dedup key (was a MainWindow field). Used only by the lifted sync engine, so it lives here.
    private string _lastTeklaSyncErrorNotice = string.Empty;

    // === Стандарт-only constants, moved verbatim from MainWindow ===
    private const string FixedTeklaStandardManifestUrl = "https://server.structura-most.ru/updates/tekla/firm/latest.json";
    private const string FixedTeklaExtensionsManifestUrl = "https://server.structura-most.ru/updates/tekla/extensions/latest.json";
    private const string FixedTeklaLibrariesManifestUrl = "https://server.structura-most.ru/updates/tekla/libraries/latest.json";
    private const string DefaultTeklaStandardLocalPath = @"C:\Company\TeklaFirm";
    private const string DefaultTeklaExtensionsLocalPath = @"C:\TeklaStructures\2025.0\Environments\common\Extensions";
    private static readonly string DefaultTeklaLibrariesLocalPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Grasshopper",
        "Libraries");
    private const string DefaultTeklaPublishSourcePath = @"\\62.113.36.107\BIM_Models\Tekla\02_ПАПКА ФИРМЫ\01_XS_FIRM";
    private const string DefaultTeklaExtensionsPublishSourcePath = @"\\62.113.36.107\BIM_Models\Tekla\02_ПАПКА ФИРМЫ\07_Extensions";
    private const string DefaultTeklaLibrariesPublishSourcePath = @"\\62.113.36.107\BIM_Models\Tekla\02_ПАПКА ФИРМЫ\02_Grasshopper\Libraries\8";

    public StandardView(IShellHost host, TeklaStandardService teklaStandardService)
    {
        _host = host;
        _teklaStandardService = teklaStandardService;
        InitializeComponent();
    }

    // ===== UI refresh (lifted from UpdateTeklaUi) ====================================================
    // Writes ONLY the Стандарт-tab controls that now live in this View. The two cross-tab writes the original made
    // (ConnectorTeklaSyncStatusTextBlock/Button on the Коннектор tab) and the header write (HeaderFirmStatusTextBlock,
    // formerly via UpdateHeaderStatusUi) are pushed to the shell at the end through _host.OnTeklaStatusChanged.
    public void UpdateTeklaUi()
    {
        var firmTargetRevision = string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardTargetRevision)
            ? "-"
            : _host.Settings.TeklaStandardTargetRevision.Trim();
        var firmHasTargetRevision = !string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardTargetRevision);
        TeklaCurrentVersionTextBlock.Text = firmTargetRevision;
        TeklaUpToDateTextBlock.Text = firmHasTargetRevision &&
                                      !_teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaStandardInstalledRevision, _host.Settings.TeklaStandardTargetRevision)
            ? "да"
            : "нет";
        TeklaStatusTextBlock.Text = BuildFirmStatusText();
        TeklaRoleTextBlock.Text = _host.Settings.IsFirmAdmin ? "Роль администратора: да" : "Роль администратора: нет";
        TeklaFirmLocalPathTextBox.Text = string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardLocalPath)
            ? DefaultTeklaStandardLocalPath
            : _host.Settings.TeklaStandardLocalPath;
        TeklaExtensionsCurrentVersionTextBlock.Text = string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsTargetRevision)
            ? "-"
            : _host.Settings.TeklaExtensionsTargetRevision.Trim();
        TeklaExtensionsUpToDateTextBlock.Text = !string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsTargetRevision) &&
                                                !_teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaExtensionsInstalledRevision, _host.Settings.TeklaExtensionsTargetRevision)
            ? "да"
            : "нет";
        TeklaExtensionsStatusTextBlock.Text = BuildExtensionsStatusText();
        TeklaExtensionsLocalPathTextBox.Text = string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsLocalPath)
            ? DefaultTeklaExtensionsLocalPath
            : _host.Settings.TeklaExtensionsLocalPath;
        TeklaLibrariesCurrentVersionTextBlock.Text = string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesTargetRevision)
            ? "-"
            : _host.Settings.TeklaLibrariesTargetRevision.Trim();
        TeklaLibrariesUpToDateTextBlock.Text = !string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesTargetRevision) &&
                                               !_teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaLibrariesInstalledRevision, _host.Settings.TeklaLibrariesTargetRevision)
            ? "да"
            : "нет";
        TeklaLibrariesStatusTextBlock.Text = BuildLibrariesStatusText();
        TeklaLibrariesLocalPathTextBox.Text = string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesLocalPath)
            ? DefaultTeklaLibrariesLocalPath
            : _host.Settings.TeklaLibrariesLocalPath;
        TeklaPublishFirmSourcePathTextBox.Text = string.IsNullOrWhiteSpace(_host.Settings.TeklaPublishSourcePath)
            ? DefaultTeklaPublishSourcePath
            : _host.Settings.TeklaPublishSourcePath;
        TeklaPublishExtensionsSourcePathTextBox.Text = string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsPublishSourcePath)
            ? DefaultTeklaExtensionsPublishSourcePath
            : _host.Settings.TeklaExtensionsPublishSourcePath;
        TeklaPublishLibrariesSourcePathTextBox.Text = string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesPublishSourcePath)
            ? DefaultTeklaLibrariesPublishSourcePath
            : _host.Settings.TeklaLibrariesPublishSourcePath;
        UpdateTeklaActionButtonUi();

        TeklaPublishPanel.Visibility = _host.Settings.IsFirmAdmin ? Visibility.Visible : Visibility.Collapsed;
        TeklaPublishButton.IsEnabled = _host.Settings.IsFirmAdmin;
        if (string.IsNullOrWhiteSpace(TeklaPublishNotesTextBox.Text))
        {
            TeklaPublishNotesTextBox.Text = "Публикация из Structura Connector";
        }

        var canRestartTeklaServer = _host.Settings.IsSystemAdmin || _host.Settings.IsFirmAdmin;
        ServerActionsPanel.Visibility = canRestartTeklaServer ? Visibility.Visible : Visibility.Collapsed;
        RestartTeklaServerButton.IsEnabled = canRestartTeklaServer;
        var (teklaOverallText, teklaOverallBrush) = BuildTeklaOverallStatus();
        TeklaOverallStatusTextBlock.Text = teklaOverallText;
        TeklaOverallStatusTextBlock.Foreground = teklaOverallBrush;

        // Push the computed overall status + action-button presentation to the shell, which mirrors them onto the
        // Коннектор tab (ConnectorTeklaSyncStatusTextBlock/Button) and the window header (HeaderFirmStatusTextBlock).
        _host.OnTeklaStatusChanged(
            teklaOverallText,
            teklaOverallBrush,
            _lastActionButtonContent,
            _lastActionIsSyncStyle,
            _host.TeklaCheckInProgress);
    }

    // The action-button presentation last computed by UpdateTeklaActionButtonUi, so UpdateTeklaUi can mirror it
    // onto the Коннектор tab via the shell. (The original wrote ConnectorTeklaSyncButton directly here.)
    private string _lastActionButtonContent = "Проверить и синхронизировать Tekla";
    private bool _lastActionIsSyncStyle = true;

    // ===== Action button (lifted from UpdateTeklaActionButtonUi) =====================================
    // Sets the local TeklaCheckButton (in this View) and records the presentation so the shell can mirror it onto
    // ConnectorTeklaSyncButton (which is no longer in this View).
    private void UpdateTeklaActionButtonUi()
    {
        if (_host.TeklaCheckInProgress)
        {
            const string inProgressText = "Идет синхронизация Tekla...";
            TeklaCheckButton.Content = inProgressText;
            TeklaCheckButton.Style = (Style)System.Windows.Application.Current.FindResource("SyncButton");
            TeklaCheckButton.IsEnabled = true;
            _lastActionButtonContent = inProgressText;
            _lastActionIsSyncStyle = true;
            return;
        }

        var hasUpdate =
            _teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaStandardInstalledRevision, _host.Settings.TeklaStandardTargetRevision) ||
            _teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaExtensionsInstalledRevision, _host.Settings.TeklaExtensionsTargetRevision) ||
            _teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaLibrariesInstalledRevision, _host.Settings.TeklaLibrariesTargetRevision);
        var hasError = !string.IsNullOrWhiteSpace(FirstNonEmpty(
            _host.Settings.TeklaStandardLastError,
            _host.Settings.TeklaExtensionsLastError,
            _host.Settings.TeklaLibrariesLastError));

        TeklaCheckButton.Content = hasError
            ? "Повторить синхронизацию Tekla"
            : hasUpdate
                ? "Обновить Tekla сейчас"
                : "Проверить и синхронизировать Tekla";
        TeklaCheckButton.Style = (Style)System.Windows.Application.Current.FindResource(hasUpdate || hasError ? "SyncButton" : "SecondaryButton");
        TeklaCheckButton.IsEnabled = true;
        _lastActionButtonContent = (string)TeklaCheckButton.Content;
        _lastActionIsSyncStyle = hasUpdate || hasError;
    }

    private (string Text, System.Windows.Media.Brush Brush) BuildTeklaOverallStatus()
    {
        var hasKnownState = false;
        var outdated = new List<string>();
        var errors = new List<string>();

        void Collect(string name, string installedRevision, string targetRevision, string lastError, bool pendingAfterClose)
        {
            if (!string.IsNullOrWhiteSpace(installedRevision) || !string.IsNullOrWhiteSpace(targetRevision))
            {
                hasKnownState = true;
            }

            if (pendingAfterClose || _teklaStandardService.IsUpdateAvailable(installedRevision, targetRevision))
            {
                outdated.Add(name);
            }

            if (string.IsNullOrWhiteSpace(lastError))
            {
                return;
            }

            if (string.Equals(lastError, "manifest_not_received", StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(name + " (нет связи с сервером обновлений)");
                return;
            }

            errors.Add(name);
        }

        Collect(
            "папка фирмы",
            _host.Settings.TeklaStandardInstalledRevision,
            _host.Settings.TeklaStandardTargetRevision,
            _host.Settings.TeklaStandardLastError,
            _host.Settings.TeklaStandardPendingAfterClose);
        Collect(
            "пользовательские приложения",
            _host.Settings.TeklaExtensionsInstalledRevision,
            _host.Settings.TeklaExtensionsTargetRevision,
            _host.Settings.TeklaExtensionsLastError,
            _host.Settings.TeklaExtensionsPendingAfterClose);
        Collect(
            "Grasshopper Libraries",
            _host.Settings.TeklaLibrariesInstalledRevision,
            _host.Settings.TeklaLibrariesTargetRevision,
            _host.Settings.TeklaLibrariesLastError,
            _host.Settings.TeklaLibrariesPendingAfterClose);

        if (!hasKnownState)
        {
            return ("Tekla Sync: проверка еще не выполнялась", System.Windows.Media.Brushes.DarkGray);
        }

        if (errors.Count > 0)
        {
            return ("Tekla Sync: есть ошибки в разделах: " + string.Join(", ", errors), System.Windows.Media.Brushes.Orange);
        }

        if (outdated.Count > 0)
        {
            return ("Tekla Sync: требуется обновить: " + string.Join(", ", outdated), System.Windows.Media.Brushes.Orange);
        }

        return ("Tekla Sync: все разделы актуальны", System.Windows.Media.Brushes.MediumSpringGreen);
    }

    // Lifted from ShowTeklaSyncFailedBalloon (called by the sync engine). Tray balloon + visibility check now go
    // through the seam; the rest is verbatim.
    private void ShowTeklaSyncFailedBalloon(string targetName, string message)
    {
        var noticeKey = targetName + "|" + message;
        if (string.Equals(_lastTeklaSyncErrorNotice, noticeKey, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _lastTeklaSyncErrorNotice = noticeKey;
        var balloonMessage = string.IsNullOrWhiteSpace(message)
            ? targetName + ": не удалось синхронизировать. Закройте блокирующую программу и повторите синхронизацию."
            : message;
        if (balloonMessage.Length > 240)
        {
            balloonMessage = balloonMessage[..237] + "...";
        }
        _host.ShowTrayBalloon(
            5000,
            "Стандарт Tekla",
            balloonMessage,
            isWarning: true);

        if (_host.IsWindowVisible)
        {
            _host.ShowDialog(
                string.IsNullOrWhiteSpace(message)
                    ? targetName + ": не удалось синхронизировать. Закройте блокирующую программу и повторите синхронизацию."
                    : message,
                "Не удалось обновить: " + targetName,
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    // ===== Sync engine (lifted from RunTeklaSyncCycleAsync) ==========================================
    private async Task<List<string>> RunTeklaSyncCycleAsync(
        bool showDialogs,
        bool forceRefresh,
        bool autoApplyIfPossible,
        OperationProgressWindow? progressReporter = null)
    {
        var summaryLines = new List<string>();
        if (_host.TeklaCheckInProgress)
        {
            summaryLines.Add("Синхронизация уже выполняется");
            return summaryLines;
        }

        _host.TeklaCheckInProgress = true;
        UpdateTeklaActionButtonUi();
        // The original pushed the in-progress button state to the Коннектор mirror via the shared control writes
        // inside UpdateTeklaActionButtonUi; with the mirror now behind the seam, push the in-progress snapshot.
        PushActionButtonOnly();

        try
        {
            NormalizeTeklaSettings();
            ApplyAndPersistTeklaPathsOnly();

            const int totalSteps = 3;
            var currentStep = 0;

            currentStep++;
            progressReporter?.UpdateStep(
                "Проверяем раздел: Папка фирмы",
                "Получаем данные и применяем обновление при необходимости",
                currentStep,
                totalSteps,
                EstimateOperationEta(currentStep, totalSteps));
            await CheckAndApplyTeklaTargetAsync(CreateFirmTargetState(), ApplyFirmTargetState, forceRefresh, autoApplyIfPossible, summaryLines, progressReporter);

            currentStep++;
            progressReporter?.UpdateStep(
                "Проверяем раздел: Пользовательские приложения",
                "Получаем данные и применяем обновление при необходимости",
                currentStep,
                totalSteps,
                EstimateOperationEta(currentStep, totalSteps));
            await CheckAndApplyTeklaTargetAsync(CreateExtensionsTargetState(), ApplyExtensionsTargetState, forceRefresh, autoApplyIfPossible, summaryLines, progressReporter);

            currentStep++;
            progressReporter?.UpdateStep(
                "Проверяем раздел: Grasshopper Libraries",
                "Получаем данные и применяем обновление при необходимости",
                currentStep,
                totalSteps,
                EstimateOperationEta(currentStep, totalSteps));
            await CheckAndApplyTeklaTargetAsync(CreateLibrariesTargetState(), ApplyLibrariesTargetState, forceRefresh, autoApplyIfPossible, summaryLines, progressReporter);

            _host.SaveSettings();
            UpdateTeklaUi();

            if (showDialogs)
            {
                var message = summaryLines.Count == 0
                    ? "Проверка Tekla sync завершена"
                    : string.Join(Environment.NewLine, summaryLines);
                _host.ShowDialog(
                    message,
                    HasTeklaSyncErrors() ? "Стандарт Tekla: есть ошибки" : "Стандарт Tekla",
                    MessageBoxButton.OK,
                    HasTeklaSyncErrors() ? MessageBoxImage.Warning : MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            _host.Settings.TeklaStandardLastError = ex.Message;
            _host.Settings.TeklaExtensionsLastError = ex.Message;
            _host.Settings.TeklaLibrariesLastError = ex.Message;
            _host.Settings.TeklaStandardLastTechnicalError = ex.ToString();
            _host.Settings.TeklaExtensionsLastTechnicalError = ex.ToString();
            _host.Settings.TeklaLibrariesLastTechnicalError = ex.ToString();
            _host.SaveSettings();
            summaryLines.Add("Синхронизация завершилась с ошибкой: " + ex.Message);
            _host.Log("Ошибка проверки Tekla sync: " + ex.Message);
            _host.Log("Технические детали Tekla sync: " + ex);
            if (showDialogs)
            {
                _host.ShowDialog(ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        finally
        {
            _host.TeklaCheckInProgress = false;
            UpdateTeklaUi();
        }

        return summaryLines;
    }

    // Recompute + push only the action-button presentation (used at sync start so the Коннектор mirror flips to the
    // in-progress state immediately, matching the original UpdateTeklaActionButtonUi side effect).
    private void PushActionButtonOnly()
    {
        var (overallText, overallBrush) = BuildTeklaOverallStatus();
        _host.OnTeklaStatusChanged(overallText, overallBrush, _lastActionButtonContent, _lastActionIsSyncStyle, _host.TeklaCheckInProgress);
    }

    private async Task CheckAndApplyTeklaTargetAsync(
        TeklaManagedTargetState target,
        Action<TeklaManagedTargetState> persist,
        bool forceRefresh,
        bool autoApplyIfPossible,
        List<string> summaryLines,
        OperationProgressWindow? progressReporter = null)
    {
        if (forceRefresh)
        {
            target.TargetRevision = string.Empty;
        }

        var manifest = await _teklaStandardService.TryGetManifestAsync(target.ManifestUrl, CancellationToken.None);
        target.LastCheckUtc = DateTimeOffset.UtcNow;

        if (manifest is null)
        {
            target.TargetVersion = string.Empty;
            target.TargetRevision = string.Empty;
            target.LastError = "manifest_not_received";
            target.LastTechnicalError = "Не удалось получить manifest по адресу: " + target.ManifestUrl;
            target.PendingAfterClose = false;
            persist(target);
            summaryLines.Add(target.DisplayName + ": данные обновления недоступны");
            _host.Log(target.DisplayName + ": технические детали: " + target.LastTechnicalError);
            return;
        }

        target.TargetVersion = manifest.Version;
        target.TargetRevision = manifest.Revision;
        target.RepoUrl = manifest.RepoUrl;
        target.RepoRef = manifest.RepoRef;
        target.RepoSubdir = manifest.RepoSubdir;

        var updateAvailable = _teklaStandardService.IsUpdateAvailable(target.InstalledRevision, target.TargetRevision);
        if (!updateAvailable)
        {
            target.PendingAfterClose = false;
            target.LastError = string.Empty;
            target.LastTechnicalError = string.Empty;
            _lastTeklaSyncErrorNotice = string.Empty;
            persist(target);
            if (target.Key == "firm")
            {
                _host.ResetTeklaPendingBalloon();
            }
            summaryLines.Add(target.DisplayName + ": актуально (" + target.TargetRevision + ")");
            return;
        }

        _host.Log(target.DisplayName + ": найдена новая ревизия " + target.TargetRevision);

        if (!autoApplyIfPossible)
        {
            target.PendingAfterClose = false;
            target.LastError = string.Empty;
            target.LastTechnicalError = string.Empty;
            persist(target);
            summaryLines.Add(target.DisplayName + ": доступна ревизия " + target.TargetRevision);
            return;
        }

        progressReporter?.UpdateDetail("Копируем обновления в локальную папку: " + target.DisplayName);
        var applyResult = await Task.Run(() => _teklaStandardService.ApplyPendingGitUpdate(new TeklaManagedSyncRequest
        {
            TargetKey = target.Key,
            DisplayName = target.DisplayName,
            RepoUrl = target.RepoUrl,
            RepoRef = target.RepoRef,
            RepoSubdir = target.RepoSubdir,
            LocalPath = target.LocalPath,
            TargetRevision = target.TargetRevision,
            Mode = target.SyncMode
        }));

        if (applyResult.IsSuccess)
        {
            _teklaStandardService.AppendLog(applyResult.Message);
            _host.Log(applyResult.Message);
            target.InstalledVersion = target.TargetVersion;
            target.InstalledRevision = applyResult.InstalledRevision;
            target.LastSuccessUtc = DateTimeOffset.UtcNow;
            target.PendingAfterClose = false;
            target.LastError = string.Empty;
            target.LastTechnicalError = string.Empty;
            _lastTeklaSyncErrorNotice = string.Empty;
            if (target.Key == "firm")
            {
                _host.ResetTeklaPendingBalloon();
            }
            persist(target);
            summaryLines.Add(target.DisplayName + ": синхронизация выполнена");
            return;
        }

        target.LastError = applyResult.Message;
        target.LastTechnicalError = applyResult.TechnicalDetails;
        target.PendingAfterClose = false;
        persist(target);
        if (!string.IsNullOrWhiteSpace(applyResult.TechnicalDetails))
        {
            _host.Log(applyResult.TechnicalDetails);
        }
        if (autoApplyIfPossible && !string.IsNullOrWhiteSpace(applyResult.Message))
        {
            ShowTeklaSyncFailedBalloon(target.DisplayName, applyResult.Message);
        }
        summaryLines.Add(applyResult.Message);
    }

    private async void TeklaUpdateAction_Click(object sender, RoutedEventArgs e) => await RunInteractiveSyncAsync();

    // Interactive sync with a progress window + "already running" notice. Public so the shell can route the
    // Коннектор-tab mirror button (its Click forwarder) through the SAME path as the Стандарт-tab button —
    // identical progress window + guard, no behavior delta on the mirror.
    public async Task RunInteractiveSyncAsync()
    {
        if (_host.TeklaCheckInProgress)
        {
            _host.ShowDialog(
                "Синхронизация уже выполняется. Дождитесь завершения текущей операции",
                "Стандарт Tekla",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
            return;
        }

        var progressWindow = new OperationProgressWindow("Синхронизация Tekla", "Проверяем и обновляем локальные папки");
        progressWindow.Owner = _host.OwnerWindow;
        progressWindow.Show();
        try
        {
            var summaryLines = await RunTeklaSyncCycleAsync(
                showDialogs: false,
                forceRefresh: true,
                autoApplyIfPossible: true,
                progressReporter: progressWindow);

            var resultText = summaryLines.Count == 0
                ? "Синхронизация завершена"
                : string.Join(Environment.NewLine, summaryLines);
            if (HasTeklaSyncErrors())
            {
                progressWindow.MarkFailed(resultText);
            }
            else
            {
                progressWindow.MarkSucceeded(resultText);
            }
        }
        catch (Exception ex)
        {
            progressWindow.MarkFailed(ex.Message);
        }
    }

    private void TeklaOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenTeklaFolder(
            string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardLocalPath) ? DefaultTeklaStandardLocalPath : _host.Settings.TeklaStandardLocalPath.Trim(),
            "Папка фирмы Tekla");
    }

    private void TeklaFirmBrowse_Click(object sender, RoutedEventArgs e)
    {
        BrowseLocalTeklaFolder(
            "Выберите локальную папку фирмы Tekla",
            TeklaFirmLocalPathTextBox,
            path => _host.Settings.TeklaStandardLocalPath = path,
            "папки фирмы");
    }

    private void TeklaExtensionsOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenTeklaFolder(
            string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsLocalPath) ? DefaultTeklaExtensionsLocalPath : _host.Settings.TeklaExtensionsLocalPath.Trim(),
            "Extensions Tekla");
    }

    private void TeklaExtensionsBrowse_Click(object sender, RoutedEventArgs e)
    {
        BrowseLocalTeklaFolder(
            "Выберите локальную папку Extensions для Tekla",
            TeklaExtensionsLocalPathTextBox,
            path => _host.Settings.TeklaExtensionsLocalPath = path,
            "Extensions");
    }

    private void TeklaLibrariesOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        OpenTeklaFolder(
            string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesLocalPath) ? DefaultTeklaLibrariesLocalPath : _host.Settings.TeklaLibrariesLocalPath.Trim(),
            "Grasshopper Libraries");
    }

    private void TeklaLibrariesBrowse_Click(object sender, RoutedEventArgs e)
    {
        BrowseLocalTeklaFolder(
            "Выберите локальную папку Grasshopper Libraries",
            TeklaLibrariesLocalPathTextBox,
            path => _host.Settings.TeklaLibrariesLocalPath = path,
            "Libraries");
    }

    private void BrowseLocalTeklaFolder(
        string description,
        System.Windows.Controls.TextBox targetTextBox,
        Action<string> applyPath,
        string label)
    {
        try
        {
            using var dialog = new Forms.FolderBrowserDialog
            {
                Description = description,
                ShowNewFolderButton = true,
                SelectedPath = (targetTextBox.Text ?? string.Empty).Trim()
            };

            if (dialog.ShowDialog() == Forms.DialogResult.OK)
            {
                targetTextBox.Text = dialog.SelectedPath;
                applyPath(dialog.SelectedPath);
                _host.SaveSettings();
                UpdateTeklaUi();
            }
        }
        catch (Exception ex)
        {
            _host.Log("Ошибка выбора папки " + label + ": " + ex.Message);
            _host.ShowDialog(ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenTeklaFolder(string folderPath, string label)
    {
        try
        {
            Directory.CreateDirectory(folderPath);
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = folderPath,
                UseShellExecute = true
            });
            _host.Log("Открыта папка " + label + ": " + folderPath);
        }
        catch (Exception ex)
        {
            _host.Log("Ошибка открытия папки " + label + ": " + ex.Message);
            _host.ShowDialog(ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void NormalizeTeklaSettings()
    {
        _host.Settings.TeklaStandardManifestUrl = string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardManifestUrl)
            ? FixedTeklaStandardManifestUrl
            : _host.Settings.TeklaStandardManifestUrl;
        _host.Settings.TeklaStandardLocalPath = string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardLocalPath)
            ? DefaultTeklaStandardLocalPath
            : _host.Settings.TeklaStandardLocalPath;
        _host.Settings.TeklaExtensionsManifestUrl = string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsManifestUrl)
            ? FixedTeklaExtensionsManifestUrl
            : _host.Settings.TeklaExtensionsManifestUrl;
        _host.Settings.TeklaExtensionsLocalPath = string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsLocalPath)
            ? DefaultTeklaExtensionsLocalPath
            : _host.Settings.TeklaExtensionsLocalPath;
        _host.Settings.TeklaLibrariesManifestUrl = string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesManifestUrl)
            ? FixedTeklaLibrariesManifestUrl
            : _host.Settings.TeklaLibrariesManifestUrl;
        _host.Settings.TeklaLibrariesLocalPath = string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesLocalPath)
            ? DefaultTeklaLibrariesLocalPath
            : _host.Settings.TeklaLibrariesLocalPath;
        _host.Settings.TeklaPublishSourcePath =
            string.IsNullOrWhiteSpace(_host.Settings.TeklaPublishSourcePath) ||
            string.Equals(_host.Settings.TeklaPublishSourcePath, @"\\62.113.36.107\BIM_Models\Tekla\XS_FIRM", StringComparison.OrdinalIgnoreCase)
            ? DefaultTeklaPublishSourcePath
            : _host.Settings.TeklaPublishSourcePath;
        _host.Settings.TeklaExtensionsPublishSourcePath =
            string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsPublishSourcePath) ||
            string.Equals(_host.Settings.TeklaExtensionsPublishSourcePath, @"\\62.113.36.107\BIM_Models\Tekla\Extension", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(_host.Settings.TeklaExtensionsPublishSourcePath, @"\\62.113.36.107\BIM_Models\Tekla\Extensions", StringComparison.OrdinalIgnoreCase)
            ? DefaultTeklaExtensionsPublishSourcePath
            : _host.Settings.TeklaExtensionsPublishSourcePath;
        _host.Settings.TeklaLibrariesPublishSourcePath = string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesPublishSourcePath)
            ? DefaultTeklaLibrariesPublishSourcePath
            : _host.Settings.TeklaLibrariesPublishSourcePath;
    }

    private void ApplyAndPersistTeklaPathsOnly()
    {
        var firmPath = (TeklaFirmLocalPathTextBox.Text ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(firmPath))
        {
            _host.Settings.TeklaStandardLocalPath = firmPath;
        }

        var extensionsPath = (TeklaExtensionsLocalPathTextBox.Text ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(extensionsPath))
        {
            _host.Settings.TeklaExtensionsLocalPath = extensionsPath;
        }

        var librariesPath = (TeklaLibrariesLocalPathTextBox.Text ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(librariesPath))
        {
            _host.Settings.TeklaLibrariesLocalPath = librariesPath;
        }

        var firmPublishSourcePath = (TeklaPublishFirmSourcePathTextBox.Text ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(firmPublishSourcePath))
        {
            _host.Settings.TeklaPublishSourcePath = firmPublishSourcePath;
        }

        var extensionsPublishSourcePath = (TeklaPublishExtensionsSourcePathTextBox.Text ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(extensionsPublishSourcePath))
        {
            _host.Settings.TeklaExtensionsPublishSourcePath = extensionsPublishSourcePath;
        }

        var librariesPublishSourcePath = (TeklaPublishLibrariesSourcePathTextBox.Text ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(librariesPublishSourcePath))
        {
            _host.Settings.TeklaLibrariesPublishSourcePath = librariesPublishSourcePath;
        }
    }

    // Flush the editable path TextBoxes into _host.Settings. The shell calls this before ReadSettingsFromUi
    // (Save/Start/SendNow) so a path TYPED on the Стандарт tab — without an intervening Browse or sync — is
    // captured, restoring the pre-migration "ReadSettingsFromUi read the TextBoxes directly" semantics.
    public void FlushPathEdits() => ApplyAndPersistTeklaPathsOnly();

    private TeklaManagedTargetState CreateFirmTargetState()
    {
        return new TeklaManagedTargetState
        {
            Key = "firm",
            DisplayName = "Папка фирмы Tekla",
            ManifestUrl = _host.Settings.TeklaStandardManifestUrl,
            LocalPath = _host.Settings.TeklaStandardLocalPath,
            InstalledVersion = _host.Settings.TeklaStandardInstalledVersion,
            TargetVersion = _host.Settings.TeklaStandardTargetVersion,
            InstalledRevision = _host.Settings.TeklaStandardInstalledRevision,
            TargetRevision = _host.Settings.TeklaStandardTargetRevision,
            LastCheckUtc = _host.Settings.TeklaStandardLastCheckUtc,
            LastSuccessUtc = _host.Settings.TeklaStandardLastSuccessUtc,
            PendingAfterClose = _host.Settings.TeklaStandardPendingAfterClose,
            LastError = _host.Settings.TeklaStandardLastError,
            LastTechnicalError = _host.Settings.TeklaStandardLastTechnicalError,
            RepoUrl = _host.Settings.TeklaStandardRepoUrl,
            RepoRef = _host.Settings.TeklaStandardRepoRef,
            RepoSubdir = _host.Settings.TeklaStandardRepoSubdir,
            SyncMode = TeklaManagedSyncMode.Strict,
            DelayWhenTeklaRunning = true
        };
    }

    private void ApplyFirmTargetState(TeklaManagedTargetState target)
    {
        _host.Settings.TeklaStandardManifestUrl = target.ManifestUrl;
        _host.Settings.TeklaStandardLocalPath = target.LocalPath;
        _host.Settings.TeklaStandardInstalledVersion = target.InstalledVersion;
        _host.Settings.TeklaStandardTargetVersion = target.TargetVersion;
        _host.Settings.TeklaStandardInstalledRevision = target.InstalledRevision;
        _host.Settings.TeklaStandardTargetRevision = target.TargetRevision;
        _host.Settings.TeklaStandardLastCheckUtc = target.LastCheckUtc;
        _host.Settings.TeklaStandardLastSuccessUtc = target.LastSuccessUtc;
        _host.Settings.TeklaStandardPendingAfterClose = target.PendingAfterClose;
        _host.Settings.TeklaStandardLastError = target.LastError;
        _host.Settings.TeklaStandardLastTechnicalError = target.LastTechnicalError;
        _host.Settings.TeklaStandardRepoUrl = target.RepoUrl;
        _host.Settings.TeklaStandardRepoRef = target.RepoRef;
        _host.Settings.TeklaStandardRepoSubdir = target.RepoSubdir;
        TeklaStatusTextBlock.Text = BuildFirmStatusText();
    }

    private TeklaManagedTargetState CreateExtensionsTargetState()
    {
        return new TeklaManagedTargetState
        {
            Key = "extensions",
            DisplayName = "Extensions Tekla",
            ManifestUrl = _host.Settings.TeklaExtensionsManifestUrl,
            LocalPath = _host.Settings.TeklaExtensionsLocalPath,
            InstalledVersion = _host.Settings.TeklaExtensionsInstalledVersion,
            TargetVersion = _host.Settings.TeklaExtensionsTargetVersion,
            InstalledRevision = _host.Settings.TeklaExtensionsInstalledRevision,
            TargetRevision = _host.Settings.TeklaExtensionsTargetRevision,
            LastCheckUtc = _host.Settings.TeklaExtensionsLastCheckUtc,
            LastSuccessUtc = _host.Settings.TeklaExtensionsLastSuccessUtc,
            PendingAfterClose = _host.Settings.TeklaExtensionsPendingAfterClose,
            LastError = _host.Settings.TeklaExtensionsLastError,
            LastTechnicalError = _host.Settings.TeklaExtensionsLastTechnicalError,
            RepoUrl = _host.Settings.TeklaExtensionsRepoUrl,
            RepoRef = _host.Settings.TeklaExtensionsRepoRef,
            RepoSubdir = _host.Settings.TeklaExtensionsRepoSubdir,
            SyncMode = TeklaManagedSyncMode.OverlayNoDelete,
            DelayWhenTeklaRunning = true
        };
    }

    private void ApplyExtensionsTargetState(TeklaManagedTargetState target)
    {
        _host.Settings.TeklaExtensionsManifestUrl = target.ManifestUrl;
        _host.Settings.TeklaExtensionsLocalPath = target.LocalPath;
        _host.Settings.TeklaExtensionsInstalledVersion = target.InstalledVersion;
        _host.Settings.TeklaExtensionsTargetVersion = target.TargetVersion;
        _host.Settings.TeklaExtensionsInstalledRevision = target.InstalledRevision;
        _host.Settings.TeklaExtensionsTargetRevision = target.TargetRevision;
        _host.Settings.TeklaExtensionsLastCheckUtc = target.LastCheckUtc;
        _host.Settings.TeklaExtensionsLastSuccessUtc = target.LastSuccessUtc;
        _host.Settings.TeklaExtensionsPendingAfterClose = target.PendingAfterClose;
        _host.Settings.TeklaExtensionsLastError = target.LastError;
        _host.Settings.TeklaExtensionsLastTechnicalError = target.LastTechnicalError;
        _host.Settings.TeklaExtensionsRepoUrl = target.RepoUrl;
        _host.Settings.TeklaExtensionsRepoRef = target.RepoRef;
        _host.Settings.TeklaExtensionsRepoSubdir = target.RepoSubdir;
        TeklaExtensionsStatusTextBlock.Text = BuildExtensionsStatusText();
    }

    private TeklaManagedTargetState CreateLibrariesTargetState()
    {
        return new TeklaManagedTargetState
        {
            Key = "libraries",
            DisplayName = "Grasshopper Libraries",
            ManifestUrl = _host.Settings.TeklaLibrariesManifestUrl,
            LocalPath = _host.Settings.TeklaLibrariesLocalPath,
            InstalledVersion = _host.Settings.TeklaLibrariesInstalledVersion,
            TargetVersion = _host.Settings.TeklaLibrariesTargetVersion,
            InstalledRevision = _host.Settings.TeklaLibrariesInstalledRevision,
            TargetRevision = _host.Settings.TeklaLibrariesTargetRevision,
            LastCheckUtc = _host.Settings.TeklaLibrariesLastCheckUtc,
            LastSuccessUtc = _host.Settings.TeklaLibrariesLastSuccessUtc,
            PendingAfterClose = _host.Settings.TeklaLibrariesPendingAfterClose,
            LastError = _host.Settings.TeklaLibrariesLastError,
            LastTechnicalError = _host.Settings.TeklaLibrariesLastTechnicalError,
            RepoUrl = _host.Settings.TeklaLibrariesRepoUrl,
            RepoRef = _host.Settings.TeklaLibrariesRepoRef,
            RepoSubdir = _host.Settings.TeklaLibrariesRepoSubdir,
            SyncMode = TeklaManagedSyncMode.OverlayNoDelete,
            DelayWhenTeklaRunning = false
        };
    }

    private void ApplyLibrariesTargetState(TeklaManagedTargetState target)
    {
        _host.Settings.TeklaLibrariesManifestUrl = target.ManifestUrl;
        _host.Settings.TeklaLibrariesLocalPath = target.LocalPath;
        _host.Settings.TeklaLibrariesInstalledVersion = target.InstalledVersion;
        _host.Settings.TeklaLibrariesTargetVersion = target.TargetVersion;
        _host.Settings.TeklaLibrariesInstalledRevision = target.InstalledRevision;
        _host.Settings.TeklaLibrariesTargetRevision = target.TargetRevision;
        _host.Settings.TeklaLibrariesLastCheckUtc = target.LastCheckUtc;
        _host.Settings.TeklaLibrariesLastSuccessUtc = target.LastSuccessUtc;
        _host.Settings.TeklaLibrariesPendingAfterClose = target.PendingAfterClose;
        _host.Settings.TeklaLibrariesLastError = target.LastError;
        _host.Settings.TeklaLibrariesLastTechnicalError = target.LastTechnicalError;
        _host.Settings.TeklaLibrariesRepoUrl = target.RepoUrl;
        _host.Settings.TeklaLibrariesRepoRef = target.RepoRef;
        _host.Settings.TeklaLibrariesRepoSubdir = target.RepoSubdir;
        TeklaLibrariesStatusTextBlock.Text = BuildLibrariesStatusText();
    }

    private string BuildFirmStatusText()
    {
        if (_host.Settings.TeklaStandardPendingAfterClose)
        {
            return "Папка фирмы: обновление подготовлено и будет применено после закрытия Tekla";
        }

        if (string.Equals(_host.Settings.TeklaStandardLastError, "manifest_not_received", StringComparison.OrdinalIgnoreCase))
        {
            return "Папка фирмы: не удалось получить данные обновления с сервера";
        }

        if (!string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardLastError))
        {
            return _host.Settings.TeklaStandardLastError;
        }

        if (_teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaStandardInstalledRevision, _host.Settings.TeklaStandardTargetRevision))
        {
            var currentRevision = string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardInstalledRevision)
                ? "не установлена"
                : _host.Settings.TeklaStandardInstalledRevision;
            return "Папка фирмы: требуется обновление до ревизии " + _host.Settings.TeklaStandardTargetRevision + " (сейчас " + currentRevision + ")";
        }

        if (!string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardInstalledRevision))
        {
            return "Папка фирмы: актуально (" + _host.Settings.TeklaStandardInstalledRevision + ")";
        }

        return "Папка фирмы: проверка еще не выполнялась";
    }

    private string BuildExtensionsStatusText()
    {
        if (_host.Settings.TeklaExtensionsPendingAfterClose)
        {
            return "Пользовательские приложения: обновление подготовлено и будет применено после закрытия Tekla";
        }

        if (string.Equals(_host.Settings.TeklaExtensionsLastError, "manifest_not_received", StringComparison.OrdinalIgnoreCase))
        {
            return "Пользовательские приложения: не удалось получить данные обновления с сервера";
        }

        if (!string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsLastError))
        {
            return _host.Settings.TeklaExtensionsLastError;
        }

        if (_teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaExtensionsInstalledRevision, _host.Settings.TeklaExtensionsTargetRevision))
        {
            var currentRevision = string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsInstalledRevision)
                ? "не установлена"
                : _host.Settings.TeklaExtensionsInstalledRevision;
            return "Пользовательские приложения: требуется обновление до ревизии " + _host.Settings.TeklaExtensionsTargetRevision + " (сейчас " + currentRevision + ")";
        }

        if (!string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsInstalledRevision))
        {
            return "Пользовательские приложения: актуально (" + _host.Settings.TeklaExtensionsInstalledRevision + ")";
        }

        return "Пользовательские приложения: проверка еще не выполнялась";
    }

    private string BuildLibrariesStatusText()
    {
        if (_host.Settings.TeklaLibrariesPendingAfterClose)
        {
            return "Grasshopper Libraries: обновление подготовлено и будет применено автоматически";
        }

        if (string.Equals(_host.Settings.TeklaLibrariesLastError, "manifest_not_received", StringComparison.OrdinalIgnoreCase))
        {
            return "Grasshopper Libraries: не удалось получить данные обновления с сервера";
        }

        if (!string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesLastError))
        {
            return _host.Settings.TeklaLibrariesLastError;
        }

        if (_teklaStandardService.IsUpdateAvailable(_host.Settings.TeklaLibrariesInstalledRevision, _host.Settings.TeklaLibrariesTargetRevision))
        {
            var currentRevision = string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesInstalledRevision)
                ? "не установлена"
                : _host.Settings.TeklaLibrariesInstalledRevision;
            return "Grasshopper Libraries: требуется обновление до ревизии " + _host.Settings.TeklaLibrariesTargetRevision + " (сейчас " + currentRevision + ")";
        }

        if (!string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesInstalledRevision))
        {
            return "Grasshopper Libraries: актуально (" + _host.Settings.TeklaLibrariesInstalledRevision + ")";
        }

        return "Grasshopper Libraries: проверка еще не выполнялась";
    }

    // Local private copy (the original was a shared generic helper on MainWindow; duplicated here to avoid a seam
    // dependency, per the brief).
    private static string FirstNonEmpty(params string[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return string.Empty;
    }

    private bool HasTeklaSyncErrors()
    {
        return !string.IsNullOrWhiteSpace(_host.Settings.TeklaStandardLastError) ||
               !string.IsNullOrWhiteSpace(_host.Settings.TeklaExtensionsLastError) ||
               !string.IsNullOrWhiteSpace(_host.Settings.TeklaLibrariesLastError);
    }

    private List<TeklaPublishTargetSelection> GetSelectedPublishTargets()
    {
        var targets = new List<TeklaPublishTargetSelection>();

        if (PublishFirmCheckBox.IsChecked == true)
        {
            targets.Add(new TeklaPublishTargetSelection
            {
                TargetKey = "firm",
                DisplayName = "Папка фирмы",
                SourcePath = (TeklaPublishFirmSourcePathTextBox.Text ?? string.Empty).Trim()
            });
        }

        if (PublishExtensionsCheckBox.IsChecked == true)
        {
            targets.Add(new TeklaPublishTargetSelection
            {
                TargetKey = "extensions",
                DisplayName = "Пользовательские приложения",
                SourcePath = (TeklaPublishExtensionsSourcePathTextBox.Text ?? string.Empty).Trim()
            });
        }

        if (PublishLibrariesCheckBox.IsChecked == true)
        {
            targets.Add(new TeklaPublishTargetSelection
            {
                TargetKey = "libraries",
                DisplayName = "Grasshopper Libraries",
                SourcePath = (TeklaPublishLibrariesSourcePathTextBox.Text ?? string.Empty).Trim()
            });
        }

        return targets;
    }

    private void TeklaPublishFirmBrowse_Click(object sender, RoutedEventArgs e)
    {
        BrowsePublishSourceFolder("Папка фирмы", TeklaPublishFirmSourcePathTextBox, DefaultTeklaPublishSourcePath, path => _host.Settings.TeklaPublishSourcePath = path);
    }

    private void TeklaPublishExtensionsBrowse_Click(object sender, RoutedEventArgs e)
    {
        BrowsePublishSourceFolder("Пользовательские приложения", TeklaPublishExtensionsSourcePathTextBox, DefaultTeklaExtensionsPublishSourcePath, path => _host.Settings.TeklaExtensionsPublishSourcePath = path);
    }

    private void TeklaPublishLibrariesBrowse_Click(object sender, RoutedEventArgs e)
    {
        BrowsePublishSourceFolder("Grasshopper Libraries", TeklaPublishLibrariesSourcePathTextBox, DefaultTeklaLibrariesPublishSourcePath, path => _host.Settings.TeklaLibrariesPublishSourcePath = path);
    }

    private void BrowsePublishSourceFolder(
        string label,
        System.Windows.Controls.TextBox targetTextBox,
        string fallbackPath,
        Action<string> saveToSettings)
    {
        try
        {
            using var dialog = new Forms.FolderBrowserDialog
            {
                Description = "Выберите эталонную серверную папку для раздела: " + label,
                ShowNewFolderButton = false,
                SelectedPath = string.IsNullOrWhiteSpace(targetTextBox.Text)
                    ? fallbackPath
                    : targetTextBox.Text.Trim()
            };

            if (dialog.ShowDialog() != Forms.DialogResult.OK)
            {
                return;
            }

            targetTextBox.Text = dialog.SelectedPath;
            saveToSettings(dialog.SelectedPath);
            _host.SaveSettings();
        }
        catch (Exception ex)
        {
            _host.Log("Ошибка выбора эталонной папки для раздела " + label + ": " + ex.Message);
            _host.ShowDialog(ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void TeklaOpenLog_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!File.Exists(_teklaStandardService.LogFilePath))
            {
                File.WriteAllText(_teklaStandardService.LogFilePath, string.Empty);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = "/select,\"" + _teklaStandardService.LogFilePath + "\"",
                UseShellExecute = true
            });
            _host.Log("Открыт лог Стандарт Tekla: " + _teklaStandardService.LogFilePath);
        }
        catch (Exception ex)
        {
            _host.Log("Ошибка открытия лога Стандарт Tekla: " + ex.Message);
            _host.ShowDialog(ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void TeklaPublish_Click(object sender, RoutedEventArgs e)
    {
        OperationProgressWindow? progressWindow = null;
        try
        {
            if (!_host.Settings.IsFirmAdmin)
            {
                throw new InvalidOperationException("Публикация доступна только для роли admin_firm.");
            }

            var token = SettingsService.DecryptToken(_host.Settings.TokenCipherBase64).Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidOperationException("Токен устройства не найден. Выполните подключение по токену.");
            }

            var selectedTargets = GetSelectedPublishTargets();
            if (selectedTargets.Count == 0)
            {
                throw new InvalidOperationException("Выберите хотя бы один раздел для публикации.");
            }

            var publishComment = (TeklaPublishNotesTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(publishComment))
            {
                throw new InvalidOperationException("Комментарий публикации обязателен.");
            }

            foreach (var selectedTarget in selectedTargets)
            {
                if (string.IsNullOrWhiteSpace(selectedTarget.SourcePath))
                {
                    throw new InvalidOperationException("Для раздела \"" + selectedTarget.DisplayName + "\" не указан путь к эталонной папке.");
                }
            }

            _host.Settings.TeklaPublishSourcePath = (TeklaPublishFirmSourcePathTextBox.Text ?? string.Empty).Trim();
            _host.Settings.TeklaExtensionsPublishSourcePath = (TeklaPublishExtensionsSourcePathTextBox.Text ?? string.Empty).Trim();
            _host.Settings.TeklaLibrariesPublishSourcePath = (TeklaPublishLibrariesSourcePathTextBox.Text ?? string.Empty).Trim();
            _host.SaveSettings();

            var totalSteps = selectedTargets.Count + 3;
            var currentStep = 0;
            progressWindow = new OperationProgressWindow("Публикация Tekla", "Подготавливаем публикацию");
            progressWindow.Owner = _host.OwnerWindow;
            progressWindow.Show();

            TeklaPublishButton.IsEnabled = false;

            var resultLines = new List<string>();
            foreach (var selectedTarget in selectedTargets)
            {
                currentStep++;
                progressWindow.UpdateStep(
                    "Публикуем раздел: " + selectedTarget.DisplayName,
                    "Идет проверка изменений и подготовка публикации",
                    currentStep,
                    totalSteps,
                    EstimateOperationEta(currentStep, totalSteps));

                var payload = new TeklaManifestPublishPayload
                {
                    Target = selectedTarget.TargetKey,
                    SourcePath = selectedTarget.SourcePath,
                    Comment = publishComment
                };
                var result = await _host.Heartbeat.PublishTeklaManifestAsync(_host.Settings.ServerUrl, token, payload, CancellationToken.None);

                if (result.NoChanges)
                {
                    resultLines.Add(selectedTarget.DisplayName + ": изменений не найдено");
                    _host.Log("Публикация " + selectedTarget.TargetKey + ": изменений не обнаружено.");
                }
                else
                {
                    resultLines.Add(selectedTarget.DisplayName + ": опубликовано, ревизия " + result.Revision);
                    _host.Log("Публикация " + selectedTarget.TargetKey + ": опубликовано, ревизия " + result.Revision);
                }

                if (!string.IsNullOrWhiteSpace(result.Message))
                {
                    _host.Log(result.Message);
                }
            }

            currentStep++;
            progressWindow.UpdateStep(
                "Обновляем данные по состоянию",
                "Сохраняем информацию о публикации",
                currentStep,
                totalSteps,
                EstimateOperationEta(currentStep, totalSteps));
            _host.SaveSettings();

            currentStep++;
            progressWindow.UpdateStep(
                "Запускаем синхронизацию на этом компьютере",
                "Сейчас локальные папки будут приведены к опубликованному состоянию",
                currentStep,
                totalSteps,
                EstimateOperationEta(currentStep, totalSteps));
            await RunTeklaSyncCycleAsync(
                showDialogs: false,
                forceRefresh: true,
                autoApplyIfPossible: true);

            currentStep++;
            progressWindow.UpdateStep(
                "Готово",
                "Публикация и синхронизация завершены",
                currentStep,
                totalSteps,
                TimeSpan.Zero);
            progressWindow.MarkSucceeded(string.Join(Environment.NewLine, resultLines));

            _host.ShowDialog(
                "Публикация завершена\n\n" + string.Join(Environment.NewLine, resultLines),
                "Стандарт Tekla",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            var message = GetFriendlyTeklaPublishErrorMessage(ex);
            progressWindow?.MarkFailed(message);
            _host.Log("Ошибка публикации Tekla: " + message);
            _host.ShowDialog(message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TeklaPublishButton.IsEnabled = _host.Settings.IsFirmAdmin;
        }
    }

    private static TimeSpan EstimateOperationEta(int currentStep, int totalSteps)
    {
        var remainingSteps = Math.Max(0, totalSteps - currentStep);
        return TimeSpan.FromSeconds(Math.Max(0, remainingSteps * 12));
    }

    private void TeklaPublishBrowse_Click(object sender, RoutedEventArgs e)
    {
        // Legacy hidden control left for backward XAML compatibility.
        TeklaPublishFirmBrowse_Click(sender, e);
    }

    private void TeklaPublishTarget_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        // Legacy hidden control left for backward XAML compatibility.
    }

    private static string GetFriendlyTeklaPublishErrorMessage(Exception ex)
    {
        var message = ex.Message?.Trim() ?? "Неизвестная ошибка публикации.";

        if (message.StartsWith("HTTP 409:", StringComparison.OrdinalIgnoreCase))
        {
            return "На сервере уже выполняется публикация. Дождитесь завершения текущей попытки и повторите запуск позже";
        }

        if (message.StartsWith("HTTP 504:", StringComparison.OrdinalIgnoreCase))
        {
            return message[9..].Trim();
        }

        if (message.StartsWith("HTTP 400:", StringComparison.OrdinalIgnoreCase) ||
            message.StartsWith("HTTP 500:", StringComparison.OrdinalIgnoreCase))
        {
            return message[9..].Trim();
        }

        return message;
    }

    private async void RestartTeklaServer_Click(object sender, RoutedEventArgs e)
    {
        await _host.RestartManagedServerAsync("tekla", RestartTeklaServerButton, "Tekla Server");
    }

    // ===== Public surface the shell drives (mirrors the original call sites) =========================
    // The shell calls these where it formerly called UpdateTeklaUi() / RunTeklaSyncCycleAsync(...).
    public void RefreshUi() => UpdateTeklaUi();

    public Task RunTeklaSyncAsync(bool showDialogs, bool forceRefresh, bool autoApplyIfPossible) =>
        RunTeklaSyncCycleAsync(showDialogs, forceRefresh, autoApplyIfPossible);

    // Lifted target-state DTO (was a private nested class on MainWindow).
    private sealed class TeklaManagedTargetState
    {
        public string Key { get; init; } = "";
        public string DisplayName { get; init; } = "";
        public string ManifestUrl { get; set; } = "";
        public string LocalPath { get; set; } = "";
        public string InstalledVersion { get; set; } = "";
        public string TargetVersion { get; set; } = "";
        public string InstalledRevision { get; set; } = "";
        public string TargetRevision { get; set; } = "";
        public DateTimeOffset? LastCheckUtc { get; set; }
        public DateTimeOffset? LastSuccessUtc { get; set; }
        public bool PendingAfterClose { get; set; }
        public string LastError { get; set; } = "";
        public string LastTechnicalError { get; set; } = "";
        public string RepoUrl { get; set; } = "";
        public string RepoRef { get; set; } = "";
        public string RepoSubdir { get; set; } = "";
        public TeklaManagedSyncMode SyncMode { get; init; }
        public bool DelayWhenTeklaRunning { get; init; }
    }

    // Lifted publish-selection DTO (was a private nested class on MainWindow).
    private sealed class TeklaPublishTargetSelection
    {
        public string TargetKey { get; init; } = "";
        public string DisplayName { get; init; } = "";
        public string SourcePath { get; init; } = "";
    }
}
