using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Threading;
using Connector.Desktop.Models;
using Connector.Desktop.Services;
using Forms = System.Windows.Forms;

namespace Connector.Desktop;

public partial class MainWindow : Window
{
    private const string FixedServerUrl = "https://server.structura-most.ru";
    private const string FixedUpdateManifestUrl = "https://server.structura-most.ru/updates/latest.json";
    private const int FixedHeartbeatSeconds = 60;
    private const string DefaultSmbSharePath = @"\\62.113.36.107\BIM_Models";
    private const string FixedTeklaStandardManifestUrl = "https://server.structura-most.ru/updates/tekla/firm/latest.json";
    private const string DefaultTeklaStandardLocalPath = @"C:\Company\TeklaFirm";
    private const string DefaultTeklaPublishSourcePath = @"\\62.113.36.107\BIM_Models\Tekla\XS_FIRM";

    private readonly SettingsService _settingsService = new();
    private readonly AutoStartService _autoStartService = new();
    private readonly HeartbeatClient _heartbeatClient = new(new HttpClient { Timeout = TimeSpan.FromSeconds(110) });
    private readonly UpdateService _updateService = new(new HttpClient { Timeout = TimeSpan.FromSeconds(40) });
    private readonly TeklaStandardService _teklaStandardService = new(new HttpClient { Timeout = TimeSpan.FromSeconds(25) });
    private readonly DispatcherTimer _timer = new();
    private readonly DispatcherTimer _updateTimer = new();
    private readonly Forms.NotifyIcon _trayIcon;

    private AppSettings _settings = new();
    private bool _isRunning;
    private bool _allowClose;
    private bool _trayHintShown;
    private string _activeSessionId = string.Empty;
    private UpdateManifest? _pendingUpdate;
    private string? _downloadedInstallerPath;
    private bool _updateOfferShown;
    private string _lastUpdateBalloonVersion = string.Empty;
    private string _lastUpdateWindowVersion = string.Empty;
    private bool _updateCheckInProgress;
    private bool _teklaCheckInProgress;
    private bool _teklaBalloonShown;
    private bool _serverConnectionFailed;
    private static readonly TimeSpan UpdateCheckInterval = TimeSpan.FromMinutes(30);

    public MainWindow()
    {
        InitializeComponent();
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        _timer.Tick += Timer_Tick;
        _updateTimer.Tick += UpdateTimer_Tick;
        Closing += MainWindow_Closing;
        Closed += MainWindow_Closed;
        StateChanged += MainWindow_StateChanged;
        _trayIcon = CreateTrayIcon();
        LoadSettingsToUi();
        UpdateRunStateUi();
        UpdateActionButtonUi();
        UpdateTeklaUi();
        UpdateHeaderStatusUi();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        Topmost = true;
        Activate();
        Focus();
        Topmost = false;
        AppendLog("При закрытии окно сворачивается в трей. Для полного выхода: иконка в трее -> Закрыть.");
        if (_teklaStandardService.CheckGitAvailability(out var gitPath, out var gitDetails))
        {
            AppendLog("Стандарт Tekla: git доступен (" + gitPath + ") " + gitDetails);
        }
        else
        {
            AppendLog("Стандарт Tekla: git недоступен (" + gitPath + ") " + gitDetails);
        }
        _ = TryAutoConnectAsync();
        _ = CheckUpdatesAsync(showDialogs: false);
        _ = CheckTeklaStandardAsync(showDialogs: false, autoApplyIfPossible: false);
        _updateTimer.Interval = UpdateCheckInterval;
        _updateTimer.Start();
    }

    private async Task CheckAndOfferUpdatesAsync()
    {
        await CheckUpdatesAsync(showDialogs: false);
        await OfferUpdateInstallIfAvailableAsync();
    }

    private async Task OfferUpdateInstallIfAvailableAsync()
    {
        if (_updateOfferShown)
        {
            return;
        }

        if (_pendingUpdate is null)
        {
            return;
        }

        _updateOfferShown = true;

        var result = ThemedDialogs.Show(this,
            "Доступна новая версия Structura Connector. Установить обновление сейчас?",
            "Обновление доступно",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            UpdateStateTextBlock.Text = "Обновление: загрузка установщика...";
            _downloadedInstallerPath = await _updateService.DownloadInstallerAsync(_pendingUpdate, CancellationToken.None);
            AppendLog("Скачан установщик обновления: " + _downloadedInstallerPath);
            UpdateService.RunInstaller(_downloadedInstallerPath);
            ExitFromTray();
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка автообновления: " + ex.Message);
            UpdateStateTextBlock.Text = "Обновление: ошибка установки";
            _updateOfferShown = false;
        }
    }

    private void ShowUpdateAvailableBalloon(UpdateManifest manifest)
    {
        var version = manifest.Version?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(version))
        {
            return;
        }

        if (string.Equals(_lastUpdateBalloonVersion, version, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _lastUpdateBalloonVersion = version;
        _trayIcon.ShowBalloonTip(
            4000,
            "Structura Connector",
            "Доступна новая версия обновления: " + version + ". Откройте приложение и нажмите 'Проверить обновление'.",
            Forms.ToolTipIcon.Info);
    }

    private void ShowUpdateAvailableWindowMessage(UpdateManifest manifest)
    {
        var version = manifest.Version?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(version))
        {
            return;
        }

        if (string.Equals(_lastUpdateWindowVersion, version, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        _lastUpdateWindowVersion = version;
        ThemedDialogs.Show(this,
            "Доступна новая версия Structura Connector: " + version + ".\nНажмите кнопку 'Скачать и установить'.",
            "Доступно обновление",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private async Task TryAutoConnectAsync()
    {
        try
        {
            var token = SettingsService.DecryptToken(_settings.TokenCipherBase64).Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                AppendLog("Сохраненного токена нет. Введите токен вручную.");
                return;
            }

            AppendLog("Найден сохраненный токен. Запускаю автоподключение...");
            await ConnectByTokenInternalAsync(token, showSuccessDialog: false);
        }
        catch (TaskCanceledException)
        {
            AppendLog("Автоподключение не выполнено: сервер ответил слишком медленно. Повторите подключение через кнопку.");
            _serverConnectionFailed = true;
            UpdateHeaderStatusUi();
        }
        catch (Exception ex)
        {
            AppendLog("Автоподключение не выполнено: " + ex.Message);
            _serverConnectionFailed = true;
            UpdateHeaderStatusUi();
        }
    }

    private Forms.NotifyIcon CreateTrayIcon()
    {
        var menu = new Forms.ContextMenuStrip();

        var openItem = new Forms.ToolStripMenuItem("Открыть Structura Connector");
        openItem.Click += (_, _) => ShowFromTray();

        var closeItem = new Forms.ToolStripMenuItem("Закрыть");
        closeItem.Click += (_, _) => ExitFromTray();

        menu.Items.Add(openItem);
        menu.Items.Add(new Forms.ToolStripSeparator());
        menu.Items.Add(closeItem);

        var icon = TryGetTrayIcon();
        var tray = new Forms.NotifyIcon
        {
            Icon = icon,
            Text = "Structura Connector",
            Visible = true,
            ContextMenuStrip = menu
        };
        tray.DoubleClick += (_, _) => ShowFromTray();
        return tray;
    }

    private static System.Drawing.Icon TryGetTrayIcon()
    {
        try
        {
            var exePath = Environment.ProcessPath;
            if (!string.IsNullOrWhiteSpace(exePath))
            {
                var extracted = System.Drawing.Icon.ExtractAssociatedIcon(exePath);
                if (extracted is not null)
                {
                    return extracted;
                }
            }
        }
        catch
        {
            // Ignore icon extraction errors and use fallback icon.
        }

        return System.Drawing.SystemIcons.Application;
    }

    private void MainWindow_StateChanged(object? sender, EventArgs e)
    {
        if (WindowState == WindowState.Minimized)
        {
            HideToTray();
        }
    }

    private void MainWindow_Closing(object? sender, CancelEventArgs e)
    {
        if (_allowClose)
        {
            return;
        }

        e.Cancel = true;
        HideToTray();
    }

    private void MainWindow_Closed(object? sender, EventArgs e)
    {
        _updateTimer.Stop();
        _trayIcon.Visible = false;
        _trayIcon.Dispose();
    }

    private async void UpdateTimer_Tick(object? sender, EventArgs e)
    {
        await CheckUpdatesAsync(showDialogs: false);
        await CheckTeklaStandardAsync(showDialogs: false, autoApplyIfPossible: false);
    }

    private void UpdateActionButtonUi()
    {
        if (_pendingUpdate is null)
        {
            UpdateActionButton.Content = "Проверить обновление";
            UpdateActionButton.Style = (Style)FindResource("SecondaryButton");
            return;
        }

        UpdateActionButton.Content = "Скачать и установить";
        UpdateActionButton.Style = (Style)FindResource("PrimaryButton");
    }

    private void UpdateTeklaUi()
    {
        var targetRevision = string.IsNullOrWhiteSpace(_settings.TeklaStandardTargetRevision)
            ? "-"
            : _settings.TeklaStandardTargetRevision.Trim();
        var hasTargetRevision = !string.IsNullOrWhiteSpace(_settings.TeklaStandardTargetRevision);
        TeklaCurrentVersionTextBlock.Text = targetRevision;
        TeklaUpToDateTextBlock.Text = hasTargetRevision && !_teklaStandardService.IsUpdateAvailable(_settings) ? "да" : "нет";
        TeklaRoleTextBlock.Text = _settings.IsFirmAdmin ? "Роль администратора: да" : "Роль администратора: нет";
        TeklaLocalPathInfoTextBlock.Text = "Папка стандарта на этом ПК: " +
                                           (string.IsNullOrWhiteSpace(_settings.TeklaStandardLocalPath)
                                               ? DefaultTeklaStandardLocalPath
                                               : _settings.TeklaStandardLocalPath);
        UpdateTeklaActionButtonUi();

        TeklaPublishPanel.Visibility = _settings.IsFirmAdmin ? Visibility.Visible : Visibility.Collapsed;
        TeklaPublishButton.IsEnabled = _settings.IsFirmAdmin;
        if (string.IsNullOrWhiteSpace(TeklaPublishSourcePathTextBox.Text))
        {
            TeklaPublishSourcePathTextBox.Text = string.IsNullOrWhiteSpace(_settings.TeklaPublishSourcePath)
                ? DefaultTeklaPublishSourcePath
                : _settings.TeklaPublishSourcePath;
        }
        if (string.IsNullOrWhiteSpace(TeklaPublishNotesTextBox.Text))
        {
            TeklaPublishNotesTextBox.Text = "Публикация из Structura Connector";
        }

        var canRestartTeklaServer = _settings.IsSystemAdmin || _settings.IsFirmAdmin;
        ServerActionsPanel.Visibility = canRestartTeklaServer ? Visibility.Visible : Visibility.Collapsed;
        RestartTeklaServerButton.IsEnabled = canRestartTeklaServer;
        UpdateHeaderStatusUi();
    }

    private void UpdateTeklaActionButtonUi()
    {
        var hasUpdate = !string.IsNullOrWhiteSpace(_settings.TeklaStandardTargetRevision) &&
                        _teklaStandardService.IsUpdateAvailable(_settings);

        TeklaCheckButton.Content = hasUpdate ? "Обновить" : "Проверить обновление";
        TeklaCheckButton.Style = (Style)FindResource(hasUpdate ? "PrimaryButton" : "SecondaryButton");
        TeklaCheckButton.IsEnabled = !_teklaCheckInProgress;
    }

    private void UpdateHeaderStatusUi()
    {
        var hasToken = !string.IsNullOrWhiteSpace(SettingsService.DecryptToken(_settings.TokenCipherBase64));
        if (!hasToken)
        {
            HeaderServerStatusTextBlock.Text = "Сервер: подключение не выполнено";
            HeaderServerStatusTextBlock.Foreground = System.Windows.Media.Brushes.DarkGray;
        }
        else if (_isRunning && !string.IsNullOrWhiteSpace(_activeSessionId) && !_serverConnectionFailed)
        {
            HeaderServerStatusTextBlock.Text = "Сервер: подключено";
            HeaderServerStatusTextBlock.Foreground = System.Windows.Media.Brushes.MediumSpringGreen;
        }
        else if (_serverConnectionFailed)
        {
            HeaderServerStatusTextBlock.Text = "Сервер: подключение не выполнено";
            HeaderServerStatusTextBlock.Foreground = System.Windows.Media.Brushes.Orange;
        }
        else
        {
            HeaderServerStatusTextBlock.Text = "Сервер: проверка подключения...";
            HeaderServerStatusTextBlock.Foreground = System.Windows.Media.Brushes.Gainsboro;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaStandardTargetRevision))
        {
            HeaderFirmStatusTextBlock.Text = "Папка фирмы: статус неизвестен";
            HeaderFirmStatusTextBlock.Foreground = System.Windows.Media.Brushes.DarkGray;
            return;
        }

        if (_teklaStandardService.IsUpdateAvailable(_settings))
        {
            HeaderFirmStatusTextBlock.Text = "Папка фирмы: требуется обновление";
            HeaderFirmStatusTextBlock.Foreground = System.Windows.Media.Brushes.Orange;
            return;
        }

        HeaderFirmStatusTextBlock.Text = "Папка фирмы: актуальна";
        HeaderFirmStatusTextBlock.Foreground = System.Windows.Media.Brushes.MediumSpringGreen;
    }

    private void ShowTeklaPendingBalloon(string revision)
    {
        if (_teklaBalloonShown)
        {
            return;
        }

        _teklaBalloonShown = true;
        _trayIcon.ShowBalloonTip(
            3000,
            "Стандарт Tekla",
            "Найдена ревизия " + revision + ". Закройте Tekla и нажмите 'Обновить сейчас' или дождитесь автоустановки при следующей проверке.",
            Forms.ToolTipIcon.Info);
    }

    private void HideToTray()
    {
        Hide();
        ShowInTaskbar = false;

        if (!_trayHintShown)
        {
            _trayHintShown = true;
            _trayIcon.ShowBalloonTip(2500, "Structura Connector", "Приложение работает в трее. ПКМ по иконке -> Закрыть.", Forms.ToolTipIcon.Info);
        }
    }

    private void ShowFromTray()
    {
        ShowInTaskbar = true;
        Show();
        WindowState = WindowState.Normal;
        Activate();
    }

    private void ExitFromTray()
    {
        _allowClose = true;
        _timer.Stop();
        _updateTimer.Stop();
        _isRunning = false;
        Close();
    }

    private void LoadSettingsToUi()
    {
        _settings = _settingsService.Load();
        var shouldPersist = false;

        if (!string.IsNullOrWhiteSpace(_settings.SmbLogin))
        {
            _settings.SmbLogin = string.Empty;
            shouldPersist = true;
        }

        if (!string.IsNullOrWhiteSpace(_settings.SmbPasswordCipherBase64))
        {
            _settings.SmbPasswordCipherBase64 = string.Empty;
            shouldPersist = true;
        }

        _settings.ServerUrl = FixedServerUrl;
        _settings.UpdateManifestUrl = FixedUpdateManifestUrl;
        _settings.AutoStart = true;
        if (_settings.HeartbeatSeconds < 10)
        {
            _settings.HeartbeatSeconds = FixedHeartbeatSeconds;
        }
        if (string.IsNullOrWhiteSpace(_settings.SmbSharePath))
        {
            _settings.SmbSharePath = DefaultSmbSharePath;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaStandardManifestUrl))
        {
            _settings.TeklaStandardManifestUrl = FixedTeklaStandardManifestUrl;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaStandardLocalPath))
        {
            _settings.TeklaStandardLocalPath = DefaultTeklaStandardLocalPath;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaPublishSourcePath))
        {
            _settings.TeklaPublishSourcePath = DefaultTeklaPublishSourcePath;
            shouldPersist = true;
        }

        ServerUrlTextBox.Text = _settings.ServerUrl;
        UpdateManifestUrlTextBox.Text = _settings.UpdateManifestUrl;
        DeviceIdTextBox.Text = _settings.DeviceId;
        SmbLoginTextBox.Text = string.Empty;
        SmbSharePathTextBox.Text = _settings.SmbSharePath;
        IntervalTextBox.Text = _settings.HeartbeatSeconds.ToString();
        AutoStartCheckBox.IsChecked = true;
        TeklaPublishSourcePathTextBox.Text = _settings.TeklaPublishSourcePath;

        var token = SettingsService.DecryptToken(_settings.TokenCipherBase64);
        TokenPasswordBox.Password = token;

        SmbPasswordBox.Password = string.Empty;

        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);

        if (shouldPersist)
        {
            _settingsService.Save(_settings);
        }

        AppendLog($"Настройки загружены: {_settingsService.SettingsPath}");
    }

    private AppSettings ReadSettingsFromUi()
    {
        var token = TokenPasswordBox.Password.Trim();
        if (string.IsNullOrWhiteSpace(token))
        {
            throw new InvalidOperationException("Токен не может быть пустым.");
        }

        var deviceId = string.IsNullOrWhiteSpace(_settings.DeviceId)
            ? "pc-" + Environment.MachineName.ToLowerInvariant()
            : _settings.DeviceId;

        var sec = _settings.HeartbeatSeconds >= 10 ? _settings.HeartbeatSeconds : FixedHeartbeatSeconds;
        var smbSharePath = string.IsNullOrWhiteSpace(_settings.SmbSharePath) ? DefaultSmbSharePath : _settings.SmbSharePath;
        var smbLogin = _settings.SmbLogin;
        var smbPassword = SettingsService.DecryptToken(_settings.SmbPasswordCipherBase64);
        var teklaManifestUrl = string.IsNullOrWhiteSpace(_settings.TeklaStandardManifestUrl)
            ? FixedTeklaStandardManifestUrl
            : _settings.TeklaStandardManifestUrl;
        var teklaLocalPath = string.IsNullOrWhiteSpace(_settings.TeklaStandardLocalPath)
            ? DefaultTeklaStandardLocalPath
            : _settings.TeklaStandardLocalPath;
        var teklaPublishSourcePath = string.IsNullOrWhiteSpace(TeklaPublishSourcePathTextBox.Text)
            ? (string.IsNullOrWhiteSpace(_settings.TeklaPublishSourcePath)
                ? DefaultTeklaPublishSourcePath
                : _settings.TeklaPublishSourcePath)
            : TeklaPublishSourcePathTextBox.Text.Trim();

        return new AppSettings
        {
            ServerUrl = FixedServerUrl,
            UpdateManifestUrl = string.IsNullOrWhiteSpace(_settings.UpdateManifestUrl)
                ? FixedUpdateManifestUrl
                : _settings.UpdateManifestUrl,
            DeviceId = deviceId,
            TokenCipherBase64 = SettingsService.EncryptToken(token),
            SmbLogin = smbLogin,
            SmbPasswordCipherBase64 = string.IsNullOrWhiteSpace(smbPassword)
                ? string.Empty
                : SettingsService.EncryptToken(smbPassword),
            SmbSharePath = smbSharePath,
            HeartbeatSeconds = sec,
            AutoStart = true,
            TeklaStandardManifestUrl = teklaManifestUrl,
            TeklaStandardLocalPath = teklaLocalPath,
            TeklaStandardInstalledRevision = _settings.TeklaStandardInstalledRevision,
            TeklaStandardTargetRevision = _settings.TeklaStandardTargetRevision,
            TeklaStandardLastCheckUtc = _settings.TeklaStandardLastCheckUtc,
            TeklaStandardLastSuccessUtc = _settings.TeklaStandardLastSuccessUtc,
            TeklaStandardPendingAfterClose = _settings.TeklaStandardPendingAfterClose,
            TeklaStandardLastError = _settings.TeklaStandardLastError,
            TeklaStandardRepoUrl = _settings.TeklaStandardRepoUrl,
            TeklaStandardRepoRef = _settings.TeklaStandardRepoRef,
            TeklaPublishSourcePath = teklaPublishSourcePath,
            IsSystemAdmin = _settings.IsSystemAdmin,
            IsFirmAdmin = _settings.IsFirmAdmin
        };
    }

    private void ApplyAndPersist()
    {
        _settings = ReadSettingsFromUi();
        _settingsService.Save(_settings);
        _autoStartService.SetEnabled(_settings.AutoStart);
        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);
        UpdateTeklaUi();
        AppendLog("Настройки сохранены.");
    }

    private void UpdateRunStateUi()
    {
        if (_isRunning)
        {
            RunStateTextBlock.Text = "Автоотправка heartbeat: включена";
            RunStateTextBlock.Foreground = System.Windows.Media.Brushes.MediumSpringGreen;
        }
        else
        {
            RunStateTextBlock.Text = "Автоотправка heartbeat: выключена";
            RunStateTextBlock.Foreground = System.Windows.Media.Brushes.Orange;
        }

        StartButton.IsEnabled = !_isRunning;
        StopButton.IsEnabled = _isRunning;
        UpdateHeaderStatusUi();
    }

    private async Task SendHeartbeatSafeAsync()
    {
        try
        {
            var token = SettingsService.DecryptToken(_settings.TokenCipherBase64);
            var teklaState = new TeklaHeartbeatState
            {
                InstalledVersion = _settings.TeklaStandardInstalledVersion,
                TargetVersion = _settings.TeklaStandardTargetVersion,
                InstalledRevision = _settings.TeklaStandardInstalledRevision,
                TargetRevision = _settings.TeklaStandardTargetRevision,
                PendingAfterClose = false,
                TeklaRunning = false,
                LastCheckUtc = _settings.TeklaStandardLastCheckUtc?.UtcDateTime.ToString("o") ?? string.Empty,
                LastSuccessUtc = _settings.TeklaStandardLastSuccessUtc?.UtcDateTime.ToString("o") ?? string.Empty,
                LastError = _settings.TeklaStandardLastError
            };

            await _heartbeatClient.SendHeartbeatAsync(
                _settings.ServerUrl,
                _settings.DeviceId,
                token,
                _activeSessionId,
                teklaState,
                CancellationToken.None);
            _serverConnectionFailed = false;
            AppendLog("Heartbeat отправлен успешно.");
            UpdateHeaderStatusUi();
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("HTTP 409", StringComparison.OrdinalIgnoreCase))
            {
                _timer.Stop();
                _isRunning = false;
                _serverConnectionFailed = true;
                UpdateRunStateUi();
                AppendLog("Сессия отключена: этот токен активирован на другом устройстве.");
                return;
            }

            _serverConnectionFailed = true;
            AppendLog("Ошибка heartbeat: " + ex.Message);
            UpdateHeaderStatusUi();
        }
    }

    private void AppendLog(string text)
    {
        LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {text}{Environment.NewLine}");
        LogTextBox.ScrollToEnd();
    }

    private async void Timer_Tick(object? sender, EventArgs e)
    {
        await SendHeartbeatSafeAsync();
    }

    private async void Start_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ApplyAndPersist();
            _timer.Start();
            _isRunning = true;
            UpdateRunStateUi();
            AppendLog("Фоновая отправка запущена.");
            await SendHeartbeatSafeAsync();
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка запуска: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Stop_Click(object sender, RoutedEventArgs e)
    {
        _timer.Stop();
        _isRunning = false;
        UpdateRunStateUi();
        AppendLog("Фоновая отправка остановлена.");
    }

    private async void SendNow_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ApplyAndPersist();
            await SendHeartbeatSafeAsync();
        }
        catch (Exception ex)
        {
            ThemedDialogs.Show(this, ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void TestConnection_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var serverUrl = ServerUrlTextBox.Text.Trim();
            if (!Uri.TryCreate(serverUrl, UriKind.Absolute, out _))
            {
                throw new InvalidOperationException("Введите корректный URL сервера.");
            }

            await _heartbeatClient.CheckServerHealthAsync(serverUrl, CancellationToken.None);
            try
            {
                var ip = await _heartbeatClient.ResolvePublicIpAsync(CancellationToken.None);
                AppendLog("Подключение к серверу проверено. Внешний IP: " + ip);
                ThemedDialogs.Show(this, 
                    "Сервер доступен и отвечает /health.\nВнешний IP: " + ip,
                    "Проверка подключения",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ipEx)
            {
                AppendLog("Сервер доступен, но внешний IP определить не удалось: " + ipEx.Message);
                ThemedDialogs.Show(this, 
                    "Сервер доступен и отвечает /health.\n" +
                    "Но внешний IP определить не удалось, поэтому отправка heartbeat может не работать.\n\n" +
                    ipEx.Message,
                    "Проверка подключения",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка проверки подключения: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Ошибка проверки подключения", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ApplyAndPersist();
        }
        catch (Exception ex)
        {
            ThemedDialogs.Show(this, ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ConnectByToken_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var token = TokenPasswordBox.Password.Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidOperationException("Введите токен устройства.");
            }

            await ConnectByTokenInternalAsync(token, showSuccessDialog: true);
        }
        catch (TaskCanceledException)
        {
            const string message = "Сервер отвечает дольше обычного. Подождите немного и повторите подключение.";
            AppendLog("Ошибка автоподключения по токену: " + message);
            _serverConnectionFailed = true;
            UpdateHeaderStatusUi();
            ThemedDialogs.Show(this, message, "Время ожидания истекло", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка автоподключения по токену: " + ex.Message);
            _serverConnectionFailed = true;
            UpdateHeaderStatusUi();
            ThemedDialogs.Show(this, ex.Message, "Ошибка подключения", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task ConnectByTokenInternalAsync(string token, bool showSuccessDialog)
    {
        var serverUrl = FixedServerUrl;

        AppendLog("Запрошен bootstrap по токену...");
        var bootstrap = await _heartbeatClient.BootstrapAsync(serverUrl, token, CancellationToken.None);

        if (string.IsNullOrWhiteSpace(bootstrap.DeviceId))
        {
            throw new InvalidOperationException("Сервер вернул пустой device_id.");
        }

        var sharePath = bootstrap.SmbAccess.ShareUnc;
        if (string.IsNullOrWhiteSpace(sharePath))
        {
            sharePath = bootstrap.SmbAccess.SharePath;
        }

        if (string.IsNullOrWhiteSpace(bootstrap.SmbAccess.Login) ||
            string.IsNullOrWhiteSpace(bootstrap.SmbAccess.Password) ||
            string.IsNullOrWhiteSpace(sharePath))
        {
            throw new InvalidOperationException("Сервер не вернул полный набор SMB-данных для подключения.");
        }

        _settings = new AppSettings
        {
            ServerUrl = FixedServerUrl,
            UpdateManifestUrl = string.IsNullOrWhiteSpace(bootstrap.UpdateManifestUrl)
                ? FixedUpdateManifestUrl
                : bootstrap.UpdateManifestUrl,
            DeviceId = bootstrap.DeviceId,
            TokenCipherBase64 = SettingsService.EncryptToken(token),
            SmbLogin = string.Empty,
            SmbPasswordCipherBase64 = string.Empty,
            SmbSharePath = sharePath,
            HeartbeatSeconds = bootstrap.HeartbeatSeconds >= 10 ? bootstrap.HeartbeatSeconds : FixedHeartbeatSeconds,
            AutoStart = true,
            TeklaStandardManifestUrl = string.IsNullOrWhiteSpace(_settings.TeklaStandardManifestUrl)
                ? FixedTeklaStandardManifestUrl
                : _settings.TeklaStandardManifestUrl,
            TeklaStandardLocalPath = string.IsNullOrWhiteSpace(_settings.TeklaStandardLocalPath)
                ? DefaultTeklaStandardLocalPath
                : _settings.TeklaStandardLocalPath,
            TeklaStandardInstalledRevision = _settings.TeklaStandardInstalledRevision,
            TeklaStandardTargetRevision = _settings.TeklaStandardTargetRevision,
            TeklaStandardLastCheckUtc = _settings.TeklaStandardLastCheckUtc,
            TeklaStandardLastSuccessUtc = _settings.TeklaStandardLastSuccessUtc,
            TeklaStandardPendingAfterClose = _settings.TeklaStandardPendingAfterClose,
            TeklaStandardLastError = _settings.TeklaStandardLastError,
            TeklaStandardRepoUrl = _settings.TeklaStandardRepoUrl,
            TeklaStandardRepoRef = _settings.TeklaStandardRepoRef,
            TeklaPublishSourcePath = string.IsNullOrWhiteSpace(_settings.TeklaPublishSourcePath)
                ? DefaultTeklaPublishSourcePath
                : _settings.TeklaPublishSourcePath,
            IsSystemAdmin = bootstrap.IsSystemAdmin,
            IsFirmAdmin = bootstrap.IsFirmAdmin
        };
        _settingsService.Save(_settings);
        _autoStartService.SetEnabled(true);
        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);
        _activeSessionId = bootstrap.SessionId;
        _serverConnectionFailed = false;

        DeviceIdTextBox.Text = _settings.DeviceId;
        IntervalTextBox.Text = _settings.HeartbeatSeconds.ToString();
        UpdateManifestUrlTextBox.Text = _settings.UpdateManifestUrl;
        SmbLoginTextBox.Text = bootstrap.SmbAccess.Login;
        SmbPasswordBox.Password = string.Empty;
        SmbSharePathTextBox.Text = _settings.SmbSharePath;
        UpdateTeklaUi();
        AppendLog("Настройки сохранены.");

        var smbConnected = true;
        try
        {
            await ConnectSmbInternalAsync(bootstrap.SmbAccess.Login, bootstrap.SmbAccess.Password, sharePath, openExplorer: true);
        }
        catch (Exception ex) when (IsWindowsSmbConflict(ex))
        {
            smbConnected = false;
            AppendLog("SMB подключение не переключено автоматически (конфликт 1219). Текущая сессия SMB оставлена без изменений.");
            AppendLog("Детали SMB конфликта: " + ex.Message);
        }

        _timer.Stop();
        _timer.Start();
        _isRunning = true;
        UpdateRunStateUi();

        await SendHeartbeatSafeAsync();
        AppendLog("Автоподключение по токену выполнено успешно.");

        if (showSuccessDialog)
        {
            ThemedDialogs.Show(this, 
                smbConnected
                    ? "Подключение выполнено. SMB доступ открыт, автоотправка heartbeat включена."
                    : "Подключение выполнено. Автоотправка heartbeat включена, но SMB не переключен из-за активной сессии Windows (1219).",
                "Structura Connector",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }

    private async Task CheckUpdatesAsync(bool showDialogs)
    {
        if (_updateCheckInProgress)
        {
            return;
        }

        _updateCheckInProgress = true;
        UpdateActionButton.IsEnabled = false;
        try
        {
            var manifestUrl = string.IsNullOrWhiteSpace(_settings.UpdateManifestUrl)
                ? FixedUpdateManifestUrl
                : _settings.UpdateManifestUrl.Trim();
            if (!Uri.TryCreate(manifestUrl, UriKind.Absolute, out _))
            {
                throw new InvalidOperationException("Введите корректный адрес обновлений.");
            }

            _settings.UpdateManifestUrl = manifestUrl;
            UpdateManifestUrlTextBox.Text = manifestUrl;
            _settingsService.Save(_settings);

            var manifest = await _updateService.TryGetUpdateAsync(manifestUrl, CancellationToken.None);
            if (manifest is null)
            {
                _pendingUpdate = null;
                _lastUpdateBalloonVersion = string.Empty;
                _lastUpdateWindowVersion = string.Empty;
                UpdateStateTextBlock.Text = "Обновление: не удалось получить данные";
                UpdateActionButtonUi();
                if (showDialogs)
                {
                    ThemedDialogs.Show(this, "Не удалось проверить обновления.", "Обновления", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                return;
            }

            if (_updateService.IsUpdateAvailable(manifest))
            {
                _pendingUpdate = manifest;
                UpdateStateTextBlock.Text = $"Доступно обновление: {manifest.Version}";
                AppendLog("Найдено обновление: " + manifest.Version);
                ShowUpdateAvailableBalloon(manifest);
                if (!showDialogs)
                {
                    ShowUpdateAvailableWindowMessage(manifest);
                }
                UpdateActionButtonUi();
                if (showDialogs)
                {
                    ThemedDialogs.Show(this,
                        "Доступна новая версия: " + manifest.Version +
                        (string.IsNullOrWhiteSpace(manifest.Notes) ? "" : "\n\n" + manifest.Notes),
                        "Обновления",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            else
            {
                _pendingUpdate = null;
                _lastUpdateBalloonVersion = string.Empty;
                _lastUpdateWindowVersion = string.Empty;
                UpdateStateTextBlock.Text = $"Обновление: актуально ({_updateService.CurrentVersion})";
                UpdateActionButtonUi();
                if (showDialogs)
                {
                    ThemedDialogs.Show(this, "Установлена актуальная версия.", "Обновления", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }
        catch (Exception ex)
        {
            _pendingUpdate = null;
            UpdateStateTextBlock.Text = "Обновление: ошибка проверки";
            AppendLog("Ошибка проверки обновления: " + ex.Message);
            UpdateActionButtonUi();
            if (showDialogs)
            {
                ThemedDialogs.Show(this, ex.Message, "Ошибка обновлений", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        finally
        {
            _updateCheckInProgress = false;
            UpdateActionButton.IsEnabled = true;
        }
    }

    private async void UpdateAction_Click(object sender, RoutedEventArgs e)
    {
        if (_pendingUpdate is null)
        {
            await CheckUpdatesAsync(showDialogs: true);
            return;
        }

        await InstallPendingUpdateAsync();
    }

    private async Task InstallPendingUpdateAsync()
    {
        try
        {
            if (_pendingUpdate is null)
            {
                await CheckUpdatesAsync(showDialogs: true);
                if (_pendingUpdate is null)
                {
                    return;
                }
            }

            UpdateActionButton.IsEnabled = false;
            UpdateStateTextBlock.Text = "Обновление: загрузка установщика...";
            _downloadedInstallerPath = await _updateService.DownloadInstallerAsync(_pendingUpdate, CancellationToken.None);
            AppendLog("Скачан установщик обновления: " + _downloadedInstallerPath);

            var result = ThemedDialogs.Show(this,
                "Установщик скачан. Закрыть приложение и запустить обновление сейчас?",
                "Обновление",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                UpdateService.RunInstaller(_downloadedInstallerPath);
                ExitFromTray();
            }
            else
            {
                UpdateStateTextBlock.Text = "Обновление: установщик скачан";
                UpdateActionButton.IsEnabled = true;
            }
        }
        catch (Exception ex)
        {
            UpdateActionButton.IsEnabled = _pendingUpdate is not null;
            UpdateStateTextBlock.Text = "Обновление: ошибка установки";
            AppendLog("Ошибка установки обновления: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Ошибка обновления", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async Task CheckTeklaStandardAsync(bool showDialogs, bool autoApplyIfPossible)
    {
        if (_teklaCheckInProgress)
        {
            return;
        }

        _teklaCheckInProgress = true;
        TeklaCheckButton.IsEnabled = false;

        try
        {
            _settings.TeklaStandardManifestUrl = string.IsNullOrWhiteSpace(_settings.TeklaStandardManifestUrl)
                ? FixedTeklaStandardManifestUrl
                : _settings.TeklaStandardManifestUrl;
            _settings.TeklaStandardLocalPath = string.IsNullOrWhiteSpace(_settings.TeklaStandardLocalPath)
                ? DefaultTeklaStandardLocalPath
                : _settings.TeklaStandardLocalPath;

            var manifest = await _teklaStandardService.TryGetManifestAsync(_settings.TeklaStandardManifestUrl, CancellationToken.None);
            _settings.TeklaStandardLastCheckUtc = DateTimeOffset.UtcNow;

            if (manifest is null)
            {
                _settings.TeklaStandardTargetVersion = string.Empty;
                _settings.TeklaStandardTargetRevision = string.Empty;
                _settings.TeklaStandardLastError = "manifest_not_received";
                TeklaStatusTextBlock.Text = "Стандарт Tekla: данные обновления недоступны";
                _settingsService.Save(_settings);
                UpdateTeklaUi();
                if (showDialogs)
                {
                    ThemedDialogs.Show(this, "Не удалось получить данные обновления Стандарт Tekla.", "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                return;
            }

            _settings.TeklaStandardTargetVersion = manifest.Version;
            _settings.TeklaStandardTargetRevision = manifest.Revision;
            _settings.TeklaStandardRepoUrl = manifest.RepoUrl;
            _settings.TeklaStandardRepoRef = manifest.RepoRef;
            var updateAvailable = _teklaStandardService.IsUpdateAvailable(_settings);

            if (!updateAvailable)
            {
                _settings.TeklaStandardPendingAfterClose = false;
                _settings.TeklaStandardLastError = string.Empty;
                _teklaBalloonShown = false;
                TeklaStatusTextBlock.Text = "Стандарт Tekla: актуально (" + manifest.Revision + ")";
                _settingsService.Save(_settings);
                UpdateTeklaUi();
                if (showDialogs)
                {
                    ThemedDialogs.Show(this, "Стандарт Tekla уже актуален.", "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                return;
            }

            AppendLog("Найдена новая ревизия Стандарт Tekla: " + manifest.Revision);
            TeklaStatusTextBlock.Text = "Стандарт Tekla: доступна ревизия " + manifest.Revision;
            _settings.TeklaStandardPendingAfterClose = false;
            _settings.TeklaStandardLastError = string.Empty;
            _settingsService.Save(_settings);

            if (autoApplyIfPossible)
            {
                await ApplyTeklaStandardAsync(showDialogs: false, forceRefresh: false);
                return;
            }

            UpdateTeklaUi();
            if (showDialogs)
            {
                ThemedDialogs.Show(this, 
                    "Доступна новая ревизия Стандарт Tekla: " + manifest.Revision,
                    "Стандарт Tekla",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            TeklaStatusTextBlock.Text = "Стандарт Tekla: ошибка проверки";
            _settings.TeklaStandardLastError = ex.Message;
            _settingsService.Save(_settings);
            AppendLog("Ошибка проверки Стандарт Tekla: " + ex.Message);
            if (showDialogs)
            {
                ThemedDialogs.Show(this, ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        finally
        {
            _teklaCheckInProgress = false;
            TeklaCheckButton.IsEnabled = true;
            UpdateTeklaUi();
        }
    }

    private async Task ApplyTeklaStandardAsync(bool showDialogs, bool forceRefresh)
    {
        try
        {
            if (forceRefresh || string.IsNullOrWhiteSpace(_settings.TeklaStandardTargetRevision))
            {
                await CheckTeklaStandardAsync(showDialogs: false, autoApplyIfPossible: false);
                if (string.IsNullOrWhiteSpace(_settings.TeklaStandardTargetRevision))
                {
                    return;
                }
            }

            TeklaStatusTextBlock.Text = "Стандарт Tekla: обновление выполняется...";
            UpdateTeklaUi();
            var result = _teklaStandardService.ApplyPendingGitUpdate(_settings);

            if (result.IsSuccess)
            {
                _teklaStandardService.AppendLog(result.Message);
                AppendLog(result.Message);
                _teklaBalloonShown = false;
                _settings.TeklaStandardLastError = string.Empty;
                _settings.TeklaStandardInstalledVersion = _settings.TeklaStandardTargetVersion;
                TeklaStatusTextBlock.Text = "Стандарт Tekla: применена ревизия " + _settings.TeklaStandardInstalledRevision;
                _settingsService.Save(_settings);
                UpdateTeklaUi();
                if (showDialogs)
                {
                    ThemedDialogs.Show(this, result.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                return;
            }

            TeklaStatusTextBlock.Text = "Стандарт Tekla: " + result.Message;
            _settings.TeklaStandardLastError = result.Message;
            _settingsService.Save(_settings);
            UpdateTeklaUi();

            if (showDialogs)
            {
                ThemedDialogs.Show(this, result.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            TeklaStatusTextBlock.Text = "Стандарт Tekla: ошибка применения";
            _settings.TeklaStandardLastError = ex.Message;
            _settingsService.Save(_settings);
            AppendLog("Ошибка применения Стандарт Tekla: " + ex.Message);
            if (showDialogs)
            {
                ThemedDialogs.Show(this, ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private async void TeklaUpdateAction_Click(object sender, RoutedEventArgs e)
    {
        var hasUpdate = !string.IsNullOrWhiteSpace(_settings.TeklaStandardTargetRevision) &&
                        _teklaStandardService.IsUpdateAvailable(_settings);
        if (hasUpdate)
        {
            await ApplyTeklaStandardAsync(showDialogs: true, forceRefresh: false);
            return;
        }

        await CheckTeklaStandardAsync(showDialogs: true, autoApplyIfPossible: false);
    }

    private void TeklaOpenFolder_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var folderPath = string.IsNullOrWhiteSpace(_settings.TeklaStandardLocalPath)
                ? DefaultTeklaStandardLocalPath
                : _settings.TeklaStandardLocalPath.Trim();
            Directory.CreateDirectory(folderPath);
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = folderPath,
                UseShellExecute = true
            });
            AppendLog("Открыта папка Стандарт Tekla: " + folderPath);
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка открытия папки Стандарт Tekla: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
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
            AppendLog("Открыт лог Стандарт Tekla: " + _teklaStandardService.LogFilePath);
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка открытия лога Стандарт Tekla: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void TeklaPublish_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (!_settings.IsFirmAdmin)
            {
                throw new InvalidOperationException("Публикация доступна только для роли admin_firm.");
            }

            var token = SettingsService.DecryptToken(_settings.TokenCipherBase64).Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidOperationException("Токен устройства не найден. Выполните подключение по токену.");
            }

            var payload = new TeklaManifestPublishPayload
            {
                SourcePath = (TeklaPublishSourcePathTextBox.Text ?? string.Empty).Trim(),
                Comment = (TeklaPublishNotesTextBox.Text ?? string.Empty).Trim()
            };

            if (string.IsNullOrWhiteSpace(payload.SourcePath))
            {
                throw new InvalidOperationException("Укажите путь к эталонной папке XS_FIRM.");
            }
            if (string.IsNullOrWhiteSpace(payload.Comment))
            {
                throw new InvalidOperationException("Комментарий публикации обязателен.");
            }

            _settings.TeklaPublishSourcePath = payload.SourcePath;
            _settingsService.Save(_settings);

            TeklaPublishButton.IsEnabled = false;
            var result = await _heartbeatClient.PublishTeklaManifestAsync(_settings.ServerUrl, token, payload, CancellationToken.None);
            if (result.NoChanges)
            {
                AppendLog("Публикация XS_FIRM: изменений не обнаружено.");
                if (!string.IsNullOrWhiteSpace(result.Message))
                {
                    AppendLog(result.Message);
                }
            }
            else
            {
                AppendLog("Tekla manifest опубликован через desktop UI.");
                AppendLog("Версия: " + result.Version + "; ревизия: " + result.Revision);
                if (!string.IsNullOrWhiteSpace(result.Message))
                {
                    AppendLog(result.Message);
                }
            }
            await CheckTeklaStandardAsync(showDialogs: false, autoApplyIfPossible: false);

            ThemedDialogs.Show(this, 
                result.NoChanges
                    ? "Изменений в эталонной папке не найдено. Публикация не требуется."
                    : "Публикация Tekla manifest выполнена успешно.",
                "Стандарт Tekla",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            var message = GetFriendlyTeklaPublishErrorMessage(ex);
            AppendLog("Ошибка публикации Tekla manifest: " + message);
            ThemedDialogs.Show(this, message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            TeklaPublishButton.IsEnabled = _settings.IsFirmAdmin;
        }
    }

    private void TeklaPublishBrowse_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            using var dialog = new Forms.FolderBrowserDialog
            {
                Description = "Выберите папку фирмы для публикации",
                ShowNewFolderButton = false,
                SelectedPath = string.IsNullOrWhiteSpace(TeklaPublishSourcePathTextBox.Text)
                    ? (string.IsNullOrWhiteSpace(_settings.TeklaPublishSourcePath)
                        ? DefaultTeklaPublishSourcePath
                        : _settings.TeklaPublishSourcePath)
                    : TeklaPublishSourcePathTextBox.Text.Trim()
            };

            if (dialog.ShowDialog() == Forms.DialogResult.OK)
            {
                TeklaPublishSourcePathTextBox.Text = dialog.SelectedPath;
                _settings.TeklaPublishSourcePath = dialog.SelectedPath;
                _settingsService.Save(_settings);
            }
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка выбора папки фирмы: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Стандарт Tekla", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string GetFriendlyTeklaPublishErrorMessage(Exception ex)
    {
        var message = ex.Message?.Trim() ?? "Неизвестная ошибка публикации.";

        if (message.StartsWith("HTTP 409:", StringComparison.OrdinalIgnoreCase))
        {
            return "На сервере уже выполняется публикация XS_FIRM. Дождитесь завершения текущей попытки и только потом повторите запуск.";
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
        await RestartManagedServerAsync("tekla", RestartTeklaServerButton, "Tekla Server");
    }

    private async Task RestartManagedServerAsync(string serviceKey, System.Windows.Controls.Button button, string displayName)
    {
        try
        {
            var canRestart = serviceKey.Equals("tekla", StringComparison.OrdinalIgnoreCase)
                ? (_settings.IsSystemAdmin || _settings.IsFirmAdmin)
                : _settings.IsSystemAdmin;
            if (!canRestart)
            {
                throw new InvalidOperationException(
                    serviceKey.Equals("tekla", StringComparison.OrdinalIgnoreCase)
                        ? "Перезапуск Tekla Server доступен только администратору Tekla или системному администратору."
                        : "Перезапуск Revit Server доступен только системному администратору.");
            }

            var token = SettingsService.DecryptToken(_settings.TokenCipherBase64).Trim();
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new InvalidOperationException("Токен устройства не найден. Выполните подключение по токену.");
            }

            button.IsEnabled = false;
            AppendLog("Запущен перезапуск службы: " + displayName);
            var result = await _heartbeatClient.RestartManagedServiceAsync(_settings.ServerUrl, token, serviceKey, CancellationToken.None);
            AppendLog("Служба перезапущена: " + displayName + "; ответ сервера: " + result.Result.ToString());
            ThemedDialogs.Show(this, 
                displayName + " успешно перезапущен.",
                "Серверные действия",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка перезапуска службы " + displayName + ": " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Серверные действия", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            RestartTeklaServerButton.IsEnabled = _settings.IsSystemAdmin || _settings.IsFirmAdmin;
        }
    }

    private static string GetSmbHost(string sharePath)
    {
        if (!sharePath.StartsWith("\\\\", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("SMB путь должен начинаться с \\\\, например \\\\62.113.36.107\\BIM_Models");
        }

        var parts = sharePath.TrimStart('\\').Split('\\', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
        {
            throw new InvalidOperationException("SMB путь должен содержать сервер и имя шары.");
        }

        return parts[0];
    }

    private static string GetSmbShareRoot(string sharePath)
    {
        var parts = sharePath.TrimStart('\\').Split('\\', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 2)
        {
            throw new InvalidOperationException("SMB путь должен содержать сервер и имя шары.");
        }

        return $@"\\{parts[0]}\{parts[1]}";
    }

    private static string NormalizeSmbLogin(string login, string host)
    {
        if (string.IsNullOrWhiteSpace(login))
        {
            return login;
        }

        if (login.Contains('@'))
        {
            return login;
        }

        if (login.Contains('\\'))
        {
            var parts = login.Split('\\', 2, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 2)
            {
                var prefix = parts[0].Trim();
                var user = parts[1].Trim();
                if (string.Equals(prefix, host, StringComparison.OrdinalIgnoreCase))
                {
                    return user;
                }
            }
            return login;
        }

        return $"{host}\\{login}";
    }

    private static List<string> BuildSmbLoginCandidates(string login, string host)
    {
        var candidates = new List<string>();

        void Add(string value)
        {
            var v = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(v))
            {
                return;
            }
            if (!candidates.Contains(v, StringComparer.OrdinalIgnoreCase))
            {
                candidates.Add(v);
            }
        }

        Add(NormalizeSmbLogin(login, host));
        Add(login);

        if (login.Contains('\\'))
        {
            var idx = login.LastIndexOf('\\');
            if (idx >= 0 && idx + 1 < login.Length)
            {
                Add(login[(idx + 1)..]);
            }
        }

        if (!login.Contains('\\') && !login.Contains('@'))
        {
            Add($"{host}\\{login}");
        }

        return candidates;
    }

    private static void ConnectShareWithAnyLogin(string shareRoot, string password, IEnumerable<string> loginCandidates)
    {
        Exception? last = null;
        foreach (var candidate in loginCandidates)
        {
            try
            {
                RunProcessOrThrow("net", "use", shareRoot, password, $"/user:{candidate}", "/persistent:no");
                return;
            }
            catch (InvalidOperationException ex)
            {
                last = ex;
            }
        }

        if (last is not null)
        {
            throw last;
        }

        throw new InvalidOperationException("Не удалось выполнить SMB вход: отсутствуют варианты логина.");
    }

    private static (int ExitCode, string Output, string Error) RunProcess(string fileName, params string[] args)
    {
        var psi = new ProcessStartInfo(fileName)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.GetEncoding(866),
            StandardErrorEncoding = Encoding.GetEncoding(866)
        };

        foreach (var arg in args)
        {
            psi.ArgumentList.Add(arg);
        }

        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Не удалось запустить процесс.");
        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();

        return (process.ExitCode, NormalizeCliMessage(output), NormalizeCliMessage(error));
    }

    private static void RunProcessOrThrow(string fileName, params string[] args)
    {
        var result = RunProcess(fileName, args);

        if (result.ExitCode != 0)
        {
            var details = string.IsNullOrWhiteSpace(result.Error) ? result.Output : result.Error;
            throw new InvalidOperationException(details);
        }
    }

    private static string NormalizeCliMessage(string value)
    {
        var text = (value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return "Неизвестная ошибка командной строки.";
        }

        return text.Replace("\r", string.Empty).Trim();
    }

    private static bool IsWindowsSmbConflict(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        return message.Contains("1219", StringComparison.OrdinalIgnoreCase) ||
               message.Contains("множественное подключение", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsWindowsSmbConflict(Exception ex)
    {
        if (ex is null)
        {
            return false;
        }

        if (IsWindowsSmbConflict(ex.Message))
        {
            return true;
        }

        if (ex is AggregateException agg)
        {
            foreach (var inner in agg.Flatten().InnerExceptions)
            {
                if (IsWindowsSmbConflict(inner))
                {
                    return true;
                }
            }
        }

        return ex.InnerException is not null && IsWindowsSmbConflict(ex.InnerException);
    }

    private static bool IsWindowsNetConnectionNotFound(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        return message.Contains("2250", StringComparison.OrdinalIgnoreCase) ||
               message.Contains("не удалось найти сетевое подключение", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsWindowsNetNoEntries(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        return message.Contains("2250", StringComparison.OrdinalIgnoreCase) ||
               message.Contains("нет записей", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> ExtractUncPaths(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            yield break;
        }

        var matches = Regex.Matches(text, @"\\\\[^\s]+\\[^\s]+");
        foreach (Match match in matches)
        {
            var path = match.Value.Trim();
            if (!string.IsNullOrWhiteSpace(path))
            {
                yield return path;
            }
        }
    }

    private static void DisconnectAllSmbSessions()
    {
        try
        {
            RunProcessOrThrow("net", "use", "*", "/delete", "/y");
        }
        catch (InvalidOperationException ex) when (IsWindowsNetNoEntries(ex.Message) || IsWindowsNetConnectionNotFound(ex.Message))
        {
            // No active SMB sessions in current user profile.
        }
    }

    private static void DisconnectAllSmbSessionsForHost(string host)
    {
        var list = RunProcess("net", "use");
        var hostPrefix = $@"\\{host}\";
        var hostPaths = ExtractUncPaths(list.Output)
            .Concat(ExtractUncPaths(list.Error))
            .Where(path => path.StartsWith(hostPrefix, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var path in hostPaths)
        {
            try
            {
                RunProcessOrThrow("net", "use", path, "/delete", "/y");
            }
            catch (InvalidOperationException ex) when (IsWindowsNetConnectionNotFound(ex.Message) || IsWindowsNetNoEntries(ex.Message))
            {
                // Path already disconnected.
            }
        }

        try
        {
            RunProcessOrThrow("net", "use", hostPrefix + "*", "/delete", "/y");
        }
        catch (InvalidOperationException ex) when (IsWindowsNetConnectionNotFound(ex.Message) || IsWindowsNetNoEntries(ex.Message))
        {
            // Fallback wildcard returned no active entries.
        }
    }

    private static void DeleteStoredWindowsCredentialForHost(string host)
    {
        var targets = new[]
        {
            host,
            $"TERMSRV/{host}",
            $"Microsoft_Windows_Network/{host}"
        };

        foreach (var target in targets)
        {
            var result = RunProcess("cmdkey", $"/delete:{target}");
            if (result.ExitCode == 0)
            {
                continue;
            }

            var details = string.IsNullOrWhiteSpace(result.Error) ? result.Output : result.Error;
            if (details.Contains("не найден", StringComparison.OrdinalIgnoreCase) ||
                details.Contains("cannot find", StringComparison.OrdinalIgnoreCase) ||
                details.Contains("1168", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }
        }
    }

    private async Task ConnectSmbInternalAsync(string login, string password, string sharePath, bool openExplorer)
    {
        if (string.IsNullOrWhiteSpace(login) || string.IsNullOrWhiteSpace(password))
        {
            throw new InvalidOperationException("Введите SMB логин и пароль.");
        }

        var host = GetSmbHost(sharePath);
        var shareRoot = GetSmbShareRoot(sharePath);
        var loginCandidates = BuildSmbLoginCandidates(login, host);
        SmbLoginTextBox.Text = loginCandidates.FirstOrDefault() ?? login;

        await Task.Run(() =>
        {
            DeleteStoredWindowsCredentialForHost(host);

            try
            {
                RunProcessOrThrow("net", "use", shareRoot, "/delete", "/y");
            }
            catch
            {
                // Ignore cleanup errors for non-existing mappings.
            }

            try
            {
                ConnectShareWithAnyLogin(shareRoot, password, loginCandidates);
            }
            catch (InvalidOperationException ex) when (IsWindowsSmbConflict(ex.Message))
            {
                DisconnectAllSmbSessionsForHost(host);
                try
                {
                    ConnectShareWithAnyLogin(shareRoot, password, loginCandidates);
                }
                catch (InvalidOperationException retryEx) when (IsWindowsSmbConflict(retryEx.Message))
                {
                    DisconnectAllSmbSessions();
                    try
                    {
                        RunProcessOrThrow("net", "use", $@"\\{host}\IPC$", "/delete", "/y");
                    }
                    catch (InvalidOperationException ipcEx) when (IsWindowsNetConnectionNotFound(ipcEx.Message) || IsWindowsNetNoEntries(ipcEx.Message))
                    {
                        // IPC session not present.
                    }
                    ConnectShareWithAnyLogin(shareRoot, password, loginCandidates);
                }
            }
        });

        AppendLog($"SMB вход выполнен: {shareRoot}");

        if (openExplorer)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = sharePath,
                UseShellExecute = true
            });
        }
    }

    private async void ConnectSmb_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            ApplyAndPersist();

            var login = SmbLoginTextBox.Text.Trim();
            var password = SmbPasswordBox.Password.Trim();
            var sharePath = SmbSharePathTextBox.Text.Trim();
            await ConnectSmbInternalAsync(login, password, sharePath, openExplorer: true);
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка SMB входа: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Ошибка SMB входа", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void OpenSmbFolder_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var sharePath = SmbSharePathTextBox.Text.Trim();
            if (string.IsNullOrWhiteSpace(sharePath))
            {
                throw new InvalidOperationException("Укажите путь SMB папки.");
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = sharePath,
                UseShellExecute = true
            });

            AppendLog("Открыта SMB папка: " + sharePath);
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка открытия SMB папки: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
