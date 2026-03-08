using System.ComponentModel;
using System.Diagnostics;
using System.Net.Http;
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

    private readonly SettingsService _settingsService = new();
    private readonly AutoStartService _autoStartService = new();
    private readonly HeartbeatClient _heartbeatClient = new(new HttpClient { Timeout = TimeSpan.FromSeconds(110) });
    private readonly UpdateService _updateService = new(new HttpClient { Timeout = TimeSpan.FromSeconds(40) });
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
    private bool _updateCheckInProgress;
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
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        Topmost = true;
        Activate();
        Focus();
        Topmost = false;
        AppendLog("При закрытии окно сворачивается в трей. Для полного выхода: иконка в трее -> Закрыть.");
        _ = TryAutoConnectAsync();
        _ = CheckUpdatesAsync(showDialogs: false);
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

        var result = System.Windows.MessageBox.Show(
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
        }
        catch (Exception ex)
        {
            AppendLog("Автоподключение не выполнено: " + ex.Message);
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
        var sanitizedSecrets = false;

        if (!string.IsNullOrWhiteSpace(_settings.SmbLogin))
        {
            _settings.SmbLogin = string.Empty;
            sanitizedSecrets = true;
        }

        if (!string.IsNullOrWhiteSpace(_settings.SmbPasswordCipherBase64))
        {
            _settings.SmbPasswordCipherBase64 = string.Empty;
            sanitizedSecrets = true;
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
        }

        ServerUrlTextBox.Text = _settings.ServerUrl;
        UpdateManifestUrlTextBox.Text = _settings.UpdateManifestUrl;
        DeviceIdTextBox.Text = _settings.DeviceId;
        SmbLoginTextBox.Text = string.Empty;
        SmbSharePathTextBox.Text = _settings.SmbSharePath;
        IntervalTextBox.Text = _settings.HeartbeatSeconds.ToString();
        AutoStartCheckBox.IsChecked = true;

        var token = SettingsService.DecryptToken(_settings.TokenCipherBase64);
        TokenPasswordBox.Password = token;

        SmbPasswordBox.Password = string.Empty;

        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);

        if (sanitizedSecrets)
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
            AutoStart = true
        };
    }

    private void ApplyAndPersist()
    {
        _settings = ReadSettingsFromUi();
        _settingsService.Save(_settings);
        _autoStartService.SetEnabled(_settings.AutoStart);
        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);
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
    }

    private async Task SendHeartbeatSafeAsync()
    {
        try
        {
            var token = SettingsService.DecryptToken(_settings.TokenCipherBase64);
            await _heartbeatClient.SendHeartbeatAsync(_settings.ServerUrl, _settings.DeviceId, token, _activeSessionId, CancellationToken.None);
            AppendLog("Heartbeat отправлен успешно.");
        }
        catch (Exception ex)
        {
            if (ex.Message.Contains("HTTP 409", StringComparison.OrdinalIgnoreCase))
            {
                _timer.Stop();
                _isRunning = false;
                UpdateRunStateUi();
                AppendLog("Сессия отключена: этот токен активирован на другом устройстве.");
                return;
            }

            AppendLog("Ошибка heartbeat: " + ex.Message);
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
            System.Windows.MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
            System.Windows.MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
                System.Windows.MessageBox.Show(
                    "Сервер доступен и отвечает /health.\nВнешний IP: " + ip,
                    "Проверка подключения",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ipEx)
            {
                AppendLog("Сервер доступен, но внешний IP определить не удалось: " + ipEx.Message);
                System.Windows.MessageBox.Show(
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
            System.Windows.MessageBox.Show(ex.Message, "Ошибка проверки подключения", MessageBoxButton.OK, MessageBoxImage.Error);
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
            System.Windows.MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
            System.Windows.MessageBox.Show(message, "Время ожидания истекло", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка автоподключения по токену: " + ex.Message);
            System.Windows.MessageBox.Show(ex.Message, "Ошибка подключения", MessageBoxButton.OK, MessageBoxImage.Error);
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
        };
        _settingsService.Save(_settings);
        _autoStartService.SetEnabled(true);
        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);
        _activeSessionId = bootstrap.SessionId;

        DeviceIdTextBox.Text = _settings.DeviceId;
        IntervalTextBox.Text = _settings.HeartbeatSeconds.ToString();
        UpdateManifestUrlTextBox.Text = _settings.UpdateManifestUrl;
        SmbLoginTextBox.Text = bootstrap.SmbAccess.Login;
        SmbPasswordBox.Password = string.Empty;
        SmbSharePathTextBox.Text = _settings.SmbSharePath;
        AppendLog("Настройки сохранены.");

        await ConnectSmbInternalAsync(bootstrap.SmbAccess.Login, bootstrap.SmbAccess.Password, sharePath, openExplorer: true);

        _timer.Stop();
        _timer.Start();
        _isRunning = true;
        UpdateRunStateUi();

        await SendHeartbeatSafeAsync();
        AppendLog("Автоподключение по токену выполнено успешно.");

        if (showSuccessDialog)
        {
            System.Windows.MessageBox.Show(
                "Подключение выполнено. SMB доступ открыт, автоотправка heartbeat включена.",
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
                throw new InvalidOperationException("Введите корректный URL manifest обновлений.");
            }

            _settings.UpdateManifestUrl = manifestUrl;
            UpdateManifestUrlTextBox.Text = manifestUrl;
            _settingsService.Save(_settings);

            var manifest = await _updateService.TryGetUpdateAsync(manifestUrl, CancellationToken.None);
            if (manifest is null)
            {
                _pendingUpdate = null;
                UpdateStateTextBlock.Text = "Обновление: не удалось получить manifest";
                UpdateActionButtonUi();
                if (showDialogs)
                {
                    System.Windows.MessageBox.Show("Не удалось проверить обновления.", "Обновления", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                return;
            }

            if (_updateService.IsUpdateAvailable(manifest))
            {
                _pendingUpdate = manifest;
                UpdateStateTextBlock.Text = $"Доступно обновление: {manifest.Version}";
                AppendLog("Найдено обновление: " + manifest.Version);
                UpdateActionButtonUi();
                if (showDialogs)
                {
                    System.Windows.MessageBox.Show(
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
                UpdateStateTextBlock.Text = $"Обновление: актуально ({_updateService.CurrentVersion})";
                UpdateActionButtonUi();
                if (showDialogs)
                {
                    System.Windows.MessageBox.Show("Установлена актуальная версия.", "Обновления", MessageBoxButton.OK, MessageBoxImage.Information);
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
                System.Windows.MessageBox.Show(ex.Message, "Ошибка обновлений", MessageBoxButton.OK, MessageBoxImage.Error);
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

            var result = System.Windows.MessageBox.Show(
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
            System.Windows.MessageBox.Show(ex.Message, "Ошибка обновления", MessageBoxButton.OK, MessageBoxImage.Error);
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

        if (login.Contains('\\') || login.Contains('@'))
        {
            return login;
        }

        return $"{host}\\{login}";
    }

    private static void RunProcessOrThrow(string fileName, params string[] args)
    {
        var psi = new ProcessStartInfo(fileName)
        {
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        foreach (var arg in args)
        {
            psi.ArgumentList.Add(arg);
        }

        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Не удалось запустить процесс.");
        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (process.ExitCode != 0)
        {
            var details = string.IsNullOrWhiteSpace(error) ? output : error;
            throw new InvalidOperationException(details.Trim());
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
        var normalizedLogin = NormalizeSmbLogin(login, host);
        SmbLoginTextBox.Text = normalizedLogin;

        await Task.Run(() =>
        {
            try
            {
                RunProcessOrThrow("net", "use", shareRoot, "/delete", "/y");
            }
            catch
            {
                // Ignore cleanup errors for non-existing mappings.
            }

            RunProcessOrThrow("net", "use", shareRoot, password, $"/user:{normalizedLogin}", "/persistent:no");
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
            System.Windows.MessageBox.Show(ex.Message, "Ошибка SMB входа", MessageBoxButton.OK, MessageBoxImage.Error);
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
            System.Windows.MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
