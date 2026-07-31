using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Threading;
using Connector.Desktop.Features.Connector;
using Connector.Desktop.Features.Tekla.Standard;
using Connector.Desktop.Models;
using Connector.Desktop.Services;
using Forms = System.Windows.Forms;

namespace Connector.Desktop;

public partial class MainWindow : Window, IShellHost, IConnectorHost
{
    private const string FixedServerUrl = "https://server.structura-most.ru";
    private const string FixedUpdateManifestUrl = "https://server.structura-most.ru/updates/latest.json";
    private const int FixedHeartbeatSeconds = 60;
    private const string DefaultSmbSharePath = @"\\62.113.36.107\BIM_Models";
    // Tekla defaults: still used by the shell-owned settings bootstrap/persistence (LoadSettingsToUi, ReadSettingsFromUi,
    // ConnectByTokenInternalAsync). The lifted StandardView keeps its own private copies for its UI; these stay here
    // because the connector's settings layer references them independently of the Стандарт module.
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

    private readonly SettingsService _settingsService = new();
    private readonly AutoStartService _autoStartService = new();
    private readonly HeartbeatClient _heartbeatClient = new(new HttpClient { Timeout = TimeSpan.FromSeconds(110) });
    private readonly UpdateService _updateService = new(new HttpClient { Timeout = TimeSpan.FromSeconds(40) });
    private readonly TeklaStandardService _teklaStandardService = new(new HttpClient { Timeout = TimeSpan.FromSeconds(25) });
    private readonly TcpConnectivityProbe _tcpConnectivityProbe = new();
    // Model Sharing / VPN services now live inside their feature modules (the shell catalog). See ComposeFeatureModules().
    private readonly Shell.ShellViewModel _shell;   // modular feature catalog (domains → modules); assigned in ctor (needs `this` as IShellHost)
    private StandardModule? _standard;   // Tekla domain → "Стандарт" module (the lifted firm/extensions/libraries sync engine)
    private ConnectorModule? _connector;   // "Коннектор" domain module (the lifted login/heartbeat FRONT-END view); engine stays here
    private string _lastVpnConfig = string.Empty;   // last config delivered by bootstrap (kept in memory, not persisted)
    private string _lastSmbLogin = string.Empty;    // SMB creds from bootstrap, for mounting the share over VPN
    private string _lastSmbPassword = string.Empty;
    private readonly DispatcherTimer _timer = new();
    private readonly DispatcherTimer _updateTimer = new();
    private readonly DispatcherTimer _teklaSyncTimer = new();
    private readonly Forms.NotifyIcon _trayIcon;
    private static readonly IReadOnlyList<ReleaseNoteItem> ReleaseNotes = new List<ReleaseNoteItem>
    {
        new()
        {
            Version = "1.0.29",
            PublishedAt = "31.07.2026",
            Title = "VPN включается автоматически",
            Changes = new[]
            {
                "После подключения по токену Connector сам устанавливает и включает VPN; пользователю нужно только один раз подтвердить UAC",
                "Если актуальный VPN уже работает, Connector использует его без переустановки и повторных запросов прав администратора",
                "SMB, heartbeat, обновления, синхронизация и Model Sharing используют защищённый маршрут к серверу по умолчанию",
                "Исправлено восстановление VPN-конфигурации после перезапуска и состояние кнопок управления VPN"
            }
        },
        new()
        {
            Version = "1.0.28",
            PublishedAt = "31.07.2026",
            Title = "Надёжное подключение общей папки",
            Changes = new[]
            {
                "Коннектор быстро определяет блокировку прямого SMB (TCP 445) и не ждёт долгого тайм-аута Windows",
                "Если VPN уже включён, общая папка автоматически подключается через VPN",
                "После первого включения VPN общая папка подключается автоматически"
            }
        },
        new()
        {
            Version = "1.0.27",
            PublishedAt = "31.07.2026",
            Title = "Надёжность общей папки и обновлений",
            Changes = new[]
            {
                "Исправлено подключение к общей папке: Connector больше не отключает другие сетевые диски пользователя при конфликте SMB-учётных данных",
                "Очистка конфликтующего подключения теперь ограничена только сервером общей папки; если Windows удерживает сессию, Connector показывает безопасную точечную инструкцию",
                "Установщик обновления скачивается только по HTTPS и запускается только после успешной проверки SHA-256"
            }
        },
        new()
        {
            Version = "1.0.26",
            PublishedAt = "22.06.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Добавлена вкладка «Общая папка (VPN)»: доступ к общей папке можно открыть через защищённый VPN-канал, если прямое SMB-подключение блокируется сетью",
                "VPN теперь доступен всем подключённым устройствам; конфигурация создаётся автоматически при подключении по токену",
                "Исправлено открытие общей папки через VPN: папка монтируется с выданными сервером SMB-учётными данными",
                "Добавлен раздел Tekla «Патчинг»: коннектор поставляет актуальный патч IFC-экспорта для Tekla 2025 SP7 и может установить его с резервной копией",
                "Интерфейс коннектора переведён на модульную структуру: разделы Tekla, Model Sharing, VPN, Structura и Атрибуты разделены на самостоятельные вкладки",
                "Безопасность: при отзыве доступа устройства его учётная запись общей папки отключается, VPN-peer удаляется, а активные подключения разрываются"
            }
        },
        new()
        {
            Version = "1.0.22",
            PublishedAt = "02.06.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Подключение больше не прерывается, если недоступна общая SMB-папка (например, провайдер закрывает порт 445): коннектор продолжает работать, heartbeat и Model Sharing включаются",
                "Сообщение о неподключённой SMB-папке стало понятнее (возможные причины и что это не влияет на Model Sharing)"
            }
        },
        new()
        {
            Version = "1.0.21",
            PublishedAt = "01.06.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Добавлен раздел Model Sharing: одной кнопкой Connector готовит Tekla на этом компьютере к совместной работе над моделями через сервер фирмы",
                "Настройка выполняется локально, без входа в Trimble; пользователь определяется по токену устройства",
                "Connector сам определяет папку Tekla и сохраняет резервную копию исходного файла; настройку можно повторить после обновления Tekla",
                "Выпуск предназначен для ручной установки и проверки"
            }
        },
        new()
        {
            Version = "1.0.20",
            PublishedAt = "16.04.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Исправлен сценарий обновления Connector поверх установленной предыдущей версии",
                "Устранена проблема, из-за которой после неудачного обновления приложение могло перестать запускаться",
                "Выпуск предназначен для ручной установки и проверки"
            }
        },
        new()
        {
            Version = "1.0.19",
            PublishedAt = "16.04.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Коннектор теперь понятнее показывает, какой именно файл не удалось обновить и какой процесс, вероятнее всего, его блокирует",
                "Сообщения об ошибках синхронизации выводятся не только в журнал, но и в окно приложения",
                "Проверка обновлений Tekla Sync выполняется чаще, чтобы изменения у пользователей подтягивались быстрее"
            }
        },
        new()
        {
            Version = "1.0.18",
            PublishedAt = "16.04.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Исправлена установка встроенного git в Connector",
                "Если встроенный git отсутствует после установки, коннектор теперь сам восстанавливает его из локального пакета",
                "Повышена надежность первой синхронизации на новом рабочем компьютере"
            }
        },
        new()
        {
            Version = "1.0.17",
            PublishedAt = "16.04.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Повышена надежность синхронизации папки фирмы, пользовательских приложений и Grasshopper Libraries",
                "Коннектор корректнее восстанавливает локальные данные синхронизации после сбоев",
                "Улучшена диагностика ошибок при применении обновлений на рабочем компьютере пользователя"
            }
        },
        new()
        {
            Version = "1.0.16",
            PublishedAt = "16.04.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Повышена стабильность синхронизации папки фирмы, пользовательских приложений и Grasshopper Libraries",
                "Коннектор корректнее восстанавливает локальные данные синхронизации и повторно получает обновления с сервера",
                "Улучшена надежность применения обновлений на рабочем компьютере пользователя"
            }
        },
        new()
        {
            Version = "1.0.15",
            PublishedAt = "13.04.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Исправлена синхронизация папки фирмы после изменения структуры файлов в Git",
                "Для папки фирмы сохранен строгий режим: лишние файлы удаляются, нужные файлы обновляются по эталону",
                "Повышена стабильность применения обновлений: корректно обрабатываются файлы и папки с атрибутом ReadOnly"
            }
        },
        new()
        {
            Version = "1.0.14",
            PublishedAt = "11.04.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Исправлена синхронизация папки фирмы после изменения структуры стандартов в Git. Обновление снова корректно приводит локальную папку фирмы к актуальному эталону",
                "Обновлена логика синхронизации папок Tekla. Коннектор теперь в любом случае пытается применить обновления для папки фирмы, пользовательских приложений и Grasshopper Libraries, даже если в этот момент запущены Tekla или Rhino",
                "Если обновление не удалось применить в одной из папок из-за занятых файлов, коннектор останавливает обновление только для этой папки и продолжает проверку остальных разделов",
                "Сообщения о проблемах стали понятнее. Теперь коннектор отдельно показывает, что именно помешало обновлению: запущенная Tekla, запущенный Rhino или занятый файл, открытый другой программой",
                "Ручная синхронизация остается доступной для тех случаев, когда часть файлов не удалось обновить автоматически и их нужно подтянуть после закрытия блокирующей программы"
            }
        },
        new()
        {
            Version = "1.0.13",
            PublishedAt = "11.04.2026",
            Title = "Ключевые изменения версии",
            Changes = new[]
            {
                "Добавлена синхронизация Extensions и Grasshopper Libraries через коннектор. Ранее через коннектор синхронизировалась только папка фирмы, теперь по тому же принципу можно централизованно обновлять и пользовательские приложения Tekla, и общие библиотеки Grasshopper",
                "Принцип синхронизации теперь разделен по типам папок. Папка фирмы приводится в точное соответствие опубликованному стандарту, а для Extensions и Grasshopper Libraries коннектор добавляет и обновляет только управляемые файлы, не удаляя локальные файлы пользователя, которых нет в общем контуре",
                "Синхронизация стала автоматической. Коннектор сам проверяет обновления и сам применяет их без лишних ручных действий. Если обновление не удалось применить из-за занятых файлов, коннектор сообщает об этом понятным текстом и предлагает повторить синхронизацию после освобождения файлов",
                "Раздел Стандарт Tekla переработан. Теперь папка фирмы, пользовательские приложения и Grasshopper Libraries вынесены в отдельные вкладки, а пути для каждой папки можно настраивать отдельно под конкретный компьютер и версию Tekla",
                "Для ответственных за обновление стандарта добавлена единая публикация изменений по трем разделам из одного окна, с последовательным запуском и понятным отображением результата",
                "Добавлена вкладка Structura. В одном месте собраны быстрые переходы к Speckle и Nextcloud, а также окно с доступами, где можно удобно посмотреть и скопировать домен, логин и пароль",
                "Добавлено окно прогресса для длительных операций. Во время синхронизации и публикации теперь видно, что именно делает коннектор и на каком этапе находится процесс",
                "Улучшены статусы и уведомления. Коннектор понятнее показывает, что именно требует обновления, какие действия выполняются автоматически и когда нужно вмешательство пользователя",
                "Уведомление о новой версии теперь показывается заметнее и остается на экране, пока пользователь не закроет его сам",
                "Добавлен раздел Что нового. Теперь ключевые изменения по версиям можно посмотреть прямо в коннекторе"
            }
        }
    };

    private AppSettings _settings = new();
    private bool _isRunning;
    private bool _allowClose;
    private bool _trayHintShown;
    private string _activeSessionId = string.Empty;
    private UpdateManifest? _pendingUpdate;
    private UpdateToastWindow? _updateToastWindow;
    private string? _downloadedInstallerPath;
    private bool _updateOfferShown;
    private string _lastUpdateToastVersion = string.Empty;
    private bool _updateCheckInProgress;
    private bool _teklaCheckInProgress;   // seam backing store for IShellHost.TeklaCheckInProgress (read by connect flow)
    private bool _teklaBalloonShown;      // seam backing store; reset via IShellHost.ResetTeklaPendingBalloon
    private bool _serverConnectionFailed;
    private static readonly TimeSpan UpdateCheckInterval = TimeSpan.FromMinutes(10);
    private static readonly TimeSpan TeklaSyncCheckInterval = TimeSpan.FromMinutes(2);

    public MainWindow()
    {
        InitializeComponent();
        _shell = new Shell.ShellViewModel(this, this);   // `this` satisfies IShellHost + IConnectorHost; field initializers have already run.
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ComposeFeatureModules();   // bind Tekla sub-tabs (TeklaTabs) + host Structura/VPN/Атрибуты module views + wire glue
        _timer.Tick += Timer_Tick;
        _updateTimer.Tick += UpdateTimer_Tick;
        _teklaSyncTimer.Tick += TeklaSyncTimer_Tick;
        Closing += MainWindow_Closing;
        Closed += MainWindow_Closed;
        StateChanged += MainWindow_StateChanged;
        _trayIcon = CreateTrayIcon();
        LoadSettingsToUi();
        UpdateRunStateUi();
        UpdateActionButtonUi();
        _standard?.RefreshUi();
        SyncFeatureModules();
        UpdateHeaderStatusUi();
    }

    // ===== IShellHost (seam for the lifted "Стандарт" module) =========================================
    // Implements the surface the lifted StandardView calls. These map 1:1 onto the shell fields/methods the
    // original MainWindow code touched directly, so the lifted logic stays behaviour-identical.

    // LIVE getter — MainWindow reassigns _settings (ReadSettingsFromUi / ConnectByTokenInternalAsync), never cache.
    AppSettings IShellHost.Settings => _settings;

    void IShellHost.SaveSettings() => _settingsService.Save(_settings);

    HeartbeatClient IShellHost.Heartbeat => _heartbeatClient;

    void IShellHost.Log(string message) => AppendLog(message);

    MessageBoxResult IShellHost.ShowDialog(string message, string title, MessageBoxButton buttons, MessageBoxImage image) =>
        ThemedDialogs.Show(this, message, title, buttons, image);

    Window IShellHost.OwnerWindow => this;

    bool IShellHost.TeklaCheckInProgress
    {
        get => _teklaCheckInProgress;
        set => _teklaCheckInProgress = value;
    }

    // Cross-tab mirror (Коннектор tab) + window header. The module computes the overall status and the action-button
    // presentation and pushes them here at the end of every refresh; the shell applies them to the controls that no
    // longer live in the StandardView (mirrors the old UpdateTeklaUi/UpdateTeklaActionButtonUi/UpdateHeaderStatusUi).
    void IShellHost.OnTeklaStatusChanged(string overallText, System.Windows.Media.Brush overallBrush, string actionButtonContent, bool actionIsSyncStyle, bool inProgress)
    {
        // Cross-tab mirror (ConnectorTeklaSyncStatusTextBlock + ConnectorTeklaSyncButton) moved into ConnectorView;
        // push it through the module. The window header (HeaderFirmStatusTextBlock) stays in the shell.
        _connector?.SetTeklaMirror(overallText, overallBrush, actionButtonContent, actionIsSyncStyle);
        HeaderFirmStatusTextBlock.Text = overallText;
        HeaderFirmStatusTextBlock.Foreground = overallBrush;
    }

    void IShellHost.ShowTrayBalloon(int durationMs, string title, string message, bool isWarning) =>
        _trayIcon.ShowBalloonTip(durationMs, title, message, isWarning ? Forms.ToolTipIcon.Warning : Forms.ToolTipIcon.Info);

    bool IShellHost.IsWindowVisible => IsVisible && WindowState != WindowState.Minimized;

    void IShellHost.ResetTeklaPendingBalloon() => _teklaBalloonShown = false;

    // ===== IConnectorHost (seam for the lifted "Коннектор" front-end module) ===========================
    // Implements the surface the moved ConnectorView handlers call. The connect/heartbeat ENGINE stays here and is
    // unchanged; these map onto the SACRED engine entry points + the shell primitives the moved handlers touch.
    // Explicit implementations forward to the existing (private) engine methods so their signatures stay untouched.

    // The login/heartbeat SPINE — the moved ConnectByToken_Click calls this VERBATIM engine method.
    Task IConnectorHost.ConnectByTokenAsync(string token, bool showSuccessDialog) =>
        ConnectByTokenInternalAsync(token, showSuccessDialog);

    // App self-update "check" path (UpdateAction_Click when no pending update).
    Task IConnectorHost.CheckUpdatesAsync(bool showDialogs) => CheckUpdatesAsync(showDialogs);

    // App self-update "install" path (UpdateAction_Click when an update is pending).
    Task IConnectorHost.InstallPendingUpdateAsync(bool confirmBeforeRun) => InstallPendingUpdateAsync(confirmBeforeRun);

    bool IConnectorHost.HasPendingUpdate => _pendingUpdate is not null;

    // Tekla-mirror sync button — identical to the Стандарт-tab button (OperationProgressWindow + "already running").
    Task IConnectorHost.RunTeklaInteractiveSyncAsync() =>
        _standard is null ? Task.CompletedTask : _standard.RunInteractiveSyncAsync();

    void IConnectorHost.Log(string message) => AppendLog(message);

    MessageBoxResult IConnectorHost.ShowDialog(string message, string title, MessageBoxButton buttons, MessageBoxImage image) =>
        ThemedDialogs.Show(this, message, title, buttons, image);

    Window IConnectorHost.OwnerWindow => this;

    void IConnectorHost.SetServerConnectionFailed(bool failed) => _serverConnectionFailed = failed;

    void IConnectorHost.UpdateHeaderStatus() => UpdateHeaderStatusUi();

    // LIVE getter — MainWindow reassigns _settings (ReadSettingsFromUi / ConnectByTokenInternalAsync), never cache.
    AppSettings IConnectorHost.Settings => _settings;

    IReadOnlyList<ReleaseNoteItem> IConnectorHost.ReleaseNotes => ReleaseNotes;

    // (TeklaUpdateAction_Click moved into ConnectorView — the Коннектор-tab mirror button now lives in the module and
    // routes through IConnectorHost.RunTeklaInteractiveSyncAsync.)

    // ===== Feature-module composition (MVVM migration) =================================================
    // The Model Sharing / Structura / VPN tabs now live in self-contained IFeatureModule views hosted from the
    // shell catalog. ComposeFeatureModules() hosts those views once and wires the shell-owned glue (journal,
    // persistence, themed dialogs, the SMB-mount helper). SyncFeatureModules() re-pushes the current settings
    // snapshot into them — at load and after each token-connect — replacing the old per-tab UpdateXxxUi calls.

    private void ComposeFeatureModules()
    {
        // Capture the Коннектор module + host its view BEFORE LoadSettingsToUi/UpdateRunStateUi/UpdateActionButtonUi
        // (called later in the ctor) push state into it. The connect/heartbeat engine stays here and reaches the
        // moved controls through this module's push methods (replacing the old direct control writes).
        _connector = _shell.Connector.Module("Коннектор") as ConnectorModule;
        ConnectorHost.Content = _connector?.View;

        _standard = _shell.Tekla.Module("Стандарт") as StandardModule;
        // Data-driven Tekla sub-nav: TeklaTabs renders one tab per Tekla module (Стандарт, Model Sharing, Патчинг)
        // from this collection. Adding a Tekla module needs no XAML/host change — only ShellViewModel registration.
        TeklaTabs.ItemsSource = _shell.Tekla.Modules;

        if (_shell.Tekla.Module("Model Sharing") is Features.Tekla.ModelSharing.ModelSharingModule ms)
        {
            ms.Log = AppendLog;
            // Window-owned themed dialogs (mirror VPN) instead of the module's default owner-less MessageBox.
            ms.ConfirmHandler = msg =>
                ThemedDialogs.Show(this, msg, "Model Sharing", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
            ms.ShowMessage = (msg, isError) =>
                ThemedDialogs.Show(this, msg, "Model Sharing", MessageBoxButton.OK,
                    isError ? MessageBoxImage.Warning : MessageBoxImage.Information);
            // Persistence stays in the shell: the module hands back the applied values to save the ModelSharing* keys.
            ms.OnProvisioned = info =>
            {
                _settings.ModelSharingTeklaBin = info.TeklaBin;
                _settings.ModelSharingServerHost = info.ServerHost;
                _settings.ModelSharingServerPort = info.ServerPort;
                _settings.ModelSharingIdentityEmail = info.IdentityEmail;
                _settings.ModelSharingLastAppliedUtc = info.AppliedUtc;
                _settingsService.Save(_settings);
            };
        }

        if (_shell.Tekla.Module("Патчинг") is Features.Tekla.Patching.PatchingModule patch)
        {
            patch.ConfirmHandler = msg =>
                ThemedDialogs.Show(this, msg, "Патчинг Tekla", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK;
        }

        if (_shell.Structura.Module("Structura") is Features.Structura.StructuraModule st)
        {
            StructuraHost.Content = st.View;
            st.Log = AppendLog;
            st.Decrypt = SettingsService.DecryptToken;
        }

        if (_shell.Vpn.Module("Общая папка (VPN)") is Features.Vpn.VpnModule vpn)
        {
            VpnHost.Content = vpn.View;
            vpn.Log = AppendLog;
            vpn.DialogHandler = (msg, button, image) => ThemedDialogs.Show(this, msg, "VPN", button, image);
            // The big SMB-mount helper + in-memory creds stay in the shell; the module only requests "open this UNC".
            vpn.OpenShareHandler = OpenVpnShareAsync;
        }

        // Атрибуты: placeholder domain (no shell glue yet — pure roadmap view).
        AttributesHost.Content = _shell.Attributes.Module("Атрибуты")?.View;
    }

    private void SyncFeatureModules()
    {
        if (_shell.Tekla.Module("Model Sharing") is Features.Tekla.ModelSharing.ModelSharingModule ms)
        {
            ms.Initialize(
                new Features.Tekla.ModelSharing.ModelSharingIdentity(
                    _settings.DeviceId, _settings.IssuedTo, !string.IsNullOrWhiteSpace(_settings.DeviceId)),
                _settings.ModelSharingServerHost,
                _settings.ModelSharingServerPort,
                _settings.ModelSharingTeklaBin,
                _settings.TeklaExtensionsLocalPath);
        }

        if (_shell.Structura.Module("Structura") is Features.Structura.StructuraModule st)
        {
            st.Initialize(
                _settings.StructuraSpeckleUrl, _settings.StructuraSpeckleLogin, _settings.StructuraSpecklePasswordCipherBase64,
                _settings.StructuraNextcloudUrl, _settings.StructuraNextcloudLogin, _settings.StructuraNextcloudPasswordCipherBase64);
        }

        if (_shell.Vpn.Module("Общая папка (VPN)") is Features.Vpn.VpnModule vpn)
        {
            vpn.PushContext(new Features.Vpn.VpnContext(
                _settings.VpnEnabled, _settings.VpnSmbUnc, _settings.VpnServerIp,
                _lastVpnConfig, _lastSmbLogin, _lastSmbPassword));
        }
    }

    // Shell-owned SMB-over-VPN mount, invoked by the VPN module's OpenShareHandler. Ported verbatim from the old
    // VpnOpenFolder_Click: prefer the in-memory SMB creds, fall back to a plain explorer open, warn on failure.
    private async void OpenVpnShareAsync(string unc)
    {
        if (string.IsNullOrWhiteSpace(unc))
        {
            return;
        }
        try
        {
            if (!string.IsNullOrWhiteSpace(_lastSmbLogin) && !string.IsNullOrWhiteSpace(_lastSmbPassword))
            {
                await ConnectSmbInternalAsync(_lastSmbLogin, _lastSmbPassword, unc, openExplorer: true);
                AppendLog("VPN: общая папка подключена (" + unc + ").");
            }
            else
            {
                Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = unc, UseShellExecute = true });
                AppendLog("VPN: открыта папка без явного логина; при запросе введите SMB-логин или переподключитесь по токену.");
            }
        }
        catch (Exception ex)
        {
            AppendLog("VPN open folder error: " + ex.Message);
            try { Process.Start(new ProcessStartInfo { FileName = "explorer.exe", Arguments = unc, UseShellExecute = true }); } catch { }
            ThemedDialogs.Show(this,
                "Не удалось автоматически подключить общую папку через VPN: " + ex.Message +
                "\n\nЕсли откроется окно проводника с запросом — введите SMB-логин и пароль из вкладки «Коннектор».",
                "VPN — общая папка", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private async void Window_Loaded(object sender, RoutedEventArgs e)
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
        // Establish the device session and ensure its VPN before any automatic
        // server-bound update/sync work. On an already provisioned PC the
        // automatic tunnel service is up before Connector starts; on the first
        // run bootstrap supplies the config and Connector asks for UAC once.
        await TryAutoConnectAsync();
        await CheckUpdatesAsync(showDialogs: false);
        if (_standard is not null)
        {
            await _standard.RunTeklaSyncAsync(
                showDialogs: false,
                forceRefresh: false,
                autoApplyIfPossible: true);
        }
        _updateTimer.Interval = UpdateCheckInterval;
        _updateTimer.Start();
        _teklaSyncTimer.Interval = TeklaSyncCheckInterval;
        _teklaSyncTimer.Start();
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
            _connector?.SetUpdateState("Обновление: загрузка установщика...");
            _downloadedInstallerPath = await _updateService.DownloadInstallerAsync(_pendingUpdate, CancellationToken.None);
            AppendLog("Скачан установщик обновления: " + _downloadedInstallerPath);
            UpdateService.RunInstaller(_downloadedInstallerPath);
            ExitFromTray();
        }
        catch (Exception ex)
        {
            AppendLog("Ошибка автообновления: " + ex.Message);
            _connector?.SetUpdateState("Обновление: ошибка установки");
            _updateOfferShown = false;
        }
    }

    private void ShowUpdateAvailableToast(UpdateManifest manifest)
    {
        var version = manifest.Version?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(version))
        {
            return;
        }

        if (string.Equals(_lastUpdateToastVersion, version, StringComparison.OrdinalIgnoreCase))
        {
            if (_updateToastWindow is not null && _updateToastWindow.IsVisible)
            {
                _updateToastWindow.BringToFront();
            }
            return;
        }

        if (_updateToastWindow is not null && _updateToastWindow.IsVisible)
        {
            _updateToastWindow.Close();
            _updateToastWindow = null;
        }

        _lastUpdateToastVersion = version;
        var toast = new UpdateToastWindow(
            "Structura Connector",
            "Доступна новая версия: " + version + ".",
            async () => await InstallPendingUpdateAsync(confirmBeforeRun: false));
        toast.Closed += (_, _) =>
        {
            if (ReferenceEquals(_updateToastWindow, toast))
            {
                _updateToastWindow = null;
            }
        };
        _updateToastWindow = toast;
        toast.Show();
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
        if (_standard is not null)
        {
            await _standard.RunTeklaSyncAsync(showDialogs: false, forceRefresh: false, autoApplyIfPossible: true);
        }
    }

    private async void TeklaSyncTimer_Tick(object? sender, EventArgs e)
    {
        if (_standard is not null)
        {
            await _standard.RunTeklaSyncAsync(showDialogs: false, forceRefresh: false, autoApplyIfPossible: true);
        }
    }

    private void UpdateActionButtonUi()
    {
        if (_pendingUpdate is null)
        {
            _connector?.SetUpdateAction("Проверить обновление коннектора", isPrimaryStyle: false);
            return;
        }

        _connector?.SetUpdateAction("Скачать и установить обновление", isPrimaryStyle: true);
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

        // The Tekla overall status (HeaderFirmStatusTextBlock) now arrives from the Стандарт module via
        // IShellHost.OnTeklaStatusChanged; UpdateHeaderStatusUi only owns the server-status line.
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
            "Найдена ревизия " + revision + ". Закройте Tekla, и Connector применит обновление автоматически на следующей проверке.",
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

        if (string.IsNullOrWhiteSpace(_settings.TeklaExtensionsManifestUrl))
        {
            _settings.TeklaExtensionsManifestUrl = FixedTeklaExtensionsManifestUrl;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaExtensionsLocalPath))
        {
            _settings.TeklaExtensionsLocalPath = DefaultTeklaExtensionsLocalPath;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaPublishSourcePath) ||
            string.Equals(_settings.TeklaPublishSourcePath, @"\\62.113.36.107\BIM_Models\Tekla\XS_FIRM", StringComparison.OrdinalIgnoreCase))
        {
            _settings.TeklaPublishSourcePath = DefaultTeklaPublishSourcePath;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaExtensionsPublishSourcePath) ||
            string.Equals(_settings.TeklaExtensionsPublishSourcePath, @"\\62.113.36.107\BIM_Models\Tekla\Extension", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(_settings.TeklaExtensionsPublishSourcePath, @"\\62.113.36.107\BIM_Models\Tekla\Extensions", StringComparison.OrdinalIgnoreCase))
        {
            _settings.TeklaExtensionsPublishSourcePath = DefaultTeklaExtensionsPublishSourcePath;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaLibrariesManifestUrl))
        {
            _settings.TeklaLibrariesManifestUrl = FixedTeklaLibrariesManifestUrl;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaLibrariesLocalPath))
        {
            _settings.TeklaLibrariesLocalPath = DefaultTeklaLibrariesLocalPath;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.TeklaLibrariesPublishSourcePath))
        {
            _settings.TeklaLibrariesPublishSourcePath = DefaultTeklaLibrariesPublishSourcePath;
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.StructuraSpeckleUrl))
        {
            _settings.StructuraSpeckleUrl = "https://speckle.structura-most.ru";
            shouldPersist = true;
        }

        if (string.IsNullOrWhiteSpace(_settings.StructuraNextcloudUrl))
        {
            _settings.StructuraNextcloudUrl = "https://cloud.structura-most.ru";
            shouldPersist = true;
        }

        // The moved (mostly Collapsed) Коннектор controls + the token PasswordBox now live in ConnectorView; push the
        // live settings into them through the module (replaces the ServerUrl/UpdateManifestUrl/DeviceId/SmbLogin/
        // SmbSharePath/Interval/AutoStart/Token/SmbPassword writes — same values, sourced from the live AppSettings).
        // The Tekla local-path / publish-source TextBoxes live in the Стандарт module's StandardView; the module
        // populates them from _settings in its RefreshUi() (called from the ctor and after each token-connect).
        // restore the VPN config (decrypt) so "Включить VPN" works after a restart without re-connecting
        if (!string.IsNullOrWhiteSpace(_settings.VpnConfigCipherBase64))
        {
            _lastVpnConfig = SettingsService.DecryptToken(_settings.VpnConfigCipherBase64);
        }

        var token = SettingsService.DecryptToken(_settings.TokenCipherBase64);
        _connector?.LoadFromSettings(token);
        SyncFeatureModules();

        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);

        if (shouldPersist)
        {
            _settingsService.Save(_settings);
        }

        AppendLog($"Настройки загружены: {_settingsService.SettingsPath}");
    }

    private AppSettings ReadSettingsFromUi()
    {
        // The token PasswordBox moved into ConnectorView. ApplyAndPersist captures the typed token into
        // _settings.TokenCipherBase64 (via _connector.FlushEditsToSettings) BEFORE calling this, so source the token
        // from _settings here (decrypt+trim) — the empty-token guard still fires exactly as with the old PasswordBox read.
        var token = SettingsService.DecryptToken(_settings.TokenCipherBase64).Trim();
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
        // The Tekla local-path / publish-source TextBoxes moved into the Стандарт module's StandardView. The module
        // keeps _settings in sync with those edits (ApplyAndPersistTeklaPathsOnly + SaveSettings) before any save, so
        // the shell sources these values from _settings (same default fallbacks as the original TextBox-empty branch).
        var teklaLocalPath = string.IsNullOrWhiteSpace(_settings.TeklaStandardLocalPath)
            ? DefaultTeklaStandardLocalPath
            : _settings.TeklaStandardLocalPath;
        var teklaExtensionsManifestUrl = string.IsNullOrWhiteSpace(_settings.TeklaExtensionsManifestUrl)
            ? FixedTeklaExtensionsManifestUrl
            : _settings.TeklaExtensionsManifestUrl;
        var teklaExtensionsLocalPath = string.IsNullOrWhiteSpace(_settings.TeklaExtensionsLocalPath)
            ? DefaultTeklaExtensionsLocalPath
            : _settings.TeklaExtensionsLocalPath;
        var teklaLibrariesManifestUrl = string.IsNullOrWhiteSpace(_settings.TeklaLibrariesManifestUrl)
            ? FixedTeklaLibrariesManifestUrl
            : _settings.TeklaLibrariesManifestUrl;
        var teklaLibrariesLocalPath = string.IsNullOrWhiteSpace(_settings.TeklaLibrariesLocalPath)
            ? DefaultTeklaLibrariesLocalPath
            : _settings.TeklaLibrariesLocalPath;
        var teklaFirmPublishSourcePath = string.IsNullOrWhiteSpace(_settings.TeklaPublishSourcePath)
            ? DefaultTeklaPublishSourcePath
            : _settings.TeklaPublishSourcePath;
        var teklaExtensionsPublishSourcePath = string.IsNullOrWhiteSpace(_settings.TeklaExtensionsPublishSourcePath)
            ? DefaultTeklaExtensionsPublishSourcePath
            : _settings.TeklaExtensionsPublishSourcePath;
        var teklaLibrariesPublishSourcePath = string.IsNullOrWhiteSpace(_settings.TeklaLibrariesPublishSourcePath)
            ? DefaultTeklaLibrariesPublishSourcePath
            : _settings.TeklaLibrariesPublishSourcePath;

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
            TeklaStandardLastTechnicalError = _settings.TeklaStandardLastTechnicalError,
            TeklaStandardRepoUrl = _settings.TeklaStandardRepoUrl,
            TeklaStandardRepoRef = _settings.TeklaStandardRepoRef,
            TeklaStandardRepoSubdir = _settings.TeklaStandardRepoSubdir,
            TeklaPublishSourcePath = teklaFirmPublishSourcePath,
            TeklaExtensionsManifestUrl = teklaExtensionsManifestUrl,
            TeklaExtensionsLocalPath = teklaExtensionsLocalPath,
            TeklaExtensionsInstalledVersion = _settings.TeklaExtensionsInstalledVersion,
            TeklaExtensionsTargetVersion = _settings.TeklaExtensionsTargetVersion,
            TeklaExtensionsInstalledRevision = _settings.TeklaExtensionsInstalledRevision,
            TeklaExtensionsTargetRevision = _settings.TeklaExtensionsTargetRevision,
            TeklaExtensionsLastCheckUtc = _settings.TeklaExtensionsLastCheckUtc,
            TeklaExtensionsLastSuccessUtc = _settings.TeklaExtensionsLastSuccessUtc,
            TeklaExtensionsPendingAfterClose = _settings.TeklaExtensionsPendingAfterClose,
            TeklaExtensionsLastError = _settings.TeklaExtensionsLastError,
            TeklaExtensionsLastTechnicalError = _settings.TeklaExtensionsLastTechnicalError,
            TeklaExtensionsRepoUrl = _settings.TeklaExtensionsRepoUrl,
            TeklaExtensionsRepoRef = _settings.TeklaExtensionsRepoRef,
            TeklaExtensionsRepoSubdir = _settings.TeklaExtensionsRepoSubdir,
            TeklaExtensionsPublishSourcePath = teklaExtensionsPublishSourcePath,
            TeklaLibrariesManifestUrl = teklaLibrariesManifestUrl,
            TeklaLibrariesLocalPath = teklaLibrariesLocalPath,
            TeklaLibrariesInstalledVersion = _settings.TeklaLibrariesInstalledVersion,
            TeklaLibrariesTargetVersion = _settings.TeklaLibrariesTargetVersion,
            TeklaLibrariesInstalledRevision = _settings.TeklaLibrariesInstalledRevision,
            TeklaLibrariesTargetRevision = _settings.TeklaLibrariesTargetRevision,
            TeklaLibrariesLastCheckUtc = _settings.TeklaLibrariesLastCheckUtc,
            TeklaLibrariesLastSuccessUtc = _settings.TeklaLibrariesLastSuccessUtc,
            TeklaLibrariesPendingAfterClose = _settings.TeklaLibrariesPendingAfterClose,
            TeklaLibrariesLastError = _settings.TeklaLibrariesLastError,
            TeklaLibrariesLastTechnicalError = _settings.TeklaLibrariesLastTechnicalError,
            TeklaLibrariesRepoUrl = _settings.TeklaLibrariesRepoUrl,
            TeklaLibrariesRepoRef = _settings.TeklaLibrariesRepoRef,
            TeklaLibrariesRepoSubdir = _settings.TeklaLibrariesRepoSubdir,
            TeklaLibrariesPublishSourcePath = teklaLibrariesPublishSourcePath,
            StructuraSpeckleUrl = _settings.StructuraSpeckleUrl,
            StructuraSpeckleLogin = _settings.StructuraSpeckleLogin,
            StructuraSpecklePasswordCipherBase64 = _settings.StructuraSpecklePasswordCipherBase64,
            StructuraNextcloudUrl = _settings.StructuraNextcloudUrl,
            StructuraNextcloudLogin = _settings.StructuraNextcloudLogin,
            StructuraNextcloudPasswordCipherBase64 = _settings.StructuraNextcloudPasswordCipherBase64,
            IsSystemAdmin = _settings.IsSystemAdmin,
            IsFirmAdmin = _settings.IsFirmAdmin,
            IssuedTo = _settings.IssuedTo,
            ModelSharingTeklaBin = _settings.ModelSharingTeklaBin,
            ModelSharingServerHost = string.IsNullOrWhiteSpace(_settings.ModelSharingServerHost) ? "62.113.36.107" : _settings.ModelSharingServerHost,
            ModelSharingServerPort = _settings.ModelSharingServerPort > 0 ? _settings.ModelSharingServerPort : 9990,
            ModelSharingIdentityEmail = _settings.ModelSharingIdentityEmail,
            ModelSharingLastAppliedUtc = _settings.ModelSharingLastAppliedUtc,
            VpnEnabled = _settings.VpnEnabled,
            VpnTunnelName = _settings.VpnTunnelName,
            VpnAddress = _settings.VpnAddress,
            VpnSmbUnc = _settings.VpnSmbUnc,
            VpnServerIp = _settings.VpnServerIp,
            VpnConfigReceivedUtc = _settings.VpnConfigReceivedUtc,
            VpnConfigCipherBase64 = _settings.VpnConfigCipherBase64
        };
    }

    private void ApplyAndPersist()
    {
        _standard?.FlushPathEdits();   // capture typed-but-unbrowsed Tekla path edits before reading settings
        // Capture the typed token from the moved Коннектор PasswordBox into _settings before ReadSettingsFromUi reads
        // it back (mirrors StandardModule.FlushPathEdits). ReadSettingsFromUi now sources the token from _settings.
        _settings.TokenCipherBase64 = SettingsService.EncryptToken(_connector?.FlushEditsToSettings() ?? string.Empty);
        _settings = ReadSettingsFromUi();
        _settingsService.Save(_settings);
        _autoStartService.SetEnabled(_settings.AutoStart);
        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);
        _standard?.RefreshUi();
        AppendLog("Настройки сохранены.");
    }

    private void UpdateRunStateUi()
    {
        string text;
        System.Windows.Media.Brush brush;
        if (_isRunning)
        {
            text = "Автоотправка heartbeat: включена";
            brush = System.Windows.Media.Brushes.MediumSpringGreen;
        }
        else
        {
            text = "Автоотправка heartbeat: выключена";
            brush = System.Windows.Media.Brushes.Orange;
        }

        // RunStateTextBlock + Start/StopButton moved into ConnectorView; push the computed text/brush/enabled state.
        _connector?.SetRunState(text, brush, !_isRunning, _isRunning);
        UpdateHeaderStatusUi();
    }

    private async Task SendHeartbeatSafeAsync()
    {
        try
        {
            var token = SettingsService.DecryptToken(_settings.TokenCipherBase64);
            var teklaRunning = _teklaStandardService.IsTeklaRunning();
            var teklaState = new TeklaHeartbeatState
            {
                InstalledVersion = _settings.TeklaStandardInstalledVersion,
                TargetVersion = _settings.TeklaStandardTargetVersion,
                InstalledRevision = _settings.TeklaStandardInstalledRevision,
                TargetRevision = _settings.TeklaStandardTargetRevision,
                PendingAfterClose =
                    _settings.TeklaStandardPendingAfterClose ||
                    _settings.TeklaExtensionsPendingAfterClose ||
                    _settings.TeklaLibrariesPendingAfterClose,
                TeklaRunning = teklaRunning,
                LastCheckUtc = _settings.TeklaStandardLastCheckUtc?.UtcDateTime.ToString("o") ?? string.Empty,
                LastSuccessUtc = _settings.TeklaStandardLastSuccessUtc?.UtcDateTime.ToString("o") ?? string.Empty,
                LastError = FirstNonEmpty(
                    _settings.TeklaStandardLastError,
                    _settings.TeklaExtensionsLastError,
                    _settings.TeklaLibrariesLastError)
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
            // ServerUrlTextBox moved into ConnectorView (Collapsed); re-source from _settings (LoadFromSettings keeps it
            // mirrored to _settings.ServerUrl, which is always FixedServerUrl — same value the old Collapsed box held).
            var serverUrl = _settings.ServerUrl.Trim();
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

    // (ConnectByToken_Click moved into ConnectorView — it calls the SACRED engine below via IConnectorHost.ConnectByTokenAsync.)

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
            TeklaStandardLastTechnicalError = _settings.TeklaStandardLastTechnicalError,
            TeklaStandardRepoUrl = _settings.TeklaStandardRepoUrl,
            TeklaStandardRepoRef = _settings.TeklaStandardRepoRef,
            TeklaStandardRepoSubdir = _settings.TeklaStandardRepoSubdir,
            TeklaPublishSourcePath = string.IsNullOrWhiteSpace(_settings.TeklaPublishSourcePath)
                ? DefaultTeklaPublishSourcePath
                : _settings.TeklaPublishSourcePath,
            TeklaExtensionsManifestUrl = string.IsNullOrWhiteSpace(_settings.TeklaExtensionsManifestUrl)
                ? FixedTeklaExtensionsManifestUrl
                : _settings.TeklaExtensionsManifestUrl,
            TeklaExtensionsLocalPath = string.IsNullOrWhiteSpace(_settings.TeklaExtensionsLocalPath)
                ? DefaultTeklaExtensionsLocalPath
                : _settings.TeklaExtensionsLocalPath,
            TeklaExtensionsInstalledVersion = _settings.TeklaExtensionsInstalledVersion,
            TeklaExtensionsTargetVersion = _settings.TeklaExtensionsTargetVersion,
            TeklaExtensionsInstalledRevision = _settings.TeklaExtensionsInstalledRevision,
            TeklaExtensionsTargetRevision = _settings.TeklaExtensionsTargetRevision,
            TeklaExtensionsLastCheckUtc = _settings.TeklaExtensionsLastCheckUtc,
            TeklaExtensionsLastSuccessUtc = _settings.TeklaExtensionsLastSuccessUtc,
            TeklaExtensionsPendingAfterClose = _settings.TeklaExtensionsPendingAfterClose,
            TeklaExtensionsLastError = _settings.TeklaExtensionsLastError,
            TeklaExtensionsLastTechnicalError = _settings.TeklaExtensionsLastTechnicalError,
            TeklaExtensionsRepoUrl = _settings.TeklaExtensionsRepoUrl,
            TeklaExtensionsRepoRef = _settings.TeklaExtensionsRepoRef,
            TeklaExtensionsRepoSubdir = _settings.TeklaExtensionsRepoSubdir,
            TeklaExtensionsPublishSourcePath = string.IsNullOrWhiteSpace(_settings.TeklaExtensionsPublishSourcePath)
                ? DefaultTeklaExtensionsPublishSourcePath
                : _settings.TeklaExtensionsPublishSourcePath,
            TeklaLibrariesManifestUrl = string.IsNullOrWhiteSpace(_settings.TeklaLibrariesManifestUrl)
                ? FixedTeklaLibrariesManifestUrl
                : _settings.TeklaLibrariesManifestUrl,
            TeklaLibrariesLocalPath = string.IsNullOrWhiteSpace(_settings.TeklaLibrariesLocalPath)
                ? DefaultTeklaLibrariesLocalPath
                : _settings.TeklaLibrariesLocalPath,
            TeklaLibrariesInstalledVersion = _settings.TeklaLibrariesInstalledVersion,
            TeklaLibrariesTargetVersion = _settings.TeklaLibrariesTargetVersion,
            TeklaLibrariesInstalledRevision = _settings.TeklaLibrariesInstalledRevision,
            TeklaLibrariesTargetRevision = _settings.TeklaLibrariesTargetRevision,
            TeklaLibrariesLastCheckUtc = _settings.TeklaLibrariesLastCheckUtc,
            TeklaLibrariesLastSuccessUtc = _settings.TeklaLibrariesLastSuccessUtc,
            TeklaLibrariesPendingAfterClose = _settings.TeklaLibrariesPendingAfterClose,
            TeklaLibrariesLastError = _settings.TeklaLibrariesLastError,
            TeklaLibrariesLastTechnicalError = _settings.TeklaLibrariesLastTechnicalError,
            TeklaLibrariesRepoUrl = _settings.TeklaLibrariesRepoUrl,
            TeklaLibrariesRepoRef = _settings.TeklaLibrariesRepoRef,
            TeklaLibrariesRepoSubdir = _settings.TeklaLibrariesRepoSubdir,
            TeklaLibrariesPublishSourcePath = string.IsNullOrWhiteSpace(_settings.TeklaLibrariesPublishSourcePath)
                ? DefaultTeklaLibrariesPublishSourcePath
                : _settings.TeklaLibrariesPublishSourcePath,
            StructuraSpeckleUrl = string.IsNullOrWhiteSpace(bootstrap.WebAccess.Speckle.Url)
                ? (string.IsNullOrWhiteSpace(_settings.StructuraSpeckleUrl) ? "https://speckle.structura-most.ru" : _settings.StructuraSpeckleUrl)
                : bootstrap.WebAccess.Speckle.Url,
            StructuraSpeckleLogin = bootstrap.WebAccess.Speckle.Login,
            StructuraSpecklePasswordCipherBase64 = string.IsNullOrWhiteSpace(bootstrap.WebAccess.Speckle.Password)
                ? string.Empty
                : SettingsService.EncryptToken(bootstrap.WebAccess.Speckle.Password),
            StructuraNextcloudUrl = string.IsNullOrWhiteSpace(bootstrap.WebAccess.Nextcloud.Url)
                ? (string.IsNullOrWhiteSpace(_settings.StructuraNextcloudUrl) ? "https://cloud.structura-most.ru" : _settings.StructuraNextcloudUrl)
                : bootstrap.WebAccess.Nextcloud.Url,
            StructuraNextcloudLogin = bootstrap.WebAccess.Nextcloud.Login,
            StructuraNextcloudPasswordCipherBase64 = string.IsNullOrWhiteSpace(bootstrap.WebAccess.Nextcloud.Password)
                ? string.Empty
                : SettingsService.EncryptToken(bootstrap.WebAccess.Nextcloud.Password),
            IsSystemAdmin = bootstrap.IsSystemAdmin,
            IsFirmAdmin = bootstrap.IsFirmAdmin,
            IssuedTo = string.IsNullOrWhiteSpace(bootstrap.IssuedTo) ? _settings.IssuedTo : bootstrap.IssuedTo,
            ModelSharingTeklaBin = _settings.ModelSharingTeklaBin,
            ModelSharingServerHost = string.IsNullOrWhiteSpace(_settings.ModelSharingServerHost) ? "62.113.36.107" : _settings.ModelSharingServerHost,
            ModelSharingServerPort = _settings.ModelSharingServerPort > 0 ? _settings.ModelSharingServerPort : 9990,
            ModelSharingIdentityEmail = _settings.ModelSharingIdentityEmail,
            ModelSharingLastAppliedUtc = _settings.ModelSharingLastAppliedUtc,
            // Preserve the last known-good tunnel until bootstrap supplies a
            // replacement. A transient VPN provisioning error must not disable
            // the controls or strand a working automatic tunnel.
            VpnEnabled = _settings.VpnEnabled,
            VpnTunnelName = _settings.VpnTunnelName,
            VpnAddress = _settings.VpnAddress,
            VpnSmbUnc = _settings.VpnSmbUnc,
            VpnServerIp = _settings.VpnServerIp,
            VpnConfigReceivedUtc = _settings.VpnConfigReceivedUtc,
            VpnConfigCipherBase64 = _settings.VpnConfigCipherBase64
        };
        _settingsService.Save(_settings);
        _autoStartService.SetEnabled(true);
        _timer.Interval = TimeSpan.FromSeconds(_settings.HeartbeatSeconds);
        _activeSessionId = bootstrap.SessionId;
        _serverConnectionFailed = false;

        // The six post-connect Коннектор controls moved into ConnectorView; push the same values across the seam
        // (DeviceId/Interval/UpdateManifestUrl/SmbLogin + SmbPassword cleared + SmbSharePath — exact original values).
        _connector?.ApplyConnectResult(_settings.DeviceId, _settings.HeartbeatSeconds, _settings.UpdateManifestUrl, bootstrap.SmbAccess.Login, _settings.SmbSharePath);
        // SMB creds (kept in memory) so we can mount the share over VPN with the right login.
        _lastSmbLogin = bootstrap.SmbAccess.Login;
        _lastSmbPassword = bootstrap.SmbAccess.Password;

        // VPN bundle from bootstrap (optional; gated by server). Config is persisted ENCRYPTED
        // (DPAPI) so "Включить VPN" works after a restart; kept in memory for immediate use.
        var bootstrapVpnAvailable =
            bootstrap.Vpn.Enabled &&
            !string.IsNullOrWhiteSpace(bootstrap.Vpn.Config);
        if (bootstrapVpnAvailable)
        {
            _settings.VpnEnabled = true;
            _lastVpnConfig = bootstrap.Vpn.Config;
            _settings.VpnTunnelName = bootstrap.Vpn.TunnelName;   // server-side tunnel name (informational)
            _settings.VpnAddress = bootstrap.Vpn.Address;
            _settings.VpnSmbUnc = bootstrap.Vpn.SmbUnc;
            _settings.VpnServerIp = bootstrap.Vpn.ServerVpnIp;
            _settings.VpnConfigReceivedUtc = DateTimeOffset.UtcNow;
            _settings.VpnConfigCipherBase64 = SettingsService.EncryptToken(bootstrap.Vpn.Config);
            _settingsService.Save(_settings);
            AppendLog("Получена конфигурация VPN для доступа к общей папке.");
        }
        else if (_settings.VpnEnabled && !string.IsNullOrWhiteSpace(_lastVpnConfig))
        {
            AppendLog(
                "Сервер временно не обновил VPN-конфигурацию; используется последняя сохранённая рабочая версия.");
        }

        _standard?.RefreshUi();
        SyncFeatureModules();
        AppendLog("Настройки сохранены.");

        var vpnReady = false;
        if (_settings.VpnEnabled &&
            _shell.Vpn.Module("Общая папка (VPN)") is Features.Vpn.VpnModule vpnModule)
        {
            AppendLog("VPN включён по умолчанию. Проверяю автоматический туннель...");
            var vpnResult = await vpnModule.EnsureEnabledAsync(
                showResultDialog: false,
                openShareOnSuccess: false);
            vpnReady = vpnResult.IsSuccess;
            if (vpnReady)
            {
                AppendLog("VPN готов. Дальнейшее подключение к серверу идёт через защищённый туннель.");
            }
            else
            {
                AppendLog(
                    "Автоматическое включение VPN не выполнено: " + vpnResult.Message +
                    " Connector продолжит работу и попробует прямое подключение.");
            }
        }

        var smbConnected = false;
        var smbConnectionRoute = string.Empty;
        var preferredSharePath = vpnReady ? ResolveVpnSmbUnc() : sharePath;
        if (string.IsNullOrWhiteSpace(preferredSharePath))
        {
            preferredSharePath = sharePath;
        }
        var preferredSmbHost = GetSmbHost(preferredSharePath);
        var preferredSmbReachable = await _tcpConnectivityProbe.CanConnectAsync(
            preferredSmbHost,
            445,
            TimeSpan.FromSeconds(4));

        if (preferredSmbReachable)
        {
            try
            {
                await ConnectSmbInternalAsync(
                    bootstrap.SmbAccess.Login,
                    bootstrap.SmbAccess.Password,
                    preferredSharePath,
                    openExplorer: true);
                smbConnected = true;
                smbConnectionRoute = vpnReady ? "через VPN" : "напрямую";
                if (vpnReady)
                {
                    AppendLog("Общая папка автоматически подключена через VPN.");
                }
            }
            catch (Exception ex) when (IsWindowsSmbConflict(ex))
            {
                AppendLog("SMB-подключение не переключено автоматически (конфликт 1219). Текущая сессия SMB оставлена без изменений.");
                AppendLog("Детали SMB конфликта: " + ex.Message);
            }
            catch (Exception ex)
            {
                AppendLog("SMB-подключение не выполнено: " + ex.Message);
            }
        }
        else
        {
            AppendLog(
                $"SMB недоступен: сервер {preferredSmbHost}:445 не отвечает " +
                (vpnReady ? "через VPN." : "из этой сети."));
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
                    ? $"Подключение выполнено. Общая SMB-папка подключена {smbConnectionRoute}, автоотправка heartbeat включена."
                    : "Подключение к серверу выполнено, heartbeat включён, но общая папка пока не подключена. " +
                      "Проверьте статус VPN на вкладке «Общая папка (VPN)» и повторите подключение.",
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
        _connector?.SetUpdateActionEnabled(false);
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
            _connector?.SetUpdateManifestUrl(manifestUrl);
            _settingsService.Save(_settings);

            var manifest = await _updateService.TryGetUpdateAsync(manifestUrl, CancellationToken.None);
            if (manifest is null)
            {
                _pendingUpdate = null;
                _lastUpdateToastVersion = string.Empty;
                _connector?.SetUpdateState("Обновление: не удалось получить данные");
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
                _connector?.SetUpdateState($"Доступно обновление: {manifest.Version}");
                AppendLog("Найдено обновление: " + manifest.Version);
                ShowUpdateAvailableToast(manifest);
                if (!showDialogs)
                {
                    UpdateActionButtonUi();
                }
                else
                {
                    UpdateActionButtonUi();
                }
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
                _lastUpdateToastVersion = string.Empty;
                _connector?.SetUpdateState($"Обновление: актуально ({_updateService.CurrentVersion})");
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
            _connector?.SetUpdateState("Обновление: ошибка проверки");
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
            _connector?.SetUpdateActionEnabled(true);
        }
    }

    // (UpdateAction_Click + ShowReleaseNotes_Click moved into ConnectorView — they route through IConnectorHost
    // [HasPendingUpdate/CheckUpdatesAsync/InstallPendingUpdateAsync] and IConnectorHost.ReleaseNotes/OwnerWindow.)

    private async Task InstallPendingUpdateAsync(bool confirmBeforeRun)
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

            _connector?.SetUpdateActionEnabled(false);
            _connector?.SetUpdateState("Обновление: загрузка установщика...");
            _downloadedInstallerPath = await _updateService.DownloadInstallerAsync(_pendingUpdate, CancellationToken.None);
            AppendLog("Скачан установщик обновления: " + _downloadedInstallerPath);

            var shouldRunInstaller = !confirmBeforeRun || ThemedDialogs.Show(this,
                "Установщик скачан. Закрыть приложение и запустить обновление сейчас?",
                "Обновление",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question) == MessageBoxResult.Yes;

            if (shouldRunInstaller)
            {
                UpdateService.RunInstaller(_downloadedInstallerPath);
                ExitFromTray();
            }
            else
            {
                _connector?.SetUpdateState("Обновление: установщик скачан");
                _connector?.SetUpdateActionEnabled(true);
            }
        }
        catch (Exception ex)
        {
            _connector?.SetUpdateActionEnabled(_pendingUpdate is not null);
            _connector?.SetUpdateState("Обновление: ошибка установки");
            AppendLog("Ошибка установки обновления: " + ex.Message);
            ThemedDialogs.Show(this, ex.Message, "Ошибка обновления", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

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

    private static string GetTeklaTargetDisplayName(string target)
    {
        return target switch
        {
            "extensions" => "Пользовательские приложения",
            "libraries" => "Grasshopper Libraries",
            _ => "Папка фирмы"
        };
    }

    // Public so it satisfies IShellHost (the lifted RestartTeklaServer_Click forwards here) and stays callable
    // by the shell's own Revit-server restart flow. Shared with Revit, so it lives in the shell, not the module.
    public async Task RestartManagedServerAsync(string serviceKey, System.Windows.Controls.Button button, string displayName)
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
            button.IsEnabled = _settings.IsSystemAdmin || _settings.IsFirmAdmin;
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

        try
        {
            RunProcessOrThrow("net", "use", $@"\\{host}\IPC$", "/delete", "/y");
        }
        catch (InvalidOperationException ex) when (IsWindowsNetConnectionNotFound(ex.Message) || IsWindowsNetNoEntries(ex.Message))
        {
            // IPC session not present.
        }
    }

    private string ResolveVpnSmbUnc()
    {
        if (!string.IsNullOrWhiteSpace(_settings.VpnSmbUnc))
        {
            return _settings.VpnSmbUnc.Trim();
        }

        return string.IsNullOrWhiteSpace(_settings.VpnServerIp)
            ? string.Empty
            : $@"\\{_settings.VpnServerIp.Trim()}\BIM_Models";
    }

    private static void DeleteStoredWindowsCredentialForHost(string host)
    {
        var targets = new[]
        {
            host,
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
        // SmbLoginTextBox moved into ConnectorView; push the chosen login candidate across the seam (same value).
        _connector?.SetSmbLogin(loginCandidates.FirstOrDefault() ?? login);

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
                    throw new InvalidOperationException(
                        $"Windows сохранила активное SMB-подключение к серверу {host} с другими учётными данными. " +
                        "Коннектор не стал отключать остальные сетевые диски. " +
                        $"Закройте окна папок этого сервера и удалите только его подключения командой " +
                        $"'net use \\\\{host}\\* /delete /y', затем повторите подключение.",
                        retryEx);
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

            // SmbLogin/SmbPassword/SmbSharePath boxes moved into ConnectorView (Collapsed, button unreachable).
            // Re-source from the in-memory bootstrap creds + _settings (same values the Collapsed boxes carried).
            var login = (_lastSmbLogin ?? string.Empty).Trim();
            var password = (_lastSmbPassword ?? string.Empty).Trim();
            var sharePath = (_settings.SmbSharePath ?? string.Empty).Trim();
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
            // SmbSharePathTextBox moved into ConnectorView (Collapsed, button unreachable); re-source from _settings.
            var sharePath = (_settings.SmbSharePath ?? string.Empty).Trim();
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
