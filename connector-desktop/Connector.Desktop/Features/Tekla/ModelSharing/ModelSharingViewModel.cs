using System.Diagnostics;
using System.IO;
using System.Windows.Input;
using System.Windows.Media;
using Connector.Desktop.Mvvm;
using Connector.Desktop.Services;

namespace Connector.Desktop.Features.Tekla.ModelSharing;

// ViewModel for the Tekla "Model Sharing" feature module. Drives ModelSharingProvisioningService: resolve the
// Tekla bin, surface provisioning status, and run the one-shot per-PC provisioning (IL-patch + redirect config).
//
// Cross-cutting state stays out of the VM: identity (who this device is), server endpoint, the persisted bin and
// the Extensions fallback, logging, persistence and themed dialogs/progress are all pushed IN by the shell via
// Initialize(...) + the *Handler callbacks. The VM never references a global AppSettings/SettingsService.
public sealed class ModelSharingViewModel : ObservableObject
{
    private readonly ModelSharingProvisioningService _service;

    // ---- shell-supplied glue (set once via Initialize, before the first RefreshStatus) ----
    private ModelSharingIdentity _identity = new("", "", false);
    private string _serverHost = "62.113.36.107";
    private int _serverPort = 9990;
    private string _persistedTeklaBin = "";
    private string _extensionsLocalPath = "";

    // The View sets these to show a confirm dialog / message box (the shell may swap in window-owned themed
    // dialogs). ConfirmHandler returns true to proceed; ShowMessage(text, isError) is informational.
    public Func<string, bool>? ConfirmHandler { get; set; }
    public Action<string, bool>? ShowMessage { get; set; }

    // Shell wires this to AppendLog (journal in MainWindow). Best-effort; null is fine in design-time.
    public Action<string>? Log { get; set; }

    // Shell wires this to persist the ModelSharing* settings keys + Save() — persistence stays in the shell.
    public Action<ModelSharingProvisionedInfo>? OnProvisioned { get; set; }

    // The View wires these (OperationProgressWindow is a Window-owned shell helper, so it stays in the View):
    // BeginProgress shows the window, ReportProgress/Complete drive it. All optional — null => no progress UI.
    public Func<ModelSharingProgressHandle?>? BeginProgress { get; set; }

    public ModelSharingViewModel(ModelSharingProvisioningService? service = null)
    {
        _service = service ?? new ModelSharingProvisioningService();

        BrowseCommand = new RelayCommand(() => { /* folder pick is supplied by the View */ });
        SetupCommand = new AsyncRelayCommand(SetupAsync, () => !IsBusy && SetupEnabled);
        OpenLogCommand = new RelayCommand(OpenLog);

        _teklaBin = "";
    }

    private string _teklaBin = "";
    public string TeklaBin { get => _teklaBin; set { if (SetProperty(ref _teklaBin, value)) RefreshStatus(); } }

    private string _statusLine = "Статус: проверка ещё не выполнялась";
    public string StatusLine { get => _statusLine; private set => SetProperty(ref _statusLine, value); }

    // Status colour, mirroring the legacy Orange / MediumSpringGreen / Gainsboro brushes set in code-behind.
    private System.Windows.Media.Brush _statusBrush = System.Windows.Media.Brushes.Gainsboro;
    public System.Windows.Media.Brush StatusBrush { get => _statusBrush; private set => SetProperty(ref _statusBrush, value); }

    private string _identityLine = "";
    public string IdentityLine { get => _identityLine; private set => SetProperty(ref _identityLine, value); }

    private bool _setupEnabled;
    public bool SetupEnabled { get => _setupEnabled; private set { if (SetProperty(ref _setupEnabled, value)) RaiseCanExec(); } }

    private bool _isBusy;
    public bool IsBusy { get => _isBusy; private set { if (SetProperty(ref _isBusy, value)) RaiseCanExec(); } }

    private string _resultMessage = "";
    public string ResultMessage { get => _resultMessage; private set => SetProperty(ref _resultMessage, value); }

    public ICommand BrowseCommand { get; }
    public ICommand SetupCommand { get; }
    public ICommand OpenLogCommand { get; }

    // Shell pushes identity + endpoint + persisted/fallback paths in once, then asks for a status refresh.
    public void Initialize(ModelSharingIdentity identity, string serverHost, int serverPort,
        string persistedTeklaBin, string extensionsLocalPath)
    {
        _identity = identity;
        _serverHost = string.IsNullOrWhiteSpace(serverHost) ? "62.113.36.107" : serverHost;
        _serverPort = serverPort > 0 ? serverPort : 9990;
        _persistedTeklaBin = persistedTeklaBin ?? "";
        _extensionsLocalPath = extensionsLocalPath ?? "";
        RefreshStatus();
    }

    // Ported from MainWindow.UpdateModelSharingUi: resolve the bin, fill the text box if empty, compose the
    // identity line, then read status and map it to text + brush + setup-enabled.
    public void RefreshStatus()
    {
        try
        {
            var teklaBin = ResolveTeklaBin();
            if (string.IsNullOrWhiteSpace(TeklaBin))
            {
                // Seed the bound text box without re-triggering RefreshStatus through the public setter.
                _teklaBin = teklaBin;
                OnPropertyChanged(nameof(TeklaBin));
            }

            var (email, name) = ResolveIdentity();
            var hasDevice = _identity.HasDevice;
            IdentityLine = hasDevice
                ? $"Пользователь: {name} ({email}). Сервер: {_serverHost}:{_serverPort}."
                : "Пользователь будет определён после подключения по токену устройства на вкладке \"Коннектор\".";

            var status = _service.GetStatus(teklaBin);
            string text;
            System.Windows.Media.Brush brush;
            if (!status.FeatureDllExists)
            {
                text = "Статус: SharingUIFeature.dll не найден в указанной папке bin Tekla.";
                brush = System.Windows.Media.Brushes.Orange;
            }
            else if (status.Provisioned)
            {
                var applied = status.AppliedUtc?.ToLocalTime().ToString("dd.MM.yyyy HH:mm") ?? "-";
                text = $"Статус: настроено ({status.IdentityEmail}, сервер {status.ServerHost}:{status.ServerPort}). Применено: {applied}.";
                brush = System.Windows.Media.Brushes.MediumSpringGreen;
            }
            else if (status.NeedsReapply)
            {
                text = "Статус: ранее настраивалось, но файл заменён обновлением Tekla — запустите настройку повторно.";
                brush = System.Windows.Media.Brushes.Orange;
            }
            else
            {
                text = "Статус: не настроено на этом компьютере.";
                brush = System.Windows.Media.Brushes.Gainsboro;
            }
            StatusLine = text;
            StatusBrush = brush;
            SetupEnabled = hasDevice && status.FeatureDllExists;
        }
        catch (Exception ex)
        {
            StatusLine = "Ошибка проверки статуса: " + ex.Message;
            StatusBrush = System.Windows.Media.Brushes.Orange;
            SetupEnabled = false;
        }
    }

    // Ported from MainWindow.ModelSharingSetup_Click: guard rails (device, dll, Tekla running), confirm, run
    // provisioning off-thread behind the View-supplied progress window, then hand persistence to the shell.
    private async Task SetupAsync()
    {
        // Created only after the confirm passes (matches the original lifecycle); the early guards below must not
        // touch a phantom window. The IsTeklaRunning / confirm-cancel branches Dismiss() a still-null handle (no-op).
        ModelSharingProgressHandle? progress = null;
        try
        {
            if (!_identity.HasDevice)
            {
                throw new InvalidOperationException("Сначала подключитесь по токену устройства на вкладке \"Коннектор\".");
            }

            var teklaBin = ResolveTeklaBin();
            var dll = Path.Combine(teklaBin, "Features", "SharingUIFeature.dll");
            if (!File.Exists(dll))
            {
                throw new InvalidOperationException("Не найден файл " + dll + ". Укажите правильную папку bin Tekla.");
            }

            if (_service.IsTeklaRunning())
            {
                progress?.Dismiss();
                ShowMessage?.Invoke("Сейчас запущена Tekla Structures. Закройте её и повторите настройку Model Sharing.", true);
                return;
            }

            var (email, name) = ResolveIdentity();
            var serverHost = string.IsNullOrWhiteSpace(_serverHost) ? "62.113.36.107" : _serverHost;
            var serverPort = _serverPort > 0 ? _serverPort : 9990;

            var confirm = ConfirmHandler?.Invoke(
                $"Подготовить Tekla к Model Sharing на этом компьютере?\n\nПользователь: {name} ({email})\nСервер: {serverHost}:{serverPort}\nПапка: {teklaBin}\n\nTekla должна быть закрыта.") ?? true;
            if (!confirm)
            {
                progress?.Dismiss();
                return;
            }

            var request = new ModelSharingProvisionRequest
            {
                TeklaBin = teklaBin,
                ServerHost = serverHost,
                ServerPort = serverPort,
                IdentityEmail = email,
                IdentityName = name
            };

            progress = BeginProgress?.Invoke();
            progress?.Show();
            progress?.Step("Настраиваем Tekla", "Патчим клиент и записываем адрес сервера", 1, 2, TimeSpan.FromSeconds(8));
            IsBusy = true;

            var result = await Task.Run(() => _service.Provision(request));

            ResultMessage = result.Message;

            if (result.IsSuccess)
            {
                // Persistence stays in the shell: hand back the applied values for it to save the ModelSharing* keys.
                OnProvisioned?.Invoke(new ModelSharingProvisionedInfo(teklaBin, serverHost, serverPort, email, DateTimeOffset.UtcNow));
                progress?.Step("Готово", "Tekla настроена", 2, 2, TimeSpan.Zero);
                progress?.Succeeded(result.Message);
                Log?.Invoke("Model Sharing настроен: " + email + " -> " + serverHost + ":" + serverPort);
                ShowMessage?.Invoke(result.Message, false);
            }
            else
            {
                progress?.Failed(result.Message);
                Log?.Invoke("Model Sharing не настроен: " + result.Message +
                            (string.IsNullOrWhiteSpace(result.TechnicalDetails) ? string.Empty : " | " + result.TechnicalDetails));
                ShowMessage?.Invoke(result.Message, true);
            }
        }
        catch (Exception ex)
        {
            progress?.Failed(ex.Message);
            ResultMessage = ex.Message;
            Log?.Invoke("Ошибка настройки Model Sharing: " + ex.Message);
            ShowMessage?.Invoke(ex.Message, true);
        }
        finally
        {
            IsBusy = false;
            RefreshStatus();
        }
    }

    // Ported from MainWindow.ModelSharingOpenLog_Click: ensure the log exists then open it best-effort.
    private void OpenLog()
    {
        try
        {
            if (!File.Exists(_service.LogFilePath))
            {
                File.WriteAllText(_service.LogFilePath, string.Empty);
            }
            Process.Start(new ProcessStartInfo { FileName = _service.LogFilePath, UseShellExecute = true });
        }
        catch (Exception ex)
        {
            Log?.Invoke("Не удалось открыть лог Model Sharing: " + ex.Message);
        }
    }

    // Ported from MainWindow.ResolveModelSharingTeklaBin: prefer the bound text box, fall back to the persisted
    // setting, then let the service derive/scan (using the Extensions local path as a hint).
    private string ResolveTeklaBin()
    {
        var configured = (TeklaBin ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(configured))
        {
            configured = _persistedTeklaBin;
        }
        return _service.ResolveTeklaBin(configured, _extensionsLocalPath);
    }

    // Ported from MainWindow.ResolveModelSharingIdentity: derive a stable email + display name from the device
    // identity the shell supplied (device id / issued-to), with the same fallbacks as the legacy code-behind.
    private (string Email, string Name) ResolveIdentity()
    {
        var deviceId = (_identity.DeviceId ?? string.Empty).Trim();
        var issuedTo = (_identity.IssuedTo ?? string.Empty).Trim();

        // Email = stable, unique key the coordinator maps to a user. Prefer issued_to if it is already an email,
        // else the device id (server-assigned, unique per PC) shaped as an email, then user-machine as last resort.
        string email;
        if (issuedTo.Contains('@'))
        {
            email = issuedTo.ToLowerInvariant();
        }
        else if (!string.IsNullOrWhiteSpace(deviceId))
        {
            var sanitized = new string(deviceId.ToLowerInvariant()
                .Select(ch => char.IsLetterOrDigit(ch) || ch == '-' || ch == '.' ? ch : '-')
                .ToArray());
            email = sanitized + "@structura-most.ru";
        }
        else
        {
            email = (Environment.UserName + "-" + Environment.MachineName).ToLowerInvariant() + "@structura-most.ru";
        }

        var name = !string.IsNullOrWhiteSpace(issuedTo)
            ? issuedTo
            : (!string.IsNullOrWhiteSpace(deviceId) ? deviceId : Environment.MachineName);
        return (email, name);
    }

    private void RaiseCanExec() => (SetupCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
}

// Device identity the shell pushes in (replaces direct AppSettings.DeviceId/IssuedTo reads). HasDevice gates Setup.
public sealed record ModelSharingIdentity(string DeviceId, string IssuedTo, bool HasDevice);

// Applied provisioning values handed back to the shell so it persists the ModelSharing* settings keys + Save().
public sealed record ModelSharingProvisionedInfo(string TeklaBin, string ServerHost, int ServerPort,
    string IdentityEmail, DateTimeOffset AppliedUtc);

// View-supplied progress handle wrapping the shell's OperationProgressWindow (kept out of the VM — it is a
// Window-owned helper). Every member is optional so the VM degrades gracefully when no progress UI is wired.
public sealed class ModelSharingProgressHandle
{
    public Action? ShowAction { get; init; }
    public Action<string, string, int, int, TimeSpan>? StepAction { get; init; }
    public Action<string>? SucceededAction { get; init; }
    public Action<string>? FailedAction { get; init; }
    public Action? DismissAction { get; init; }

    public void Show() => ShowAction?.Invoke();
    public void Step(string step, string details, int current, int total, TimeSpan eta) =>
        StepAction?.Invoke(step, details, current, total, eta);
    public void Succeeded(string details) => SucceededAction?.Invoke(details);
    public void Failed(string details) => FailedAction?.Invoke(details);
    public void Dismiss() => DismissAction?.Invoke();
}
