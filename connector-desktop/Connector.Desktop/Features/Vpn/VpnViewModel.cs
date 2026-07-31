using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Connector.Desktop.Mvvm;
using Connector.Desktop.Services;

namespace Connector.Desktop.Features.Vpn;

// Immutable snapshot of the VPN-related shell state the module needs to render and act. The shell owns the
// originals (in-memory config/SMB creds from the token-connect flow + the bootstrap-populated _settings.Vpn*),
// and pushes a fresh copy via VpnViewModel.SetContext after each token-connect — the VM never owns SettingsService.
public sealed record VpnContext(
    bool Enabled,
    string SmbUnc,
    string ServerIp,
    string VpnConfig,
    string SmbLogin,
    string SmbPassword)
{
    public static readonly VpnContext Empty = new(false, "", "", "", "", "");
}

// ViewModel for the "Общая папка (VPN)" feature module. Drives VpnProvisioningService: install/start, stop,
// and surface tunnel status for the firm AmneziaWG tunnel that exposes the SMB share from any network. The big
// SMB-mount helper and the live config/creds stay in the shell; the VM only requests "open this UNC" via a callback.
public sealed class VpnViewModel : ObservableObject
{
    private readonly VpnProvisioningService _vpnService;

    // The client tunnel name is always the local constant (the server-side name may differ).
    private const string VpnTunnel = "structura-vpn";

    // VPN-state snapshot injected by the shell (see SetContext). Defaults to empty until the first push.
    private VpnContext _ctx = VpnContext.Empty;

    // Shell-owned side effects, wired by VpnModule (mirror PatchingViewModel.ConfirmHandler):
    public Action<string>? Log { get; set; }                              // -> AppendLog
    public Action<string, MessageBoxButton, MessageBoxImage>? DialogHandler { get; set; }  // -> ThemedDialogs.Show
    public Action<string>? OpenShareHandler { get; set; }                 // -> shell mounts the UNC over VPN (SMB creds)

    public VpnViewModel(VpnProvisioningService? service = null)
    {
        _vpnService = service ?? new VpnProvisioningService();

        EnableCommand = new AsyncRelayCommand(EnableAsync, () => !IsBusy && CanEnable);
        DisableCommand = new AsyncRelayCommand(DisableAsync, () => !IsBusy && CanDisable);
        OpenFolderCommand = new RelayCommand(OpenFolder, () => CanOpenFolder);
        OpenLogCommand = new RelayCommand(OpenLog);
    }

    private string _statusLine = "Статус: проверка...";
    public string StatusLine { get => _statusLine; private set => SetProperty(ref _statusLine, value); }

    private System.Windows.Media.Brush _statusBrush = System.Windows.Media.Brushes.DarkGray;
    public System.Windows.Media.Brush StatusBrush { get => _statusBrush; private set => SetProperty(ref _statusBrush, value); }

    private string _shareUnc = "-";
    public string ShareUnc { get => _shareUnc; private set => SetProperty(ref _shareUnc, value); }

    private bool _canEnable;
    public bool CanEnable { get => _canEnable; private set { if (SetProperty(ref _canEnable, value)) RaiseCanExec(); } }

    private bool _canDisable;
    public bool CanDisable { get => _canDisable; private set { if (SetProperty(ref _canDisable, value)) RaiseCanExec(); } }

    private bool _canOpenFolder;
    public bool CanOpenFolder { get => _canOpenFolder; private set { if (SetProperty(ref _canOpenFolder, value)) RaiseCanExec(); } }

    private string _enableButtonText = "Включить VPN";
    public string EnableButtonText { get => _enableButtonText; private set => SetProperty(ref _enableButtonText, value); }

    private bool _isBusy;
    public bool IsBusy { get => _isBusy; private set { if (SetProperty(ref _isBusy, value)) RaiseCanExec(); } }

    public ICommand EnableCommand { get; }
    public ICommand DisableCommand { get; }
    public ICommand OpenFolderCommand { get; }
    public ICommand OpenLogCommand { get; }

    // The shell calls this after each token-connect (and at bootstrap) to push a fresh VPN-state snapshot;
    // pair it with RefreshStatus() to redraw.
    public void SetContext(VpnContext context)
    {
        _ctx = context ?? VpnContext.Empty;
        RefreshStatus();
    }

    private string ResolvedUnc => string.IsNullOrWhiteSpace(_ctx.SmbUnc)
        ? (string.IsNullOrWhiteSpace(_ctx.ServerIp) ? "" : $@"\\{_ctx.ServerIp}\BIM_Models")
        : _ctx.SmbUnc;

    // Reproduces UpdateVpnUi's branching: render status text/colour and which actions are available, from the
    // service probes (BundledClientPresent / IsTunnelInstalled / IsTunnelRunning) + the injected context.
    public void RefreshStatus()
    {
        try
        {
            var tunnel = VpnTunnel;
            var unc = ResolvedUnc;
            ShareUnc = string.IsNullOrWhiteSpace(unc) ? "-" : unc;

            if (!_ctx.Enabled)
            {
                // VPN turned off by the server — but if a tunnel is still installed from before, the user
                // must retain a way to remove it (otherwise it keeps routing across reboots).
                var stillInstalled = _vpnService.BundledClientPresent && _vpnService.IsTunnelInstalled(tunnel);
                if (stillInstalled)
                {
                    StatusLine = "VPN отключён администратором, но туннель ещё установлен на этом ПК. Нажмите «Отключить VPN», чтобы удалить его.";
                    StatusBrush = System.Windows.Media.Brushes.Orange;
                    CanEnable = false;
                    CanDisable = true;
                    CanOpenFolder = false;
                }
                else
                {
                    StatusLine = "VPN-доступ к общей папке не настроен администратором (или не требуется в вашей сети).";
                    StatusBrush = System.Windows.Media.Brushes.DarkGray;
                    CanEnable = false;
                    CanDisable = false;
                    CanOpenFolder = false;
                }
                return;
            }

            if (!_vpnService.BundledClientPresent)
            {
                StatusLine = "Встроенный VPN-клиент отсутствует в этой сборке коннектора. Обновите коннектор.";
                StatusBrush = System.Windows.Media.Brushes.Orange;
                CanEnable = false;
                CanDisable = false;
                CanOpenFolder = false;
                return;
            }

            var running = _vpnService.IsTunnelRunning(tunnel);
            if (running)
            {
                StatusLine = "VPN включён. Общая папка доступна из любой сети: " + (string.IsNullOrWhiteSpace(unc) ? "(адрес не задан)" : unc);
                StatusBrush = System.Windows.Media.Brushes.MediumSpringGreen;
                EnableButtonText = "Переподключить VPN";
                CanEnable = true;
                CanDisable = true;
                CanOpenFolder = !string.IsNullOrWhiteSpace(unc);
            }
            else
            {
                StatusLine = "VPN-доступ настроен, но не включён. Нажмите «Включить VPN» (потребуется однократное подтверждение прав администратора).";
                StatusBrush = System.Windows.Media.Brushes.Orange;
                EnableButtonText = "Включить VPN";
                CanEnable = true;
                CanDisable = _vpnService.IsTunnelInstalled(tunnel);
                CanOpenFolder = false;
            }
        }
        catch (Exception ex) { StatusLine = "Ошибка проверки статуса VPN: " + ex.Message; }
    }

    private async Task EnableAsync()
    {
        await EnsureEnabledAsync(showResultDialog: true, openShareOnSuccess: true);
    }

    public async Task<VpnResult> EnsureEnabledAsync(
        bool showResultDialog,
        bool openShareOnSuccess)
    {
        try
        {
            if (!_ctx.Enabled)
            {
                const string message = "VPN-доступ не включён администратором.";
                if (showResultDialog)
                {
                    DialogHandler?.Invoke(message, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                return VpnResult.Fail(message);
            }
            if (string.IsNullOrWhiteSpace(_ctx.VpnConfig))
            {
                const string message =
                    "Нет конфигурации VPN. Нажмите «Подключиться по токену» на вкладке «Коннектор», затем повторите.";
                if (showResultDialog)
                {
                    DialogHandler?.Invoke(message, MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                return VpnResult.Fail(message);
            }

            var tunnel = VpnTunnel;
            IsBusy = true;
            StatusLine = "Проверяем и включаем VPN. При первой установке подтвердите запрос UAC.";
            var result = await Task.Run(() => _vpnService.Enable(_ctx.VpnConfig, tunnel));
            Log?.Invoke("VPN enable: " + (result.IsSuccess ? "ok" : "fail") + " — " + result.Message);
            if (showResultDialog)
            {
                DialogHandler?.Invoke(
                    result.Message,
                    MessageBoxButton.OK,
                    result.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            if (result.IsSuccess &&
                openShareOnSuccess &&
                !string.IsNullOrWhiteSpace(ResolvedUnc))
            {
                Log?.Invoke("VPN включён. Подключаю общую папку автоматически.");
                OpenShareHandler?.Invoke(ResolvedUnc);
            }
            return result;
        }
        catch (Exception ex)
        {
            Log?.Invoke("VPN enable error: " + ex.Message);
            if (showResultDialog)
            {
                DialogHandler?.Invoke(ex.Message, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return VpnResult.Fail(ex.Message);
        }
        finally { IsBusy = false; RefreshStatus(); }
    }

    private async Task DisableAsync()
    {
        try
        {
            var tunnel = VpnTunnel;
            IsBusy = true;
            var result = await Task.Run(() => _vpnService.Disable(tunnel));
            Log?.Invoke("VPN disable: " + (result.IsSuccess ? "ok" : "fail") + " — " + result.Message);
            DialogHandler?.Invoke(result.Message,
                MessageBoxButton.OK, result.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }
        catch (Exception ex)
        {
            Log?.Invoke("VPN disable error: " + ex.Message);
            DialogHandler?.Invoke(ex.Message, MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally { IsBusy = false; RefreshStatus(); }
    }

    private void OpenFolder()
    {
        var unc = ResolvedUnc;
        if (string.IsNullOrWhiteSpace(unc)) return;
        // The share over VPN is a different host (tunnel IP), so it needs the SMB login mounted explicitly —
        // the shell owns the big ConnectSmbInternalAsync helper and the creds, so we just request the open.
        OpenShareHandler?.Invoke(unc);
    }

    private void OpenLog()
    {
        try
        {
            if (!File.Exists(_vpnService.LogFilePath))
                File.WriteAllText(_vpnService.LogFilePath, string.Empty);
            Process.Start(new ProcessStartInfo { FileName = _vpnService.LogFilePath, UseShellExecute = true });
        }
        catch (Exception ex) { Log?.Invoke("VPN open log error: " + ex.Message); }
    }

    private void RaiseCanExec()
    {
        (EnableCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (DisableCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (OpenFolderCommand as RelayCommand)?.RaiseCanExecuteChanged();
    }
}
