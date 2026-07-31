using System.Windows;
using Connector.Desktop.Mvvm;
using Connector.Desktop.Services;

namespace Connector.Desktop.Features.Vpn;

// VPN domain module: "Общая папка (VPN)" access over the firm AmneziaWG tunnel. Owns the ViewModel + View
// (composition root for the feature). The shell wires the side-effect delegates (Log / dialog / open-share)
// and pushes the live VPN-state snapshot via PushContext after each token-connect (it owns the config/creds).
public sealed class VpnModule : IFeatureModule
{
    private readonly VpnView _view;
    private readonly VpnViewModel _viewModel;

    public VpnModule(VpnProvisioningService? service = null)
    {
        _viewModel = new VpnViewModel(service)
        {
            // Sensible default dialog so the module is usable before composition; the shell overrides these
            // (Log -> AppendLog, DialogHandler -> ThemedDialogs.Show owned by MainWindow, OpenShareHandler ->
            // ConnectSmbInternalAsync with the in-memory SMB creds) when it wires the module.
            DialogHandler = (msg, button, image) => System.Windows.MessageBox.Show(msg, "VPN", button, image)
        };
        _view = new VpnView { DataContext = _viewModel };
    }

    public string Title => "Общая папка (VPN)";
    public FrameworkElement View => _view;

    // Shell-owned side effects (wired at composition; mirror PatchingModule's ConfirmHandler approach).
    public Action<string>? Log
    {
        get => _viewModel.Log;
        set => _viewModel.Log = value;
    }

    public Action<string, MessageBoxButton, MessageBoxImage>? DialogHandler
    {
        get => _viewModel.DialogHandler;
        set => _viewModel.DialogHandler = value;
    }

    public Action<string>? OpenShareHandler
    {
        get => _viewModel.OpenShareHandler;
        set => _viewModel.OpenShareHandler = value;
    }

    // The shell pushes a fresh VPN-state snapshot here after each token-connect (and at bootstrap), then the
    // module redraws. Replaces the old direct control mutation in UpdateVpnUi.
    public void PushContext(VpnContext context) => _viewModel.SetContext(context);

    public Task<VpnResult> EnsureEnabledAsync(
        bool showResultDialog = false,
        bool openShareOnSuccess = false) =>
        _viewModel.EnsureEnabledAsync(showResultDialog, openShareOnSuccess);

    public Task OnActivatedAsync()
    {
        _viewModel.RefreshStatus();
        return Task.CompletedTask;
    }
}
