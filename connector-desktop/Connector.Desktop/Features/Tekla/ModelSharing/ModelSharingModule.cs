using System.Windows;
using Connector.Desktop.Mvvm;
using Connector.Desktop.Services;

namespace Connector.Desktop.Features.Tekla.ModelSharing;

// Tekla domain module: self-hosted Model Sharing provisioning. Owns the ViewModel + View (composition root
// for the feature). The shell wires identity / server / persistence into the VM via Initialize(...) at
// composition time — the VM never reaches into AppSettings directly.
public sealed class ModelSharingModule : IFeatureModule
{
    private readonly ModelSharingView _view;
    private readonly ModelSharingViewModel _viewModel;

    public ModelSharingModule(ModelSharingProvisioningService? service = null)
    {
        _viewModel = new ModelSharingViewModel(service)
        {
            // Confirm before touching licensed Tekla binaries (mirrors the legacy ThemedDialogs Yes/No prompt;
            // the shell may override this with a window-owned themed dialog after composition).
            ConfirmHandler = msg =>
                System.Windows.MessageBox.Show(msg, "Model Sharing", MessageBoxButton.YesNo, MessageBoxImage.Question)
                    == MessageBoxResult.Yes,
            // Result / status notices (mirrors the legacy ThemedDialogs OK prompts). isError chooses the icon.
            ShowMessage = (msg, isError) =>
                System.Windows.MessageBox.Show(msg, "Model Sharing", MessageBoxButton.OK,
                    isError ? MessageBoxImage.Warning : MessageBoxImage.Information)
        };
        _view = new ModelSharingView { DataContext = _viewModel };
    }

    public string Title => "Model Sharing";
    public FrameworkElement View => _view;

    // Shell wires these after construction: Log -> AppendLog; OnProvisioned -> persist ModelSharing* keys + Save();
    // Initialize -> push current device identity / server endpoint / persisted bin / Extensions fallback into the VM.
    // ConfirmHandler/ShowMessage default to plain MessageBox (ctor) but the shell overrides them with window-owned
    // themed dialogs, mirroring the VPN module.
    public Action<string>? Log { get => _viewModel.Log; set => _viewModel.Log = value; }
    public Action<ModelSharingProvisionedInfo>? OnProvisioned { get => _viewModel.OnProvisioned; set => _viewModel.OnProvisioned = value; }
    public Func<string, bool>? ConfirmHandler { get => _viewModel.ConfirmHandler; set => _viewModel.ConfirmHandler = value; }
    public Action<string, bool>? ShowMessage { get => _viewModel.ShowMessage; set => _viewModel.ShowMessage = value; }

    public void Initialize(ModelSharingIdentity identity, string serverHost, int serverPort,
        string persistedTeklaBin, string extensionsLocalPath)
        => _viewModel.Initialize(identity, serverHost, serverPort, persistedTeklaBin, extensionsLocalPath);

    public Task OnActivatedAsync()
    {
        _viewModel.RefreshStatus();
        return Task.CompletedTask;
    }
}
