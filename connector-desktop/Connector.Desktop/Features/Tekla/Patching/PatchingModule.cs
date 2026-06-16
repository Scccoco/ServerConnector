using System.Windows;
using Connector.Desktop.Mvvm;
using Connector.Desktop.Services;

namespace Connector.Desktop.Features.Tekla.Patching;

// Tekla domain module: IFC-export patching. Owns the ViewModel + View (composition root for the feature).
public sealed class PatchingModule : IFeatureModule
{
    private readonly PatchingView _view;
    private readonly PatchingViewModel _viewModel;

    public PatchingModule(IfcExportPatchService? service = null)
    {
        _viewModel = new PatchingViewModel(service)
        {
            // any-user self-patch guardrail: confirm before touching licensed Tekla binaries
            ConfirmHandler = msg =>
                System.Windows.MessageBox.Show(msg, "Патчинг Tekla", MessageBoxButton.OKCancel, MessageBoxImage.Warning)
                    == MessageBoxResult.OK
        };
        _view = new PatchingView { DataContext = _viewModel };
    }

    public string Title => "Патчинг";
    public FrameworkElement View => _view;

    public Task OnActivatedAsync()
    {
        _viewModel.RefreshStatus();
        return Task.CompletedTask;
    }
}
