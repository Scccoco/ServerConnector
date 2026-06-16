using System.Windows;
using Connector.Desktop.Mvvm;

namespace Connector.Desktop.Features.Structura;

// Structura domain module: quick links + credentials for the firm web services (Speckle / Nextcloud).
// Owns the ViewModel + View (composition root for the feature). The shell pushes the six Structura settings
// strings (URLs/logins/encrypted passwords) plus the logging + decrypt callbacks in via the ctor.
public sealed class StructuraModule : IFeatureModule
{
    private readonly StructuraView _view;
    private readonly StructuraViewModel _viewModel;

    public StructuraModule(Action<string>? log = null, Func<string, string>? decrypt = null)
    {
        _viewModel = new StructuraViewModel(log, decrypt);
        _view = new StructuraView { DataContext = _viewModel };
    }

    public string Title => "Structura";
    public FrameworkElement View => _view;

    // Shell wires these after construction (the catalog builds the module arg-less).
    public Action<string>? Log { get => _viewModel.Log; set => _viewModel.Log = value; }
    public Func<string, string> Decrypt { get => _viewModel.Decrypt; set => _viewModel.Decrypt = value; }

    // The shell feeds the current settings snapshot here (after bootstrap / on load) so the VM can drive the UI.
    public void Initialize(
        string speckleUrl, string speckleLogin, string specklePasswordCipher,
        string nextcloudUrl, string nextcloudLogin, string nextcloudPasswordCipher)
    {
        _viewModel.Initialize(
            speckleUrl, speckleLogin, specklePasswordCipher,
            nextcloudUrl, nextcloudLogin, nextcloudPasswordCipher);
    }

    public Task OnActivatedAsync()
    {
        _viewModel.RefreshStatus();
        return Task.CompletedTask;
    }
}
