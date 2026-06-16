using System.IO;
using System.Windows;
using Forms = System.Windows.Forms;

namespace Connector.Desktop.Features.Tekla.Patching;

// View is passive: its ViewModel is supplied via DataContext by PatchingModule (the shell owns the VM/service).
public partial class PatchingView : System.Windows.Controls.UserControl
{
    public PatchingView()
    {
        InitializeComponent();
        Loaded += (_, _) => Vm?.RefreshStatus();
    }

    private PatchingViewModel? Vm => DataContext as PatchingViewModel;

    private void BrowseBin_Click(object sender, RoutedEventArgs e)
    {
        if (Vm is { } vm) Browse(p => vm.TeklaBin = p, vm.TeklaBin);
    }

    private void BrowseStaging_Click(object sender, RoutedEventArgs e)
    {
        if (Vm is { } vm) Browse(p => vm.StagingDir = p, vm.StagingDir);
    }

    private static void Browse(Action<string> set, string current)
    {
        using var dlg = new Forms.FolderBrowserDialog();
        if (!string.IsNullOrWhiteSpace(current) && Directory.Exists(current)) dlg.SelectedPath = current;
        if (dlg.ShowDialog() == Forms.DialogResult.OK) set(dlg.SelectedPath);
    }
}
