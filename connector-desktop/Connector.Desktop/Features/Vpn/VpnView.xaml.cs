namespace Connector.Desktop.Features.Vpn;

// View is passive: its ViewModel is supplied via DataContext by VpnModule (the shell owns the VM/service and
// the live config/creds it pushes in via VpnViewModel.SetContext). No Click handlers — buttons bind to commands.
public partial class VpnView : System.Windows.Controls.UserControl
{
    public VpnView()
    {
        InitializeComponent();
        Loaded += (_, _) => Vm?.RefreshStatus();
    }

    private VpnViewModel? Vm => DataContext as VpnViewModel;
}
