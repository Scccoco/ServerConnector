using System.Windows;

namespace Connector.Desktop.Features.Structura;

// View is passive: its ViewModel is supplied via DataContext by StructuraModule (the shell owns the VM).
// Owning/showing the credentials dialog is the View's concern (Window ownership), wired via the VM's
// ShowAccessHandler — the VM hands over already-decrypted credentials, the View only opens the window.
public partial class StructuraView : System.Windows.Controls.UserControl
{
    public StructuraView()
    {
        InitializeComponent();
        Loaded += (_, _) =>
        {
            if (Vm is { } vm)
            {
                vm.ShowAccessHandler = ShowAccessWindow;
                vm.RefreshStatus();
            }
        };
    }

    private StructuraViewModel? Vm => DataContext as StructuraViewModel;

    // (title, domain, login, decryptedPassword) → open the self-contained credentials window owned by this View.
    private void ShowAccessWindow(string title, string domain, string login, string password)
    {
        var dialog = new StructuraAccessWindow(title, domain, login, password)
        {
            Owner = Window.GetWindow(this)
        };
        dialog.ShowDialog();
    }
}
