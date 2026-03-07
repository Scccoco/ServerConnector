using System.Windows;

namespace Connector.Desktop;

public partial class App : System.Windows.Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var window = new MainWindow
        {
            WindowStartupLocation = WindowStartupLocation.CenterScreen
        };

        MainWindow = window;
        window.Show();
        window.Activate();
    }
}
