using System.IO;
using System.Windows;
using Forms = System.Windows.Forms;

namespace Connector.Desktop.Features.Tekla.ModelSharing;

// View is passive: its ViewModel is supplied via DataContext by ModelSharingModule (the shell owns the VM/service).
// The only logic here is the WinForms folder picker (UI-only) and wiring the Window-owned progress helper to the
// VM as a callback — the VM must not reference Window/OperationProgressWindow directly.
public partial class ModelSharingView : System.Windows.Controls.UserControl
{
    public ModelSharingView()
    {
        InitializeComponent();
        Loaded += (_, _) =>
        {
            WireProgress();
            Vm?.RefreshStatus();
        };
    }

    private ModelSharingViewModel? Vm => DataContext as ModelSharingViewModel;

    private void Browse_Click(object sender, RoutedEventArgs e)
    {
        if (Vm is not { } vm)
        {
            return;
        }

        using var dialog = new Forms.FolderBrowserDialog
        {
            Description = "Выберите папку bin Tekla Structures",
            SelectedPath = (vm.TeklaBin ?? string.Empty).Trim()
        };
        if (dialog.ShowDialog() == Forms.DialogResult.OK)
        {
            // Setting TeklaBin re-runs RefreshStatus through the VM setter.
            vm.TeklaBin = dialog.SelectedPath;
        }
    }

    // Supply the VM a factory that builds a shell OperationProgressWindow (owned by this View's Window) and
    // exposes it as a plain handle of callbacks. Mirrors how PatchingView feeds View-only concerns to its VM.
    private void WireProgress()
    {
        if (Vm is not { } vm)
        {
            return;
        }

        vm.BeginProgress = () =>
        {
            var owner = Window.GetWindow(this);
            var window = new OperationProgressWindow("Model Sharing", "Готовим Tekla к совместной работе")
            {
                Owner = owner
            };
            return new ModelSharingProgressHandle
            {
                ShowAction = window.Show,
                StepAction = (step, details, current, total, eta) => window.UpdateStep(step, details, current, total, eta),
                SucceededAction = window.MarkSucceeded,
                FailedAction = window.MarkFailed,
                DismissAction = window.Close
            };
        };
    }
}
