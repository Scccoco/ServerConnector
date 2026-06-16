using System.Diagnostics;
using System.IO;
using System.Windows.Input;
using Connector.Desktop.Mvvm;
using Connector.Desktop.Services;

namespace Connector.Desktop.Features.Tekla.Patching;

// ViewModel for the Tekla "Патчинг" (IFC-export patch) feature module. Drives IfcExportPatchService:
// detect the Tekla bin + build, apply / roll back the per-build patch set, and surface status. Apply is
// guarded by a confirmation the View supplies (any connected user may self-patch — see project decision).
public sealed class PatchingViewModel : ObservableObject
{
    private readonly IfcExportPatchService _service;

    // The View sets this to show a confirm dialog (returns true to proceed). Any-user self-patch guardrail.
    public Func<string, bool>? ConfirmHandler { get; set; }

    public PatchingViewModel(IfcExportPatchService? service = null)
    {
        _service = service ?? new IfcExportPatchService();

        DetectCommand = new AsyncRelayCommand(DetectAsync);
        ApplyCommand = new AsyncRelayCommand(ApplyAsync, () => !IsBusy && CanApply);
        RollbackCommand = new AsyncRelayCommand(RollbackAsync, () => !IsBusy && IsApplied);
        OpenLogCommand = new RelayCommand(OpenLog);

        _teklaBin = _service.ResolveTeklaBin(null, null);
        // Until the LFS delivery rail lands, default the staging set to the local PoC folder (browsable).
        _stagingDir = @"C:\Users\Lagom\ifc-poc\staging\2025";
    }

    private string _teklaBin = "";
    public string TeklaBin { get => _teklaBin; set { if (SetProperty(ref _teklaBin, value)) RefreshStatus(); } }

    private string _stagingDir = "";
    public string StagingDir { get => _stagingDir; set { SetProperty(ref _stagingDir, value); RaiseCanExec(); } }

    private string _detectedBuild = "";
    public string DetectedBuild { get => _detectedBuild; private set => SetProperty(ref _detectedBuild, value); }

    private string _statusLine = "Статус: не проверялось";
    public string StatusLine { get => _statusLine; private set => SetProperty(ref _statusLine, value); }

    private bool _isApplied;
    public bool IsApplied { get => _isApplied; private set { if (SetProperty(ref _isApplied, value)) RaiseCanExec(); } }

    private bool _isBusy;
    public bool IsBusy { get => _isBusy; private set { if (SetProperty(ref _isBusy, value)) RaiseCanExec(); } }

    private string _resultMessage = "";
    public string ResultMessage { get => _resultMessage; private set => SetProperty(ref _resultMessage, value); }

    public string LogFilePath => _service.LogFilePath;
    private bool CanApply => !string.IsNullOrWhiteSpace(TeklaBin) && Directory.Exists(StagingDir);

    public ICommand DetectCommand { get; }
    public ICommand ApplyCommand { get; }
    public ICommand RollbackCommand { get; }
    public ICommand OpenLogCommand { get; }

    public void RefreshStatus()
    {
        try
        {
            var s = _service.GetStatus(TeklaBin);
            DetectedBuild = s.DetectedBuild;
            IsApplied = s.Applied;
            StatusLine = !s.TeklaFound
                ? "Tekla не найдена по указанному пути bin"
                : s.Applied
                    ? $"Патч установлен (набор {s.SetVersion}, сборка {s.DetectedBuild})"
                    : s.NeedsReapply
                        ? "Патч был установлен, но файлы заменены обновлением Tekla — требуется переустановка"
                        : $"Патч не установлен (сборка {s.DetectedBuild})";
        }
        catch (Exception ex) { StatusLine = "Ошибка проверки статуса: " + ex.Message; }
    }

    private Task DetectAsync()
    {
        TeklaBin = _service.ResolveTeklaBin(string.IsNullOrWhiteSpace(TeklaBin) ? null : TeklaBin, null);
        RefreshStatus();
        return Task.CompletedTask;
    }

    private async Task ApplyAsync()
    {
        if (_service.IsTeklaRunning())
        {
            ResultMessage = "Сейчас запущена Tekla Structures. Закройте её и повторите.";
            return;
        }
        var confirm = ConfirmHandler?.Invoke(
            "Будут заменены системные файлы Tekla для расширения IFC-экспорта (с резервной копией для отката). " +
            "Tekla должна быть закрыта. Продолжить?") ?? true;
        if (!confirm) return;

        IsBusy = true;
        try
        {
            var req = new IfcPatchRequest { TeklaBin = TeklaBin, StagingDir = StagingDir };
            var r = await Task.Run(() => _service.Apply(req));
            ResultMessage = r.Message + (r.TechnicalDetails.Length > 0 ? "\n[detail] " + r.TechnicalDetails : "");
        }
        finally { IsBusy = false; RefreshStatus(); }
    }

    private async Task RollbackAsync()
    {
        if (_service.IsTeklaRunning())
        {
            ResultMessage = "Сейчас запущена Tekla Structures. Закройте её и повторите.";
            return;
        }
        IsBusy = true;
        try
        {
            var r = await Task.Run(() => _service.Rollback(TeklaBin));
            ResultMessage = r.Message + (r.TechnicalDetails.Length > 0 ? "\n[detail] " + r.TechnicalDetails : "");
        }
        finally { IsBusy = false; RefreshStatus(); }
    }

    private void OpenLog()
    {
        try
        {
            if (File.Exists(LogFilePath))
                Process.Start(new ProcessStartInfo { FileName = LogFilePath, UseShellExecute = true });
        }
        catch { /* best-effort */ }
    }

    private void RaiseCanExec()
    {
        (ApplyCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
        (RollbackCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
    }
}
