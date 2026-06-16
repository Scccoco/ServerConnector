using System.Net.Http;
using System.Windows;
using Connector.Desktop.Mvvm;
using Connector.Desktop.Services;

namespace Connector.Desktop.Features.Tekla.Standard;

// Tekla domain module: "Стандарт" — firm-folder / extensions / Grasshopper-libraries git-sync + admin publish +
// Tekla-server restart. Owns the View, which holds the VERBATIM-lifted MainWindow logic. The shell supplies the
// IShellHost seam (live settings, persistence, transport, dialogs, the cross-tab/header status mirror, the tray) and
// drives the engine through the public pass-throughs (RunTeklaSyncAsync / RefreshUi) where it formerly called
// RunTeklaSyncCycleAsync() / UpdateTeklaUi() in MainWindow.
public sealed class StandardModule : IFeatureModule
{
    private readonly StandardView _view;

    // Builds the module. The IShellHost is REQUIRED (the lifted logic cannot function without the shell seam);
    // the orchestrator constructs MainWindow's IShellHost implementation and passes it here when registering the
    // module in ShellViewModel. The Стандарт-only TeklaStandardService (HttpClient with a 25s timeout, matching the
    // original MainWindow field) is constructed here unless one is supplied (e.g. for testing).
    public StandardModule(IShellHost host, TeklaStandardService? teklaStandardService = null)
    {
        var service = teklaStandardService ?? new TeklaStandardService(new HttpClient { Timeout = TimeSpan.FromSeconds(25) });
        _view = new StandardView(host, service);
    }

    public string Title => "Стандарт";
    public FrameworkElement View => _view;

    // ===== Public pass-throughs the shell uses =====================================================
    // Engine entry point. The shell calls this from: the _teklaSyncTimer tick, the _updateTimer tick, Window_Loaded,
    // and after token-connect (all with showDialogs:false, forceRefresh:false, autoApplyIfPossible:true), replacing
    // its old direct RunTeklaSyncCycleAsync(...) calls.
    public Task RunTeklaSyncAsync(bool showDialogs, bool forceRefresh, bool autoApplyIfPossible) =>
        _view.RunTeklaSyncAsync(showDialogs, forceRefresh, autoApplyIfPossible);

    // Redraw the Стандарт controls + push the overall status to the shell mirror/header. The shell calls this where
    // it formerly called UpdateTeklaUi() (e.g. ApplyAndPersist, ConnectByTokenInternalAsync, the ctor bootstrap).
    public void RefreshUi() => _view.RefreshUi();

    // Interactive sync WITH a progress window + already-running notice — the shell's Коннектор-tab mirror button
    // forwarder calls this (identical to the Стандарт-tab button), restoring the progress UI on that button.
    public Task RunInteractiveSyncAsync() => _view.RunInteractiveSyncAsync();

    // Capture typed-but-unbrowsed path edits into Settings before the shell reads them (Save/Start/SendNow).
    public void FlushPathEdits() => _view.FlushPathEdits();

    // Activated => refresh, mirroring the other modules' OnActivatedAsync.
    public Task OnActivatedAsync()
    {
        _view.RefreshUi();
        return Task.CompletedTask;
    }
}
