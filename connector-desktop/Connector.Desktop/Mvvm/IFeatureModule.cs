using System.Windows;

namespace Connector.Desktop.Mvvm;

// A self-contained feature shown as one sub-tab inside a domain (e.g. Tekla → Патчинг).
// Adding a feature = implement this + register it in a FeatureDomain in ShellViewModel — no MainWindow bloat.
public interface IFeatureModule
{
    string Title { get; }                  // sub-tab header
    FrameworkElement View { get; }         // the module's UserControl (created once, lazily)
    Task OnActivatedAsync();               // called when the module's tab becomes active (refresh status, etc.)
}

// A group of feature modules under one top-level tab (e.g. "Tekla", later "Атрибуты").
public sealed class FeatureDomain
{
    public string Title { get; }
    public IReadOnlyList<IFeatureModule> Modules { get; }

    public FeatureDomain(string title, IReadOnlyList<IFeatureModule> modules)
    {
        Title = title;
        Modules = modules;
    }

    public IFeatureModule? Module(string title) =>
        Modules.FirstOrDefault(m => string.Equals(m.Title, title, StringComparison.OrdinalIgnoreCase));
}
