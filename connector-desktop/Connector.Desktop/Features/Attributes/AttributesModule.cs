using System.Windows;
using Connector.Desktop.Mvvm;

namespace Connector.Desktop.Features.Attributes;

// Placeholder module for the future "Атрибуты" domain (mapping / validation / filling / IDS-КСИ).
// It reserves the domain slot and demonstrates the extension point: future attribute features are added
// as sibling IFeatureModule entries in the Attributes FeatureDomain — no MainWindow changes required.
public sealed class AttributesModule : IFeatureModule
{
    private readonly AttributesView _view = new();

    public string Title => "Атрибуты";
    public FrameworkElement View => _view;

    public Task OnActivatedAsync() => Task.CompletedTask;
}
