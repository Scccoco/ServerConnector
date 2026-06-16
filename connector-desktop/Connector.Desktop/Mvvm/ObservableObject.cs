using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Connector.Desktop.Mvvm;

// Minimal INotifyPropertyChanged base for ViewModels. Dependency-free (no CommunityToolkit) to keep the
// app's lean package set; if the MVVM surface grows we can swap to CommunityToolkit.Mvvm later.
public abstract class ObservableObject : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    protected bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}
