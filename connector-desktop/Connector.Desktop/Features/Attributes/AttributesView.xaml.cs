namespace Connector.Desktop.Features.Attributes;

// Passive placeholder view for the "Атрибуты" domain (no ViewModel yet — it is a roadmap card).
// When real attribute features land, each becomes its own IFeatureModule (View + ViewModel) registered
// under the Attributes FeatureDomain in ShellViewModel, mirroring the Tekla domain.
public partial class AttributesView : System.Windows.Controls.UserControl
{
    public AttributesView()
    {
        InitializeComponent();
    }
}
