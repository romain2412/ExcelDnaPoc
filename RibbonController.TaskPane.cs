using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "Volet" — adaptateur ruban -> TaskPaneController.
// Le ruban se contente d'exposer les callbacks ; toute la logique du volet
// (creation, visibilite, hebergement WPF) vit dans TaskPaneController.
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    private const string TaskPaneXml = """
        <group id='grpVolet' label='Volet'>
          <toggleButton id='tbWpf' label='Volet WPF' size='large'
                        imageMso='WindowNew'
                        onAction='OnToggleWpfPane' getPressed='GetWpfPanePressed'/>
        </group>
        """;

    public bool GetWpfPanePressed(IRibbonControl control) => _taskPane.Visible;

    public void OnToggleWpfPane(IRibbonControl control, bool pressed)
        => _taskPane.SetVisible(pressed);
}
