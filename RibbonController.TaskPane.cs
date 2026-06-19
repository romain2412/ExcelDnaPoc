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
          <button id='btnSpawnPane' label='Nouveau volet' size='large'
                  imageMso='WindowSwitchWindowsMenuExcel'
                  screentip='Cree un volet supplementaire (ancre alterne D/G ; les volets s empilent)'
                  onAction='OnSpawnPane'/>
          <button id='btnCloseAllPanes' label='Tout fermer' size='large'
                  imageMso='WindowClose'
                  screentip='Ferme tous les volets crees via Nouveau volet'
                  onAction='OnCloseAllPanes'/>
        </group>
        """;

    public bool GetWpfPanePressed(IRibbonControl control) => _taskPane.Visible;

    public void OnToggleWpfPane(IRibbonControl control, bool pressed)
        => _taskPane.SetVisible(pressed);

    // Volets "spawn dynamique" : un nouveau volet par clic ; "Tout fermer" les retire.
    // Les callbacks ruban tournent sur le thread principal d'Excel -> creation de CTP OK.
    public void OnSpawnPane(IRibbonControl control) => AddInServices.MultiPane.SpawnNew();

    public void OnCloseAllPanes(IRibbonControl control) => AddInServices.MultiPane.CloseAll();
}
