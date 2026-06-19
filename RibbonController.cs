using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : personnalisation du RUBAN (CustomUI)
// =====================================================================
// ExcelDna detecte automatiquement l'unique classe derivee de ExcelRibbon
// et s'en sert pour fournir le XML CustomUI et resoudre les callbacks (par nom).
//
// CONTRAINTE : tous les callbacks du ruban DOIVENT etre des methodes de cette
// classe (ExcelDna les invoque par nom sur cet objet). Pour separer le code par
// fonctionnalite tout en respectant cette contrainte, on utilise une CLASSE
// PARTIELLE repartie sur plusieurs fichiers RibbonController.*.cs ; chacun
// fournit son fragment XML (const) + ses callbacks.
//
// Les concepts reellement autonomes sont, eux, de VRAIES classes separees vers
// lesquelles le ruban delegue : TaskPaneController (volet WPF) et ChuckTrigger/JokeJob
// (appel async). Le cycle de vie est dans AddIn.cs (IExcelAddIn).
[ComVisible(true)]
public partial class RibbonController : ExcelRibbon
{
    // Reference au ruban (capturee au onLoad) pour appeler Invalidate.
    private IRibbonUI? _ribbon;

    // Service partage (meme instance que le menu contextuel et l'interception du clic
    // droit) : le volet (toggleButton du ruban) reflete son etat de visibilite.
    private readonly TaskPaneController _taskPane = AddInServices.TaskPane;

    public RibbonController()
    {
        // Quand le volet change de visibilite, on rafraichit le toggleButton du ruban.
        _taskPane.VisibleChanged += () => _ribbon?.InvalidateControl("tbWpf");
    }

    // Callback onLoad du CustomUI : on garde la reference IRibbonUI.
    public void OnRibbonLoad(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
        Log.Info("OnRibbonLoad - ruban ACCEPTE et charge par Excel.");
    }

    // Assemble le XML CustomUI a partir des fragments fournis par les fichiers partiels.
    // ATTENTION : c'est du XML. Dans un attribut delimite par des apostrophes, une
    // apostrophe litterale doit s'ecrire &apos; (et JAMAIS '' comme en VBA).
    public override string GetCustomUI(string ribbonId)
    {
        Log.Info($"GetCustomUI({ribbonId}) appele par Excel");
        string xml = BuildRibbonXml();
        Log.Info($"GetCustomUI -> {xml.Length} caracteres retournes");
        return xml;
    }

    private static string BuildRibbonXml() =>
        $"""
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabPoc' label='POC ExcelDna'>
                {DemonstrationXml}
                {TextToolsXml}
                {InteractiveXml}
                {TaskPaneXml}
                {AsyncApiXml}
                {UdfXml}
                {RtdXml}
              </tab>
            </tabs>
          </ribbon>
          {ContextMenuXml}
        </customUI>
        """;
}
