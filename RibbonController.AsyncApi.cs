using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "Reseau" — adaptateur ruban -> JokeApiService.
// Le ruban capture la cible (cellule active), ouvre le volet, puis delegue
// tout le travail asynchrone (appel API, attente, annulation) a JokeApiService.
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    private const string AsyncApiXml = """
        <group id='grpReseau' label='Reseau'>
          <button id='btnChuck' label='Blague (async)' size='large'
                  imageMso='HyperlinkInsert' onAction='OnChuckNorrisClick'
                  screentip='Appel async a api.chucknorris.io'
                  supertip='Recupere une blague via Internet sans bloquer Excel, puis l&apos;ecrit dans la cellule active APRES une attente async de 45s.'/>
          <button id='btnCancel' label='Annuler' size='large'
                  imageMso='Cancel' onAction='OnCancelClick' getEnabled='GetCancelEnabled'
                  screentip='Annule l&apos;attente async en cours'/>
        </group>
        """;

    // Clic : s'execute sur le thread principal d'Excel. On capture la cellule active
    // de la feuille affichee, on ouvre le volet (barre de progression), puis on delegue.
    public void OnChuckNorrisClick(IRibbonControl control)
    {
        dynamic target = ((dynamic)ExcelDnaUtil.Application).ActiveCell;
        target.Value2 = "Chargement de la blague...";

        WpfPane pane = _taskPane.Show(); // ouvre le volet et renvoie son contenu WPF
        _joke.Start(target, pane);       // lance l'operation async (fire-and-forget)
    }

    public void OnCancelClick(IRibbonControl control) => _joke.Cancel();

    public bool GetCancelEnabled(IRibbonControl control) => _joke.IsRunning;
}
