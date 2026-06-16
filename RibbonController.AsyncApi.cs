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
                  supertip='Recupere une blague via Internet sans bloquer Excel, puis l&apos;ecrit dans la cellule a droite APRES une attente async de 15s.'/>
          <button id='btnCancel' label='Annuler' size='large'
                  imageMso='Cancel' onAction='OnCancelClick' getEnabled='GetCancelEnabled'
                  screentip='Annule l&apos;attente async en cours'/>
        </group>
        """;

    // Clic ruban OU entree du menu contextuel (Solution 2) : meme comportement,
    // delegue au declencheur commun (capture cellule active, ouvre le volet, lance l'async).
    public void OnChuckNorrisClick(IRibbonControl control) => ChuckTrigger.Fire();

    public void OnCancelClick(IRibbonControl control) => _joke.Cancel();

    public bool GetCancelEnabled(IRibbonControl control) => _joke.IsRunning;
}
