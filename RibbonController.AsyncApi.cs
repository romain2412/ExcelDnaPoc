using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "Reseau" — adaptateur ruban -> ChuckTrigger.
// Le bouton delegue au declencheur commun, qui cree un traitement independant (JokeJob)
// avec sa propre barre + bouton Annuler dans le volet WPF.
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    private const string AsyncApiXml = """
        <group id='grpReseau' label='Reseau'>
          <button id='btnChuck' label='Blague (async)' size='large'
                  imageMso='HyperlinkInsert' onAction='OnChuckNorrisClick'
                  screentip='Appel async a api.chucknorris.io'
                  supertip='Recupere une blague via Internet sans bloquer Excel, puis l&apos;ecrit dans la cellule a droite APRES une attente async de 15s.'/>
        </group>
        """;

    // Clic ruban OU entree du menu contextuel (Solution 2) : meme comportement,
    // delegue au declencheur commun (capture cellule active, ouvre le volet, lance l'async).
    // NB : l'annulation se fait via le bouton "Annuler" du volet WPF (a cote de la barre).
    public void OnChuckNorrisClick(IRibbonControl control) => ChuckTrigger.Fire();
}
