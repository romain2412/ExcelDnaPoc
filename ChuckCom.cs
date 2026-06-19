using System.Runtime.InteropServices;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : objet COM (C#) consomme depuis VBA
// =====================================================================
// ExcelDna fait de l'add-in un SERVEUR COM : une classe [ComVisible] avec un ProgId
// devient instanciable depuis VBA par CreateObject("ExcelDnaPoc.Chuck"). Cote build,
// <ExcelAddInComServer>true</ExcelAddInComServer> (-> ComServer="true" dans le .dna)
// et ComServer.DllRegisterServer() dans AutoOpen (cf. AddIn.cs) rendent le ProgId
// resoluble.
//
// On expose une INTERFACE IDispatch explicite : c'est ce que VBA utilise en
// late-binding (Dim chuck As Object). [ClassInterface(None)] + l'interface = pattern
// recommande (vs AutoDual).
//
// LancerBlague() DELEGUE au declencheur commun ChuckTrigger (le meme que le bouton
// ruban "Blague (async)", les menus contextuels et le double-clic) : il ouvre le
// VOLET WPF, cree un JokeJob (sa propre ligne avec BARRE DE PROGRESSION + bouton
// ANNULER) et lance l'async (API + attente 15s) en FIRE-AND-FORGET. La methode COM
// rend donc la main aussitot -> la macro VBA se termine, Excel reste reactif, et le
// traitement est suivable/annulable depuis le volet. (Pas de logique async dupliquee
// ici : tout vit deja dans ChuckTrigger/JokeJob/WpfPane.)
[ComVisible(true)]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IChuckCom
{
    void LancerBlague();
}

[ComVisible(true)]
[Guid("7E3C1A92-4B8D-4E2F-9C6A-1D5B8E0F2A34")]
[ProgId("ExcelDnaPoc.Chuck")]
[ClassInterface(ClassInterfaceType.None)]
public class ChuckCom : IChuckCom
{
    public void LancerBlague() => ChuckTrigger.Fire();
}
