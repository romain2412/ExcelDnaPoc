using System;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelDna.Integration;

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
// FIRE-AND-FORGET : LancerBlague() demarre le travail async et REND LA MAIN tout de
// suite (la methode COM retourne immediatement). VBA n'attend pas -> Excel reste
// reactif pendant TOUT l'IO. L'objet finit le travail en arriere-plan (vrai async,
// aucun thread tenu) et ecrit la blague lui-meme via QueueAsMacro, comme JokeJob.
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
    private const string JokeUrl = "https://api.chucknorris.io/jokes/random";
    private static readonly HttpClient _http = new();

    public void LancerBlague()
    {
        // Sur le thread appelant (COM/VBA = thread principal d'Excel) : on capture la
        // cible MAINTENANT = cellule a droite de la cellule active.
        dynamic app = ExcelDnaUtil.Application;
        dynamic active = app.ActiveCell;
        dynamic target = active.Worksheet.Cells[(int)active.Row, (int)active.Column + 1];

        target.Value2 = "Chargement de la blague (Excel reste utilisable)...";
        app.StatusBar = "Chuck (COM) : appel API + attente 15s en arriere-plan...";

        _ = RunAsync(target); // FIRE-AND-FORGET : on ne await pas -> retour immediat
    }

    private static async Task RunAsync(dynamic target)
    {
        try
        {
            string json = await _http.GetStringAsync(JokeUrl);   // async : aucun thread bloque
            using var doc = JsonDocument.Parse(json);
            string joke = doc.RootElement.GetProperty("value").GetString() ?? "(reponse vide)";

            await Task.Delay(15_000);                            // attente async (timer, pas Thread.Sleep)

            // Ecriture de la cellule + reset du statut sur le thread d'Excel.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                target.Value2 = joke;
                ((dynamic)ExcelDnaUtil.Application).StatusBar = false; // restaure le statut par defaut
            });
        }
        catch (Exception ex)
        {
            Log.Error("ChuckCom.RunAsync", ex);
            ExcelAsyncUtil.QueueAsMacro(() =>
                ((dynamic)ExcelDnaUtil.Application).StatusBar = "Chuck (COM) : erreur - " + ex.Message);
        }
    }
}
