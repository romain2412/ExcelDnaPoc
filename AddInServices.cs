using ExcelDna.Integration;

namespace ExcelDnaPoc;

// Services partages par TOUS les points d'entree (ruban, menu contextuel CustomUI,
// interception du clic droit) : une seule instance du volet et du service async,
// afin que le comportement et l'etat soient communs quel que soit le declencheur.
public static class AddInServices
{
    public static readonly TaskPaneController TaskPane = new();
    public static readonly JokeApiService Joke = new();
}

// Declencheur commun du comportement "Blague async".
// Capture la cellule active, ouvre le volet WPF, lance l'operation asynchrone.
// Appele par : bouton ruban, entree du menu contextuel (Solution 2) et popup (Solution 1).
public static class ChuckTrigger
{
    // Logique ASYNC de bout en bout : capture la cible, ouvre le volet, AWAIT le service.
    public static async Task RunAsync()
    {
        dynamic active = ((dynamic)ExcelDnaUtil.Application).ActiveCell;
        // Cible = cellule immediatement a DROITE de la cellule active.
        dynamic target = active.Worksheet.Cells[(int)active.Row, (int)active.Column + 1];

        target.Value2 = "Chargement de la blague...";

        WpfPane pane = AddInServices.TaskPane.Show();
        await AddInServices.Joke.RunAsync(target, pane);
    }

    // FRONTIERE fire-and-forget UNIQUE. Les callbacks Excel (ruban, menus, double-clic,
    // macros) sont forcement synchrones (void) et n'attendent pas le Task : on lance ici,
    // une seule fois, et on observe les exceptions (log). C'est le seul point ou l'async
    // "redescend" en synchrone, au plus haut niveau possible (la frontiere Excel).
    public static void Fire() => _ = FireAndLogAsync();

    private static async Task FireAndLogAsync()
    {
        try { await RunAsync(); }
        catch (Exception ex) { Log.Error("ChuckTrigger.Fire", ex); }
    }
}
