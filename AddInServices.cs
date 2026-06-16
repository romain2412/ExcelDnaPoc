using ExcelDna.Integration;

namespace ExcelDnaPoc;

// Services partages par TOUS les points d'entree (ruban, menu contextuel CustomUI,
// interception du clic droit, double-clic) : le volet est unique ; les TRAITEMENTS,
// eux, sont independants (un JokeJob par declenchement) et tournent en parallele.
public static class AddInServices
{
    public static readonly TaskPaneController TaskPane = new();
}

// Declencheur commun du comportement "Blague async".
// Capture la cellule cible, ouvre le volet WPF, cree un traitement INDEPENDANT (JokeJob)
// avec sa propre ligne d'UI (barre + Annuler), puis le lance. Plusieurs declenchements
// => plusieurs jobs en parallele, chacun annulable separement.
// Appele par : bouton ruban, menu contextuel (Solution 2), menus custom (Solution 1), double-clic.
public static class ChuckTrigger
{
    // Logique ASYNC de bout en bout : capture la cible, ouvre le volet, cree+lance un job.
    public static async Task RunAsync()
    {
        dynamic active = ((dynamic)ExcelDnaUtil.Application).ActiveCell;
        // Cible = cellule immediatement a DROITE de la cellule active.
        dynamic target = active.Worksheet.Cells[(int)active.Row, (int)active.Column + 1];
        string label = (string)target.Address;

        target.Value2 = "Chargement de la blague...";

        WpfPane pane = AddInServices.TaskPane.Show();
        var job = new JokeJob(target);
        pane.AddJob(job, label);     // cree la ligne d'UI (barre + Annuler) du traitement
        await job.RunAsync();        // ne touche PAS aux autres jobs en cours
    }

    // Frontiere fire-and-forget pour les callbacks Excel/COM SYNCHRONES (ruban onAction,
    // sinks d'evenements, macros) : ils ne peuvent PAS awaiter -> on lance le Task ici.
    public static void Fire() => _ = RunSafeAsync();

    // Variante AWAITABLE qui observe les exceptions (log). A utiliser depuis les handlers
    // d'evenements UI WinForms/WPF, qui peuvent eux etre "async void" et l'awaiter
    // directement -> l'async remonte alors jusqu'au handler (cf. menus contextuels custom).
    public static async Task RunSafeAsync()
    {
        try { await RunAsync(); }
        catch (Exception ex) { Log.Error("ChuckTrigger", ex); }
    }
}
