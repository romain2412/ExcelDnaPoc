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
    public static void Run()
    {
        dynamic target = ((dynamic)ExcelDnaUtil.Application).ActiveCell;
        target.Value2 = "Chargement de la blague...";

        WpfPane pane = AddInServices.TaskPane.Show();
        AddInServices.Joke.Start(target, pane);
    }
}
