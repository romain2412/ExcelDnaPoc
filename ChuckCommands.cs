using ExcelDna.Integration;

namespace ExcelDnaPoc;

// Macros ExcelDna (commandes) appelables par nom.
// Comme ExcelAddInExplicitExports=true, seules les methodes marquees [ExcelCommand]
// sont exportees/enregistrees aupres d'Excel.
public static class ChuckCommands
{
    // Appelee par le bouton du menu popup de la Solution 1 (CommandBarButton.OnAction).
    [ExcelCommand(Name = "ShowChuckJokeFromMenu")]
    public static void ShowChuckJokeFromMenu() => ChuckTrigger.Run();
}
