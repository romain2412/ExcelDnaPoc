using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : cycle de vie de l'ADD-IN (IExcelAddIn)
// =====================================================================
// ExcelDna detecte automatiquement les classes implementant IExcelAddIn et
// appelle AutoOpen au chargement du .xll, AutoClose au dechargement.
// C'est le point d'entree pour initialiser/liberer des ressources globales
// (logging, connexions, prechauffage de caches, etc.).
public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        Log.Info("=== AutoOpen debut ===");
        CellRightClickInterceptor.Hook();
        StartupLoader.Run(); // ouvre les classeurs definis dans startup.json (async)
        Log.Info("=== AutoOpen fin ===");
    }

    public void AutoClose()
    {
        // Desabonnement propre au dechargement de l'add-in.
        CellRightClickInterceptor.Unhook();
        WorkbookEventsBinder.UnbindAll();
    }
}
