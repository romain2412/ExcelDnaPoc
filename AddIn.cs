using ExcelDna.Integration;
using ExcelDna.IntelliSense;

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
        // Active l'IntelliSense en ligne pour nos UDF (liste + infobulle d'arguments).
        // Uninstall() OBLIGATOIRE dans AutoClose, sinon Excel crashe au dechargement.
        IntelliSenseServer.Install();
        // Enregistre le serveur COM (ProgId ExcelDnaPoc.Chuck) -> CreateObject() en VBA.
        ExcelDna.ComInterop.ComServer.DllRegisterServer();
        CellRightClickInterceptor.Hook();
        StartupLoader.Run(); // ouvre les classeurs definis dans startup.json (async)
        Log.Info("=== AutoOpen fin ===");
    }

    public void AutoClose()
    {
        // Desabonnement propre au dechargement de l'add-in.
        IntelliSenseServer.Uninstall(); // pendant obligatoire de Install() (cf. AutoOpen)
        ExcelDna.ComInterop.ComServer.DllUnregisterServer(); // pendant de DllRegisterServer()
        CellRightClickInterceptor.Unhook();
        WorkbookEventsBinder.UnbindAll();
        WorksheetEventsBinder.UnbindAll();
    }
}
