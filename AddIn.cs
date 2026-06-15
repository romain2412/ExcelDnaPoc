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
        // Exemple de point d'initialisation au chargement de l'add-in.
        // (Le ruban et le volet sont, eux, crees a la demande par Excel/ExcelDna.)
    }

    public void AutoClose()
    {
        // Exemple de point de nettoyage au dechargement de l'add-in.
    }
}
