using System;
using System.IO;
using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  DEMARRAGE de l'add-in : ouverture de classeurs pilotee par la config
// =====================================================================
// Appele depuis AddIn.AutoOpen. Lit StartupConfig et ouvre les classeurs demandes
// de maniere ASYNCHRONE (non bloquante) : on differe chaque ouverture sur le thread
// principal d'Excel via ExcelAsyncUtil.QueueAsMacro, APRES le chargement de l'add-in
// (xlAutoOpen ne doit pas etre bloque). Le critere actuel = liste de chemins ; il
// pourra evoluer (conditions d'ouverture, feuille a activer, etc.).
public static class StartupLoader
{
    public static void Run()
    {
        StartupConfig? config = StartupConfig.Load();
        if (config is null)
            return;

        foreach (string path in config.OpenWorkbooks)
        {
            string p = path; // capture pour le lambda
            ExcelAsyncUtil.QueueAsMacro(() => OpenWorkbook(p));
        }
    }

    private static void OpenWorkbook(string path)
    {
        try
        {
            string full = ResolvePath(path);
            if (!File.Exists(full))
            {
                Log.Error($"Startup: classeur introuvable -> {full}");
                return;
            }

            dynamic app = ExcelDnaUtil.Application;
            string name = Path.GetFileName(full);

            // Ne pas rouvrir un classeur deja ouvert.
            foreach (dynamic wb in app.Workbooks)
            {
                if (string.Equals((string)wb.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    Log.Info($"Startup: classeur deja ouvert -> {name}");
                    return;
                }
            }

            dynamic opened = app.Workbooks.Open(full);
            Log.Info($"Startup: classeur ouvert -> {full}");

            // On binde les evenements Excel de ce classeur sur des callbacks C#.
            WorkbookEventsBinder.Bind(opened);
        }
        catch (Exception ex)
        {
            Log.Error("Startup: ouverture du classeur echouee", ex);
        }
    }

    // Un chemin relatif est resolu par rapport au DOSSIER DU .xll (commun a F5 et au
    // script Launch-AddIn.ps1) -> config machine-independante, versionnable.
    private static string ResolvePath(string path)
    {
        if (Path.IsPathRooted(path))
            return path;

        string? xllDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
        return string.IsNullOrEmpty(xllDir) ? path : Path.Combine(xllDir, path);
    }
}
