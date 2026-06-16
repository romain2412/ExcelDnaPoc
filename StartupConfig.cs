using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  Systeme de CONFIGURATION de demarrage de l'add-in
// =====================================================================
// Lue depuis un fichier "startup.json" place A COTE du .xll. Concue pour evoluer :
// on pourra y ajouter des criteres (conditions, feuilles a activer, parametres...).
// Pour l'instant : une simple liste de classeurs a ouvrir au demarrage.
//
// Exemple de startup.json :
//   { "openWorkbooks": [ "C:\\chemin\\TestAddin.xlsx" ] }
public class StartupConfig
{
    public const string FileName = "startup.json";

    // Classeurs a ouvrir au demarrage de l'add-in.
    public List<string> OpenWorkbooks { get; set; } = new();

    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
    };

    // Charge la config depuis le dossier du .xll. Renvoie null si absente/illisible.
    public static StartupConfig? Load()
    {
        try
        {
            string? dir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            if (dir is null)
                return null;

            string path = Path.Combine(dir, FileName);
            if (!File.Exists(path))
            {
                Log.Info($"Startup: aucune config ({path})");
                return null;
            }

            StartupConfig? config = JsonSerializer.Deserialize<StartupConfig>(
                File.ReadAllText(path), Options);
            Log.Info($"Startup: config chargee ({config?.OpenWorkbooks.Count ?? 0} classeur(s))");
            return config;
        }
        catch (System.Exception ex)
        {
            Log.Error("Startup: lecture de la config echouee", ex);
            return null;
        }
    }
}
