using System;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "VBA + COM" — un bouton qui appelle une MACRO VBA, laquelle
// utilise l'objet COM C# (ChuckCom) en fire-and-forget. Demontre l'aller-retour
// complet ruban (C#) -> macro VBA -> objet COM (C#).
// Application.Run resout la macro dans n'importe quel classeur ouvert : le module
// ChuckMacro.bas doit donc etre importe dans un .xlsm ouvert (cf. README).
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    private const string VbaComXml = """
        <group id='grpVbaCom' label='VBA + COM'>
          <button id='btnVbaChuck' label='Blague COM (VBA)' size='large'
                  imageMso='MacroPlay'
                  screentip='Appelle la macro VBA LancerBlagueAsync (qui utilise l objet COM ChuckCom en fire-and-forget)'
                  onAction='OnRunVbaChuck'/>
        </group>
        """;

    public void OnRunVbaChuck(IRibbonControl control)
    {
        try
        {
            ((dynamic)ExcelDnaUtil.Application).Run("LancerBlagueAsync");
        }
        catch (Exception ex)
        {
            // Cas le plus courant : le module VBA n'est pas importe dans un classeur ouvert.
            Log.Error("OnRunVbaChuck (macro VBA introuvable ?)", ex);
            try
            {
                ((dynamic)ExcelDnaUtil.Application).StatusBar =
                    "Macro 'LancerBlagueAsync' introuvable : importe ChuckMacro.bas dans un classeur .xlsm ouvert.";
            }
            catch { /* ignore */ }
        }
    }
}
