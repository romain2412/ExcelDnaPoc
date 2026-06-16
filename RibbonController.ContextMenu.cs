using System;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// =====================================================================
//  SOLUTION 2 : menu contextuel via CustomUI (callback getVisible)
// =====================================================================
// L'entree "Blague (async)" s'AJOUTE au menu Excel normal, mais uniquement quand
// la cellule active vaut "BlagueInterception" (sinon le menu reste celui d'Excel).
// Realise sans interop : un <button> dans <contextMenu idMso='ContextMenuCell'>,
// dont la visibilite est pilotee par getVisible (reevalue a chaque clic droit).
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    // Contenu de cellule qui declenche l'enrichissement du menu Excel.
    private const string CtxTrigger = "BlagueMenuExcel";

    // Section <contextMenus> du CustomUI. Le coeur (RibbonController.cs) l'insere
    // APRES <ribbon> (ordre impose par le schema : commands, ribbon, backstage, contextMenus).
    private const string ContextMenuXml = """
        <contextMenus>
          <contextMenu idMso='ContextMenuCell'>
            <button id='ctxChuck' label='Blague (async)'
                    imageMso='HyperlinkInsert' onAction='OnChuckNorrisClick'
                    getVisible='GetChuckCtxVisible' insertBeforeMso='Cut'/>
          </contextMenu>
        </contextMenus>
        """;

    // Visible seulement si la cellule active vaut "BlagueInterception".
    // Les menus contextuels etant reconstruits a chaque ouverture, ce callback est
    // reevalue a chaque clic droit -> apparition/disparition dynamique de l'entree.
    public bool GetChuckCtxVisible(IRibbonControl control)
    {
        try
        {
            object? value = ((dynamic)ExcelDnaUtil.Application).ActiveCell.Value2;
            return value is string s &&
                   string.Equals(s.Trim(), CtxTrigger, StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }
}
