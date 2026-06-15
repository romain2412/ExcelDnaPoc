using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "Outils texte" — un bouton avec MENU DEROULANT (<menu>).
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    private const string TextToolsXml = """
        <group id='grpMenu' label='Outils texte'>
          <menu id='mnuActions' label='Transformer' size='large'
                imageMso='ListMacros'
                screentip='Actions sur la cellule active'
                supertip='Ouvre un menu d&apos;actions appliquees a la cellule active.'>
            <button id='miUpper' label='MAJUSCULES'
                    imageMso='FontDialog' onAction='OnUpperClick'/>
            <button id='miLower' label='minuscules'
                    imageMso='FontDialog' onAction='OnLowerClick'/>
            <menuSeparator id='sep1'/>
            <button id='miNow' label='Inserer date/heure'
                    imageMso='DateAndTimeInsert' onAction='OnNowClick'/>
            <menuSeparator id='sep2'/>
            <button id='miClear' label='Effacer la cellule'
                    imageMso='Delete' onAction='OnClearClick'/>
          </menu>
        </group>
        """;

    // Met la cellule active en MAJUSCULES.
    public void OnUpperClick(IRibbonControl control)
    {
        dynamic cell = ((dynamic)ExcelDnaUtil.Application).ActiveCell;
        cell.Value2 = (cell.Value2?.ToString() ?? string.Empty).ToUpperInvariant();
    }

    // Met la cellule active en minuscules.
    public void OnLowerClick(IRibbonControl control)
    {
        dynamic cell = ((dynamic)ExcelDnaUtil.Application).ActiveCell;
        cell.Value2 = (cell.Value2?.ToString() ?? string.Empty).ToLowerInvariant();
    }

    // Insere la date et l'heure courantes dans la cellule active.
    public void OnNowClick(IRibbonControl control)
    {
        dynamic cell = ((dynamic)ExcelDnaUtil.Application).ActiveCell;
        cell.Value2 = System.DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
    }

    // Efface le contenu de la cellule active.
    public void OnClearClick(IRibbonControl control)
    {
        ((dynamic)ExcelDnaUtil.Application).ActiveCell.ClearContents();
    }
}
