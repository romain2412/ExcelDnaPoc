using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "Interactif" — controles a etat (editBox, checkBox,
// toggleButton, dropDown, comboBox) + bouton dynamique rafraichi via Invalidate.
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    // Etat memorise (le nom saisi dans l'editBox).
    private string _name = "";

    private const string InteractiveXml = """
        <group id='grpInteractif' label='Interactif'>
          <editBox id='ebName' label='Nom' sizeString='WWWWWWWWWWWW'
                   onChange='OnNameChange' getText='GetNameText'/>
          <button id='btnGreet' getLabel='GetGreetLabel'
                  imageMso='HappyFace' onAction='OnGreet'/>
          <separator id='sepI1'/>
          <checkBox id='cbBold' label='Gras'
                    onAction='OnBoldToggle' getPressed='GetBoldPressed'/>
          <toggleButton id='tbItalic' label='Italique' imageMso='Italic'
                        onAction='OnItalicToggle' getPressed='GetItalicPressed'/>
          <separator id='sepI2'/>
          <dropDown id='ddColor' label='Fond'
                    onAction='OnColorChange' getSelectedItemIndex='GetColorIndex'>
            <item id='colNone' label='Aucune'/>
            <item id='colYellow' label='Jaune'/>
            <item id='colGreen' label='Vert'/>
            <item id='colRed' label='Rouge'/>
          </dropDown>
          <comboBox id='cboFont' label='Police' sizeString='WWWWWWWWWWWW'
                    onChange='OnFontChange'>
            <item id='fCalibri' label='Calibri'/>
            <item id='fArial' label='Arial'/>
            <item id='fTimes' label='Times New Roman'/>
          </comboBox>
        </group>
        """;

    // -- editBox "Nom" --
    public string GetNameText(IRibbonControl control) => _name;

    public void OnNameChange(IRibbonControl control, string text)
    {
        _name = text;
        // Le libelle du bouton "btnGreet" depend de _name : on demande son rafraichissement.
        _ribbon?.InvalidateControl("btnGreet");
    }

    // -- button dynamique "btnGreet" : libelle calcule via getLabel --
    public string GetGreetLabel(IRibbonControl control)
        => string.IsNullOrWhiteSpace(_name) ? "Dire bonjour" : $"Bonjour {_name} !";

    public void OnGreet(IRibbonControl control)
    {
        string qui = string.IsNullOrWhiteSpace(_name) ? "tout le monde" : _name;
        MessageBox.Show($"Bonjour {qui} !", "POC ExcelDna",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // -- checkBox "Gras" : reflete (getPressed) et applique (onAction) le gras --
    public bool GetBoldPressed(IRibbonControl control)
    {
        try
        {
            object? b = ((dynamic)ExcelDnaUtil.Application).ActiveCell.Font.Bold;
            return b is bool flag && flag;
        }
        catch { return false; }
    }

    public void OnBoldToggle(IRibbonControl control, bool pressed)
    {
        ((dynamic)ExcelDnaUtil.Application).ActiveCell.Font.Bold = pressed;
    }

    // -- toggleButton "Italique" --
    public bool GetItalicPressed(IRibbonControl control)
    {
        try
        {
            object? i = ((dynamic)ExcelDnaUtil.Application).ActiveCell.Font.Italic;
            return i is bool flag && flag;
        }
        catch { return false; }
    }

    public void OnItalicToggle(IRibbonControl control, bool pressed)
    {
        ((dynamic)ExcelDnaUtil.Application).ActiveCell.Font.Italic = pressed;
    }

    // -- dropDown "Fond" : couleur d'arriere-plan de la cellule active --
    // Index : 0 = Aucune, 1 = Jaune, 2 = Vert, 3 = Rouge.
    private static readonly int[] _colors =
    {
        0,        // Aucune (gere a part)
        0x00FFFF, // Jaune  (BGR)
        0x00B050, // Vert
        0x0000FF, // Rouge
    };

    public int GetColorIndex(IRibbonControl control) => 0;

    public void OnColorChange(IRibbonControl control, string selectedId, int selectedIndex)
    {
        dynamic interior = ((dynamic)ExcelDnaUtil.Application).ActiveCell.Interior;
        if (selectedIndex <= 0)
            interior.ColorIndex = -4142; // xlColorIndexNone : aucune couleur
        else
            interior.Color = _colors[selectedIndex];
    }

    // -- comboBox "Police" : applique la police saisie/choisie a la cellule active --
    public void OnFontChange(IRibbonControl control, string text)
    {
        if (!string.IsNullOrWhiteSpace(text))
            ((dynamic)ExcelDnaUtil.Application).ActiveCell.Font.Name = text;
    }
}
