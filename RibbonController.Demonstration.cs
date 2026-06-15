using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "Demonstration" — commandes simples sur la feuille.
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    private const string DemonstrationXml = """
        <group id='grpDemo' label='Demonstration'>
          <button id='btnHello' label='Dire Bonjour' size='large'
                  imageMso='HappyFace' onAction='OnHelloClick'/>
          <button id='btnFill' label='Ecrire dans la cellule' size='large'
                  imageMso='FileSave' onAction='OnFillClick'/>
          <button id='btnSum' label='Somme selection' size='large'
                  imageMso='AutoSum' onAction='OnSumClick'/>
        </group>
        """;

    // Affiche une boite de dialogue.
    public void OnHelloClick(IRibbonControl control)
    {
        MessageBox.Show(
            "Bonjour depuis un add-in ExcelDna en .NET 8 !\nAucun code VBA n'est utilise.",
            "POC ExcelDna",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    // Ecrit une valeur dans la cellule active via le modele objet Excel.
    public void OnFillClick(IRibbonControl control)
    {
        dynamic app = ExcelDnaUtil.Application;
        app.ActiveCell.Value2 = $"Ecrit par l'add-in .NET a {System.DateTime.Now:HH:mm:ss}";
    }

    // Calcule la somme de la selection et l'affiche.
    public void OnSumClick(IRibbonControl control)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic selection = app.Selection;
        double total = 0;
        foreach (dynamic cell in selection.Cells)
        {
            if (cell.Value2 is double d)
                total += d;
        }

        MessageBox.Show(
            $"Somme des cellules selectionnees : {total}",
            "POC ExcelDna",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }
}
