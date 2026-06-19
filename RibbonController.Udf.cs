using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "UDF" — un bouton "UDF Sync" avec un menu deroulant.
// Chaque entree ecrit une formule UDF (exemple PRET A L'EMPLOI) dans la/les
// cellule(s) selectionnee(s). La formule est portee par l'attribut `tag` du
// <button> ; UN seul callback (OnInsertFormula) la lit et l'ecrit -> facile a
// etendre (ajouter un <button tag='=...'/>, zero code).
// Les exemples a plage utilisent des CONSTANTES MATRICIELLES ({10;20;30}) : ils
// marchent depuis n'importe quelle cellule (pas de dependance a des donnees voisines).
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    private const string UdfXml = """
        <group id='grpUdf' label='UDF'>
          <menu id='mnuUdf' label='UDF Sync' size='large' imageMso='ListMacros'
                screentip='Insere une formule UDF (exemple pret a l emploi) dans la selection'>
            <button id='udfAddition' label='POC.ADDITION(2;3)'
                    onAction='OnInsertFormula' tag='=POC.ADDITION(2,3)'/>
            <button id='udfBonjour' label='POC.BONJOUR(&quot;Romain&quot;)'
                    onAction='OnInsertFormula' tag='=POC.BONJOUR(&quot;Romain&quot;)'/>
            <menuSeparator id='udfSep1'/>
            <button id='udfSomme' label='POC.SOMMEPERSO({10;20;30})'
                    onAction='OnInsertFormula' tag='=POC.SOMMEPERSO({10;20;30})'/>
            <button id='udfConcat' label='POC.CONCAT({&quot;a&quot;;&quot;b&quot;;&quot;c&quot;};&quot; | &quot;)'
                    onAction='OnInsertFormula' tag='=POC.CONCAT({&quot;a&quot;;&quot;b&quot;;&quot;c&quot;},&quot; | &quot;)'/>
            <button id='udfMoyenne' label='POC.MOYENNE({1;2;3;4})'
                    onAction='OnInsertFormula' tag='=POC.MOYENNE({1;2;3;4})'/>
            <button id='udfCompte' label='POC.COMPTENB({1;2;&quot;x&quot;})'
                    onAction='OnInsertFormula' tag='=POC.COMPTENB({1;2;&quot;x&quot;})'/>
            <button id='udfInfo' label='POC.INFOPLAGE({1;&quot;x&quot;;2})'
                    onAction='OnInsertFormula' tag='=POC.INFOPLAGE({1;&quot;x&quot;;2})'/>
            <button id='udfDoubler' label='POC.DOUBLER({1;2;3})'
                    onAction='OnInsertFormula' tag='=POC.DOUBLER({1;2;3})'/>
            <button id='udfScalaire' label='POC.PRODUITSCALAIRE({1;2;3};{4;5;6})'
                    onAction='OnInsertFormula' tag='=POC.PRODUITSCALAIRE({1;2;3},{4;5;6})'/>
            <button id='udfChuckSync' label='POC.CHUCKNORRIS() (bloque Excel)'
                    onAction='OnInsertFormula' tag='=POC.CHUCKNORRIS()'/>
            <menuSeparator id='udfSep2' title='Async (non bloquant)'/>
            <button id='udfChuck' label='POC.CHUCKNORRISASYNC()'
                    onAction='OnInsertFormula' tag='=POC.CHUCKNORRISASYNC()'/>
            <button id='udfAttendre' label='POC.ATTENDRE(5)'
                    onAction='OnInsertFormula' tag='=POC.ATTENDRE(5)'/>
            <menuSeparator id='udfSep3' title='API C (XlCall / ExcelReference)'/>
            <button id='udfOuSuisJe' label='POC.OUSUISJE()'
                    onAction='OnInsertFormula' tag='=POC.OUSUISJE()'/>
            <button id='udfVoisine' label='POC.VALEURVOISINE()'
                    onAction='OnInsertFormula' tag='=POC.VALEURVOISINE()'/>
          </menu>
        </group>
        """;

    // Ecrit la formule (portee par le tag) dans la selection courante.
    // .Formula attend la syntaxe US : ',' separe les arguments, ';' separe les
    // lignes d'une constante matricielle (ex. {10;20;30} = une colonne de 3 valeurs).
    public void OnInsertFormula(IRibbonControl control)
    {
        try
        {
            dynamic selection = ((dynamic)ExcelDnaUtil.Application).Selection;
            selection.Formula = control.Tag;
        }
        catch (System.Exception ex)
        {
            Log.Error("OnInsertFormula a echoue", ex);
        }
    }
}
