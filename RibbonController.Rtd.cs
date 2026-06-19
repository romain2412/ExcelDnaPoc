using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// Fonctionnalite : groupe "RTD" — declenche le PUSH vers les cellules abonnees.
// "Inserer =POC.PUSH" reutilise le callback OnInsertFormula (cf. RibbonController.Udf.cs)
// via l'attribut tag ; "Pousser une valeur" appelle PushSource.Push -> le serveur
// RTD met a jour toutes les cellules =POC.PUSH("demo"). Aucun timer : tant qu'on ne
// clique pas, rien ne bouge.
// (Partie de la classe partielle RibbonController.)
public partial class RibbonController
{
    private const string RtdXml = """
        <group id='grpRtd' label='RTD'>
          <button id='btnRtdInsert' label='Inserer =POC.PUSH' size='large'
                  imageMso='TableInsert'
                  screentip='Insere =POC.PUSH(&quot;demo&quot;) dans la selection'
                  onAction='OnInsertFormula' tag='=POC.PUSH(&quot;demo&quot;)'/>
          <button id='btnRtdPush' label='Pousser une valeur' size='large'
                  imageMso='Refresh'
                  screentip='Declenche un push : toutes les cellules =POC.PUSH(&quot;demo&quot;) se mettent a jour'
                  onAction='OnPushRtd'/>
        </group>
        """;

    public void OnPushRtd(IRibbonControl control) => PushSource.Push("demo");
}
