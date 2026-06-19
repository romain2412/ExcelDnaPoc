Attribute VB_Name = "ChuckMacro"
' =====================================================================
'  Macro VBA qui utilise l'objet COM C# (ExcelDna) en FIRE-AND-FORGET
' =====================================================================
' L'objet COM ExcelDnaPoc.Chuck (classe ChuckCom du add-in) fait, de maniere
' ASYNCHRONE, l'appel a l'API Chuck Norris PUIS une attente de 15 s.
'
' Cette macro le DEMARRE et REND LA MAIN tout de suite : la methode COM retourne
' immediatement, la macro se termine, donc Excel reste REACTIF pendant tout l'IO.
' L'objet COM ecrit la blague lui-meme (cellule a droite de la cellule active)
' quand il a fini -- on n'attend pas ici (ni DoEvents, ni evenement).
'
' Pre-requis : l'add-in ExcelDnaPoc charge (il enregistre le serveur COM).
' Importer ce module dans un classeur (.xlsm), puis declencher via le bouton du
' ruban "Blague COM (VBA)" ou en lancant LancerBlagueAsync.
Public Sub LancerBlagueAsync()
    Dim chuck As Object
    Set chuck = CreateObject("ExcelDnaPoc.Chuck")
    chuck.LancerBlague   ' demarre l'async ; rend la main immediatement
End Sub
