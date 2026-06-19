using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : API C bas niveau (XlCall + ExcelReference)
// =====================================================================
// La couche SOUS le modele COM (ExcelDnaUtil.Application) : l'interface native des
// .xll. On y descend pour ce que le COM fait mal ou pas du tout DEPUIS une UDF en
// calcul (contexte ou le modele COM est largement hors-jeu).
//
// - XlCall.Excel(id, args...) appelle une fonction de l'API C (ou une fonction
//   integree d'Excel).
// - ExcelReference = une reference de plage au niveau C, avec GetValue()/SetValue()
//   (lecture/ecriture sans passer par Range COM).
//
// IsMacroType = true est REQUIS : une UDF normale (type "fonction de feuille") n'a
// le droit d'acceder qu'a ses arguments. Pour appeler xlfCaller et LIRE d'autres
// cellules, il faut l'enregistrer en "macro sheet equivalent" (ce flag).
public static class FunctionsCApi
{
    [ExcelFunction(
        Name = "POC.OUSUISJE",
        Description = "Renvoie l'adresse de SA PROPRE cellule (via xlfCaller de l'API C).",
        Category = "POC ExcelDna API C",
        IsMacroType = true)]
    public static object OuSuisJe()
    {
        // xlfCaller : "qui m'a appele ?" -> une ExcelReference vers la cellule appelante.
        // Infaisable proprement via COM depuis une UDF.
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller)
            return "(pas appele depuis une cellule)";

        // ExcelReference porte les indices 0-based (RowFirst/ColumnFirst) + la feuille.
        // xlSheetNm donne le nom de feuille d'une reference (forme "[Classeur]Feuille").
        string feuille = (string)XlCall.Excel(XlCall.xlSheetNm, caller);
        string a1 = ColonneEnLettres(caller.ColumnFirst) + (caller.RowFirst + 1);
        return $"{feuille}!{a1}";
    }

    // Convertit un index de colonne 0-based en lettres Excel (0->A, 25->Z, 26->AA...).
    private static string ColonneEnLettres(int col0)
    {
        string s = "";
        for (int n = col0 + 1; n > 0; n = (n - 1) / 26)
            s = (char)('A' + (n - 1) % 26) + s;
        return s;
    }

    [ExcelFunction(
        Name = "POC.VALEURVOISINE",
        Description = "Lit la cellule a GAUCHE d'elle-meme (via ExcelReference de l'API C).",
        Category = "POC ExcelDna API C",
        IsMacroType = true)]
    public static object ValeurVoisine()
    {
        if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller)
            return "(pas appele depuis une cellule)";

        if (caller.ColumnFirst == 0)
            return ExcelError.ExcelErrorRef; // colonne A : pas de voisine a gauche -> #REF!

        // Reference vers la cellule immediatement a gauche, sur la MEME feuille (SheetId).
        var gauche = new ExcelReference(
            caller.RowFirst, caller.RowFirst,
            caller.ColumnFirst - 1, caller.ColumnFirst - 1,
            caller.SheetId);

        object valeur = gauche.GetValue(); // lecture directe, sans Range COM
        return valeur is ExcelEmpty ? "(voisine vide)" : valeur;
    }
}
