using System.Collections.Generic;
using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : UDF (fonctions de feuille) =MaFonction(...)
// =====================================================================
// Les methodes marquees [ExcelFunction] sont enregistrees comme fonctions
// appelables depuis une cellule (avec ExcelAddInExplicitExports=true, seules
// celles marquees sont exportees). Name = nom dans la feuille, Description +
// Category + [ExcelArgument] alimentent l'Assistant de fonction (fx).
public static class Functions
{
    [ExcelFunction(
        Name = "POC.ADDITION",
        Description = "Additionne deux nombres.",
        Category = "POC ExcelDna")]
    public static double Addition(
        [ExcelArgument(Name = "a", Description = "le premier nombre")] double a,
        [ExcelArgument(Name = "b", Description = "le second nombre")] double b)
        => a + b;

    [ExcelFunction(
        Name = "POC.BONJOUR",
        Description = "Renvoie une salutation pour le nom donne.",
        Category = "POC ExcelDna")]
    public static string Bonjour(
        [ExcelArgument(Name = "nom", Description = "le nom a saluer (optionnel)")] string nom)
        => string.IsNullOrWhiteSpace(nom) ? "Bonjour tout le monde !" : $"Bonjour {nom} !";

    // ----- UDF prenant une PLAGE en parametre -----
    // Un parametre object[,] recoit le tableau 2D des valeurs de la plage :
    // nombre -> double, texte -> string, vide -> ExcelEmpty, erreur -> ExcelError.

    [ExcelFunction(
        Name = "POC.SOMMEPERSO",
        Description = "Additionne uniquement les nombres d'une plage.",
        Category = "POC ExcelDna")]
    public static double SommePerso(
        [ExcelArgument(Name = "plage", Description = "la plage de cellules a sommer")] object[,] plage)
    {
        double total = 0;
        foreach (object v in plage)
            if (v is double d)
                total += d;
        return total;
    }

    [ExcelFunction(
        Name = "POC.CONCAT",
        Description = "Concatene les valeurs non vides d'une plage.",
        Category = "POC ExcelDna")]
    public static string Concat(
        [ExcelArgument(Name = "plage", Description = "la plage de cellules")] object[,] plage,
        [ExcelArgument(Name = "separateur", Description = "separateur entre valeurs (defaut : \", \")")] object separateur)
    {
        // Argument optionnel : Excel passe ExcelMissing si omis.
        string sep = separateur is string s ? s : ", ";

        var parts = new List<string>();
        foreach (object v in plage)
            if (v is not null and not ExcelEmpty and not ExcelError and not ExcelMissing)
                parts.Add(v.ToString()!);

        return string.Join(sep, parts);
    }

    // ----- UDF qui RENVOIE une plage (tableau qui "spill") -----

    [ExcelFunction(
        Name = "POC.DOUBLER",
        Description = "Renvoie la plage avec chaque nombre multiplie par 2.",
        Category = "POC ExcelDna")]
    public static object[,] Doubler(
        [ExcelArgument(Name = "plage", Description = "la plage a doubler")] object[,] plage)
    {
        int rows = plage.GetLength(0);
        int cols = plage.GetLength(1);
        var result = new object[rows, cols];

        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                result[r, c] = plage[r, c] is double d ? d * 2 : plage[r, c];

        return result;
    }

    // ----- Variante double[,] : ExcelDna convertit la plage en nombres -----
    // Cellule VIDE -> 0 ; cellule NUMERIQUE -> sa valeur ; cellule TEXTE -> la conversion
    // echoue et la fonction renvoie #VALEUR! (la plage doit donc etre numerique/vide).
    // A l'inverse, object[,] (cf. SOMMEPERSO) accepte tout et laisse trier au code.

    [ExcelFunction(
        Name = "POC.MOYENNE",
        Description = "Moyenne d'une plage numerique (cellule vide = 0 ; du texte donne #VALEUR!).",
        Category = "POC ExcelDna")]
    public static double Moyenne(
        [ExcelArgument(Name = "plage", Description = "la plage de nombres")] double[,] plage)
    {
        double total = 0;
        int n = plage.Length;
        foreach (double v in plage)
            total += v;
        return n == 0 ? 0 : total / n;
    }

    // ----- Variante object[] : plage 1D (une seule ligne OU une seule colonne) -----

    [ExcelFunction(
        Name = "POC.COMPTENB",
        Description = "Compte les nombres d'une plage sur une ligne ou une colonne.",
        Category = "POC ExcelDna")]
    public static double CompteNb(
        [ExcelArgument(Name = "plage", Description = "une plage 1D (ligne ou colonne)")] object[] plage)
    {
        int count = 0;
        foreach (object v in plage)
            if (v is double)
                count++;
        return count;
    }

    // ----- Gestion FINE des types (nombres / textes / vides / erreurs) -----

    [ExcelFunction(
        Name = "POC.INFOPLAGE",
        Description = "Decrit le contenu d'une plage : nombres, textes, vides, erreurs.",
        Category = "POC ExcelDna")]
    public static string InfoPlage(
        [ExcelArgument(Name = "plage", Description = "la plage a analyser")] object[,] plage)
    {
        int nombres = 0, textes = 0, vides = 0, erreurs = 0, autres = 0;
        foreach (object v in plage)
        {
            if (v is double) nombres++;
            else if (v is string) textes++;
            else if (v is ExcelEmpty) vides++;
            else if (v is ExcelError) erreurs++;
            else autres++; // booleens, etc.
        }
        return $"{nombres} nombre(s), {textes} texte(s), {vides} vide(s), {erreurs} erreur(s)"
             + (autres > 0 ? $", {autres} autre(s)" : "");
    }

    // ----- UDF a PLUSIEURS plages : produit scalaire (renvoie #VALEUR! si tailles differentes) -----

    [ExcelFunction(
        Name = "POC.PRODUITSCALAIRE",
        Description = "Produit scalaire de deux plages de meme dimension (somme des produits).",
        Category = "POC ExcelDna")]
    public static object ProduitScalaire(
        [ExcelArgument(Name = "plage1", Description = "premiere plage de nombres")] double[,] plage1,
        [ExcelArgument(Name = "plage2", Description = "seconde plage de nombres (meme taille)")] double[,] plage2)
    {
        if (plage1.GetLength(0) != plage2.GetLength(0) ||
            plage1.GetLength(1) != plage2.GetLength(1))
            return ExcelError.ExcelErrorValue; // -> #VALEUR! dans la cellule

        double total = 0;
        int rows = plage1.GetLength(0), cols = plage1.GetLength(1);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                total += plage1[r, c] * plage2[r, c];

        return total;
    }
}
