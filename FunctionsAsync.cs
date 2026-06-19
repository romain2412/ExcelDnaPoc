using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : UDF ASYNCHRONE (fonction de feuille non bloquante)
// =====================================================================
// Une UDF normale (cf. Functions.cs) s'execute sur le thread de calcul d'Excel :
// tant qu'elle tourne, Excel est FIGE. Inacceptable pour un I/O reseau.
//
// ExcelDna offre DEUX APIs pour rendre la main a Excel. Elles different sur un
// point essentiel : tiennent-elles un thread pendant l'attente ?
//
//   (A) ExcelAsyncUtil.Run(nom, params, Func<object>)  --  SIMPLE mais SYNC-OVER-ASYNC
//       Le delegue est SYNCHRONE. ExcelDna le lance sur un thread du ThreadPool,
//       rend #N/A a Excel, puis ecrit le retour. Excel n'est plus bloque, MAIS le
//       delegue doit bloquer SON thread (.GetAwaiter().GetResult()) pendant tout
//       l'I/O -> un thread du pool gaspille a ne rien faire. On a juste DEPLACE le
//       blocage hors d'Excel ; ce n'est pas la philosophie de l'async I/O.
//
//   (B) ExcelAsyncUtil.Observe(nom, params, () => IExcelObservable)  --  VRAI ASYNC
//       On fournit un IExcelObservable. On y BRANCHE une Task (adaptateur
//       ExcelTaskObservable ci-dessous). Le `await _http.GetStringAsync(...)` se
//       termine via le port de completion d'I/O : AUCUN thread n'est tenu pendant
//       l'attente reseau. A la completion, l'observateur pousse la valeur a Excel.
//       C'est l'API utilisee ici. (RTD-based en interne, mais zero blocage.)
//
// Points communs aux deux : Excel recoit #N/A immediatement puis la cellule est
// mise a jour ; l'IDENTITE de l'appel = (nom + parametres) sert de cle (dedup si
// memes params ; changer un param -- d'ou `actualiser` -- ou F9 force un nouvel
// appel). Async natif `Task<T>` directement sur [ExcelFunction] = encore une autre
// voie, qui demande le package ExcelDna.Registration (ProcessAsyncRegistrations) ;
// non retenu ici pour rester sans dependance.
//
// Difference avec RTD : une UDF async calcule UNE fois (puis se reactualise sur
// recalcul/changement d'entree). Un rafraichissement TOUT SEUL sur timer = RTD
// (ExcelRtdServer) -- concept distinct.
public static class FunctionsAsync
{
    private const string JokeUrl = "https://api.chucknorris.io/jokes/random";
    private static readonly HttpClient _http = new();

    // Le pendant SYNCHRONE (le contre-exemple qui BLOQUE Excel) est POC.CHUCKNORRIS,
    // range avec les autres UDF synchrones dans Functions.cs.

    // ----- VRAI async : Observe + Task, aucun thread tenu pendant l'I/O -----
    [ExcelFunction(
        Name = "POC.CHUCKNORRISASYNC",
        Description = "Appelle l'API Chuck Norris sans bloquer Excel NI aucun thread pendant l'attente reseau.",
        Category = "POC ExcelDna Async")]
    public static object ChuckNorrisAsync(
        [ExcelArgument(Name = "actualiser",
            Description = "valeur quelconque ; la changer (ou F9) force une nouvelle blague (identite = nom + parametres)")]
        object actualiser)
    {
        // La source n'est appelee qu'une fois par identite (nom + params). Elle
        // DEMARRE la Task et la branche sur l'observable : le retour est immediat,
        // l'attente reseau ne tient aucun thread.
        return ExcelAsyncUtil.Observe(nameof(ChuckNorrisAsync), new[] { actualiser },
            () => FetchJokeAsync().ToExcelObservable());
    }

    private static async Task<string> FetchJokeAsync()
    {
        try
        {
            string json = await _http.GetStringAsync(JokeUrl); // VRAI await : pas de thread bloque
            using var doc = JsonDocument.Parse(json);
            return doc.RootElement.GetProperty("value").GetString() ?? "(reponse vide)";
        }
        catch (Exception ex)
        {
            return "Erreur API : " + ex.Message;
        }
    }

    [ExcelFunction(
        Name = "POC.ATTENDRE",
        Description = "Attend N secondes sans bloquer aucun thread (Task.Delay) : rend visible le passage #N/A -> valeur.",
        Category = "POC ExcelDna Async")]
    public static object Attendre(
        [ExcelArgument(Name = "secondes", Description = "duree d'attente simulee (defaut : 3)")] object secondes)
    {
        int sec = secondes is double d ? (int)d : 3;
        if (sec < 0) sec = 0;
        if (sec > 30) sec = 30;

        return ExcelAsyncUtil.Observe(nameof(Attendre), new object[] { sec },
            () => DelayAsync(sec).ToExcelObservable());
    }

    private static async Task<string> DelayAsync(int sec)
    {
        await Task.Delay(sec * 1000); // timer, pas de Thread.Sleep -> aucun thread tenu
        return $"Attendu {sec}s (vrai async : aucun thread tenu pendant l'attente)";
    }
}

// ---------------------------------------------------------------------
//  Pont Task -> IExcelObservable (ce qui permet le VRAI async cote ExcelDna)
// ---------------------------------------------------------------------
// ExcelAsyncUtil.Observe attend un IExcelObservable. Cet adaptateur abonne
// Excel a la fin d'une Task : a la completion, il pousse le resultat (OnNext +
// OnCompleted) ou l'erreur (OnError). Aucun polling, aucun thread d'attente :
// c'est la continuation de la Task qui notifie Excel.
internal sealed class ExcelTaskObservable<T> : IExcelObservable
{
    private readonly Task<T> _task;
    public ExcelTaskObservable(Task<T> task) => _task = task;

    public IDisposable Subscribe(IExcelObserver observer)
    {
        _task.ContinueWith(t =>
        {
            if (t.IsFaulted)
                observer.OnError(t.Exception?.InnerException ?? t.Exception!);
            else if (t.IsCanceled)
                observer.OnError(new TaskCanceledException(t));
            else
            {
                observer.OnNext(t.Result!);
                observer.OnCompleted();
            }
        }, TaskScheduler.Default);

        return NullDisposable.Instance; // rien a desabonner : la Task notifie une fois
    }

    private sealed class NullDisposable : IDisposable
    {
        public static readonly NullDisposable Instance = new();
        public void Dispose() { }
    }
}

internal static class ExcelObservableExtensions
{
    public static IExcelObservable ToExcelObservable<T>(this Task<T> task)
        => new ExcelTaskObservable<T>(task);
}
