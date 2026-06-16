using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  Un TRAITEMENT asynchrone INDEPENDANT (parallelisable + annulable)
// =====================================================================
// Chaque declenchement (bouton, menu, double-clic...) cree un JokeJob a part :
// son propre CancellationTokenSource, sa cible (cellule), et il rapporte sa
// progression via des callbacks (branches sur sa propre ligne d'UI - JobRow).
// Plusieurs jobs tournent ainsi EN PARALLELE, chacun annulable separement.
public class JokeJob
{
    private const string JokeUrl = "https://api.chucknorris.io/jokes/random";
    private static readonly HttpClient _http = new();

    private readonly CancellationTokenSource _cts = new();
    private readonly dynamic _target; // cellule ou ecrire la blague

    // Callbacks de progression (branches par WpfPane sur la JobRow ; thread-safe cote UI).
    public Action<string>? OnStatus;
    public Action<double>? OnProgress;
    public Action<bool>? OnBusy;
    public Action? OnDone;

    public JokeJob(dynamic target) => _target = target;

    public void Cancel()
    {
        try { _cts.Cancel(); } catch (ObjectDisposedException) { }
    }

    public async Task RunAsync()
    {
        CancellationToken ct = _cts.Token;

        OnBusy?.Invoke(true);
        OnStatus?.Invoke("Appel de l'API Chuck Norris...");

        string joke;
        try
        {
            string json = await _http.GetStringAsync(JokeUrl, ct);
            using var doc = JsonDocument.Parse(json);
            joke = doc.RootElement.GetProperty("value").GetString() ?? "(reponse vide)";
        }
        catch (OperationCanceledException)
        {
            Finish("(annule)", "Annule pendant l'appel API.");
            return;
        }
        catch (Exception ex)
        {
            Finish("Erreur d'appel API : " + ex.Message, "Erreur API.");
            return;
        }

        // Attente de 15s, decoupee pour faire avancer la barre (compte a rebours).
        try
        {
            const int totalMs = 15_000;
            const int stepMs = 500;
            for (int elapsed = stepMs; elapsed <= totalMs; elapsed += stepMs)
            {
                await Task.Delay(stepMs, ct);
                int remaining = (totalMs - elapsed) / 1000;
                OnProgress?.Invoke(elapsed * 100.0 / totalMs);
                OnStatus?.Invoke($"Attente... {remaining}s restantes");
            }
        }
        catch (OperationCanceledException)
        {
            Finish("(annule)", "Attente annulee : blague non ecrite.");
            return;
        }

        OnProgress?.Invoke(100);
        Finish(joke, "Termine : blague ecrite.");
    }

    // Ecriture de la cellule sur le thread principal d'Excel (QueueAsMacro) + fin d'UI.
    private void Finish(string? cellText, string status)
    {
        if (cellText != null)
            ExcelAsyncUtil.QueueAsMacro(() => _target.Value2 = cellText);

        OnStatus?.Invoke(status);
        OnDone?.Invoke();
    }
}
