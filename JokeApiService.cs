using System;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : travail ASYNCHRONE + marshaling vers Excel
// =====================================================================
// Appel reseau non bloquant, attente async ANNULABLE (CancellationToken), puis
// ecriture dans Excel via ExcelAsyncUtil.QueueAsMacro (le modele objet Excel est
// STA mono-thread : on ne peut y toucher que sur le thread principal).
//
// Decouple du ruban : signale son etat "en cours" via RunningChanged (le ruban
// s'y abonne pour activer/desactiver le bouton Annuler). Le contenu WPF (WpfPane)
// est pilote pour afficher statut + barre de progression.
public class JokeApiService
{
    private const string JokeUrl = "https://api.chucknorris.io/jokes/random";

    // HttpClient partage (bonne pratique : un seul pour toute l'application).
    private static readonly HttpClient _http = new();

    private CancellationTokenSource? _cts;
    private bool _running;

    public bool IsRunning => _running;

    // Leve a chaque changement de l'etat "en cours".
    public event Action? RunningChanged;

    // Declenche l'annulation de l'operation en cours.
    public void Cancel() => _cts?.Cancel();

    private void SetRunning(bool running)
    {
        _running = running;
        RunningChanged?.Invoke();
    }

    // Operation asynchrone AWAITABLE ("async jusqu'en haut"). La gestion du token et de
    // l'etat "en cours" se fait SYNCHRONIQUEMENT au tout debut (avant le 1er await) ->
    // le bouton Annuler est actif immediatement. A appeler sur le thread principal
    // d'Excel ; le fire-and-forget est fait a la frontiere (ChuckTrigger.Fire).
    // Blague recuperee, PUIS attente 15s (annulable), PUIS ecriture (via QueueAsMacro).
    public async Task RunAsync(dynamic target, WpfPane pane)
    {
        // Annule une eventuelle operation precedente, puis demarre la nouvelle.
        _cts?.Cancel();
        var cts = new CancellationTokenSource();
        _cts = cts;
        SetRunning(true);
        CancellationToken ct = cts.Token;

        // 1) Appel reseau (duree inconnue) -> barre en mode indetermine.
        pane.SetBusy(true);
        pane.SetStatus("Appel de l'API Chuck Norris...");

        string joke;
        try
        {
            string json = await _http.GetStringAsync(JokeUrl, ct);
            using var doc = JsonDocument.Parse(json);
            joke = doc.RootElement.GetProperty("value").GetString() ?? "(reponse vide)";
        }
        catch (OperationCanceledException)
        {
            Finish(cts, target, pane, "Annule pendant l'appel API.", "(annule)", 0);
            return;
        }
        catch (Exception ex)
        {
            Finish(cts, target, pane, "Erreur API.", "Erreur d'appel API : " + ex.Message, 0);
            return;
        }

        // 2) Attente async de 15s, decoupee pour faire avancer la barre (compte a rebours).
        //    La blague n'est PAS encore ecrite. Le thread UI d'Excel reste libre.
        try
        {
            const int totalMs = 15_000;
            const int stepMs = 500;
            for (int elapsed = stepMs; elapsed <= totalMs; elapsed += stepMs)
            {
                await Task.Delay(stepMs, ct);
                int remaining = (totalMs - elapsed) / 1000;
                pane.SetProgress(elapsed * 100.0 / totalMs);
                pane.SetStatus($"Attente async... {remaining}s restantes (Annuler possible, Excel utilisable)");
            }
        }
        catch (OperationCanceledException)
        {
            Finish(cts, target, pane, "Attente annulee : blague non ecrite.", "(annule)", 0);
            return;
        }

        // 3) APRES l'attente : on ecrit enfin la blague dans la cellule.
        Finish(cts, target, pane, "Termine : blague ecrite apres 15s.", joke, 100);
    }

    // Centralise la fin d'operation : ecriture cellule (thread principal Excel),
    // mise a jour du volet, et reactualisation de l'etat "en cours".
    private void Finish(CancellationTokenSource cts, dynamic target, WpfPane pane,
                        string status, string? cellText, double progress)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            if (cellText != null)
                target.Value2 = cellText;

            // Ne touche l'etat global que si cette operation est toujours la courante
            // (evite qu'une operation annulee ne reinitialise une nouvelle en cours).
            if (ReferenceEquals(_cts, cts))
            {
                SetRunning(false);
                cts.Dispose();
                _cts = null; // IMPORTANT : sinon le prochain Start ferait Cancel() sur
                             // un CTS dispose -> ObjectDisposedException (=> "marche
                             // une seule fois").
            }
        });

        pane.SetProgress(progress);
        pane.SetStatus(status);
    }
}
