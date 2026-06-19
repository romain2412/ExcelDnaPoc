using System;
using System.Collections.Generic;

namespace ExcelDnaPoc;

// =====================================================================
//  "Flux" event-driven INTERNE (ni timer, ni reseau)
// =====================================================================
// Demontre le mecanisme PUSH du RTD a l'etat pur : RIEN ne bouge tant qu'aucun
// evenement n'arrive. Un bouton du ruban appelle Push(sujet) -> l'evenement
// ValuePushed se declenche -> le serveur RTD abonne (PushRtdServer) pousse la
// valeur aux cellules concernees.
//
// Le point a retenir : une vraie source temps reel (WebSocket, stream gRPC, file
// de messages) remplacerait SEULEMENT l'appelant de Push (webSocket.OnMessage,
// `await foreach` gRPC...) -- le serveur RTD, lui, ne changerait pas d'une ligne.
public static class PushSource
{
    // (sujet, valeur). Le serveur RTD s'y abonne dans ServerStart.
    public static event Action<string, object>? ValuePushed;

    // Derniere valeur par sujet : un nouvel abonne (ConnectData) recoit l'etat courant.
    private static readonly Dictionary<string, object> _last = new();
    private static readonly object _gate = new();
    private static int _counter;

    // Declenche un "push" pour un sujet (appele par le bouton ruban).
    public static void Push(string sujet)
    {
        object valeur;
        lock (_gate)
        {
            _counter++;
            valeur = $"push #{_counter} @ {DateTime.Now:HH:mm:ss}";
            _last[sujet] = valeur;
        }
        ValuePushed?.Invoke(sujet, valeur); // notification -> serveur RTD
    }

    // Etat courant d'un sujet (null si rien n'a encore ete pousse).
    public static object? LastValue(string sujet)
    {
        lock (_gate) return _last.TryGetValue(sujet, out object? v) ? v : null;
    }
}
