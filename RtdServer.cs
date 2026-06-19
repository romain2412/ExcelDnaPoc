using System.Collections.Generic;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.Rtd;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : serveur RTD (Real-Time Data)
// =====================================================================
// Des cellules qui se mettent a jour TOUTES SEULES quand le serveur pousse une
// valeur -- PUSH, pas polling. Contrairement a une UDF async (qui calcule une fois),
// un RTD reste "branche" et peut emettre de nouvelles valeurs indefiniment.
//
// Ici la source est PushSource (evenement .NET interne declenche par un bouton du
// ruban) : AUCUN timer. UpdateValue peut etre appele depuis n'importe quel thread,
// ExcelDna re-marshale vers Excel. (NB : Excel regroupe les updates RTD selon
// Application.RTD.ThrottleInterval -- 2s par defaut.)
//
// Cycle de vie (appele par Excel) :
//   ServerStart      -> s'abonner a la source
//   ConnectData      -> une cellule =POC.PUSH(sujet) s'abonne ; renvoyer la valeur initiale
//   (push)           -> UpdateValue sur les cellules du sujet
//   DisconnectData   -> la cellule disparait -> se desabonner
//   ServerTerminate  -> fermer la source
//
// [ComVisible(true)] requis (l'assembly est ComVisible(false) par defaut) ;
// le ProgId sert de cle dans XlCall.RTD (cf. RtdFunctions.Push).
[ComVisible(true)]
[ProgId(ServerProgId)]
public class PushRtdServer : ExcelRtdServer
{
    public const string ServerProgId = "ExcelDnaPoc.PushRtdServer";

    // sujet -> cellules abonnees (plusieurs cellules peuvent suivre le meme sujet).
    private readonly Dictionary<string, List<Topic>> _topicsBySubject = new();
    private readonly object _gate = new();

    protected override bool ServerStart()
    {
        PushSource.ValuePushed += OnValuePushed;
        return true;
    }

    protected override void ServerTerminate()
    {
        PushSource.ValuePushed -= OnValuePushed;
    }

    // topicInfo[0] = le 1er argument passe a RTD (ici le sujet).
    protected override object ConnectData(Topic topic, IList<string> topicInfo, ref bool newValues)
    {
        string sujet = topicInfo.Count > 0 ? topicInfo[0] : "";
        lock (_gate)
        {
            if (!_topicsBySubject.TryGetValue(sujet, out List<Topic>? list))
                _topicsBySubject[sujet] = list = new List<Topic>();
            list.Add(topic);
        }
        return PushSource.LastValue(sujet) ?? "(en attente d'un push...)";
    }

    protected override void DisconnectData(Topic topic)
    {
        lock (_gate)
        {
            foreach (List<Topic> list in _topicsBySubject.Values)
                if (list.Remove(topic)) break;
        }
    }

    // Notifie par PushSource : on pousse la valeur aux cellules abonnees a CE sujet.
    // Peut arriver de n'importe quel thread -> on copie la liste sous verrou.
    private void OnValuePushed(string sujet, object valeur)
    {
        List<Topic>? cibles;
        lock (_gate)
        {
            cibles = _topicsBySubject.TryGetValue(sujet, out List<Topic>? list)
                ? new List<Topic>(list)
                : null;
        }
        if (cibles == null) return;
        foreach (Topic t in cibles)
            t.UpdateValue(valeur); // push vers Excel
    }
}

// Fonction de feuille qui BRANCHE une cellule sur le serveur RTD : XlCall.RTD
// renvoie un handle "vivant" ; Excel re-affiche la cellule a chaque UpdateValue.
public static class RtdFunctions
{
    [ExcelFunction(
        Name = "POC.PUSH",
        Description = "Cellule RTD : derniere valeur poussee pour un sujet (mise a jour PUSH, sans timer).",
        Category = "POC ExcelDna RTD")]
    public static object Push(
        [ExcelArgument(Name = "sujet", Description = "le canal a suivre (ex. \"demo\")")] string sujet)
        => XlCall.RTD(PushRtdServer.ServerProgId, null, sujet);
}
