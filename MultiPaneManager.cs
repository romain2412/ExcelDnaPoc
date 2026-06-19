using System.Collections.Generic;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : PLUSIEURS Custom Task Panes simultanes
// =====================================================================
// Rien ne limite Office/ExcelDna a un seul volet : CHAQUE appel a
// CustomTaskPaneFactory.CreateCustomTaskPane(...) cree un volet INDEPENDANT.
// Ici on en cree un nouveau a la demande (bouton "Nouveau volet") et on les garde
// dans une liste pour pouvoir TOUS les fermer (cycle de vie a notre charge).
//
// - Ancrage alterne Droite/Gauche : montre DockPosition et separe visuellement les
//   volets. Plusieurs volets du MEME cote s'empilent automatiquement (Office gere).
// - Un CTP est rattache a la fenetre Excel ACTIVE au moment de sa creation.
// - Delete() retire le volet et libere le COM sous-jacent.
public class MultiPaneManager
{
    private readonly List<CustomTaskPane> _panes = new();
    private int _counter;

    public int Count => _panes.Count;

    // Cree un nouveau volet, l'ancre, l'affiche, et le garde en reference.
    public void SpawnNew()
    {
        _counter++;
        bool right = _counter % 2 == 1;            // alterne le cote
        string dockLabel = right ? "Droite" : "Gauche";

        CustomTaskPane pane = CustomTaskPaneFactory.CreateCustomTaskPane(
            typeof(SpawnedPaneHost), $"POC Volet #{_counter}");
        pane.DockPosition = right
            ? MsoCTPDockPosition.msoCTPDockPositionRight
            : MsoCTPDockPosition.msoCTPDockPositionLeft;
        pane.Width = 260;

        ((SpawnedPaneHost)pane.ContentControl).SetInfo($"Volet #{_counter}", dockLabel);

        pane.Visible = true;
        _panes.Add(pane);
    }

    // Ferme et libere tous les volets crees ici (sans toucher au volet WPF).
    public void CloseAll()
    {
        foreach (CustomTaskPane pane in _panes)
        {
            try { pane.Delete(); }
            catch { /* deja ferme par l'utilisateur (croix) */ }
        }
        _panes.Clear();
        _counter = 0;
    }
}
