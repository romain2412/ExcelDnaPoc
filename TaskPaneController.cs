using System;
using ExcelDna.Integration.CustomUI;

namespace ExcelDnaPoc;

// =====================================================================
//  Concept ExcelDna : Custom Task Pane (volet ancre)
// =====================================================================
// Gere le cycle de vie du volet : creation paresseuse, visibilite, et acces au
// contenu WPF. Totalement decouple du ruban : il signale ses changements via
// l'evenement VisibleChanged (le ruban s'y abonne pour rafraichir son bouton).
//
// Chaine d'hebergement : CustomTaskPane -> WpfPaneHost (UserControl WinForms,
// ComVisible) -> ElementHost -> WpfPane (UserControl WPF/XAML).
public class TaskPaneController
{
    private CustomTaskPane? _pane;

    // Leve quand la visibilite du volet change (ex. fermeture via la croix).
    public event Action? VisibleChanged;

    public bool Visible => _pane?.Visible ?? false;

    // Cree le volet a la 1ere demande et renvoie son contenu WPF.
    private WpfPane EnsureCreated()
    {
        if (_pane == null)
        {
            _pane = CustomTaskPaneFactory.CreateCustomTaskPane(
                typeof(WpfPaneHost), "POC ExcelDna - Volet WPF");
            _pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
            _pane.Width = 360;
            _pane.VisibleStateChange += _ => VisibleChanged?.Invoke();
        }
        return ((WpfPaneHost)_pane.ContentControl).Pane;
    }

    // Affiche/masque le volet (le cree si besoin).
    public void SetVisible(bool visible)
    {
        EnsureCreated();
        _pane!.Visible = visible;
    }

    // Rend le volet visible et renvoie son contenu WPF (pour le piloter).
    public WpfPane Show()
    {
        WpfPane pane = EnsureCreated();
        _pane!.Visible = true;
        return pane;
    }
}
