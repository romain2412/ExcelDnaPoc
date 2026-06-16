using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ExcelDna.Integration;

namespace ExcelDnaPoc;

// =====================================================================
//  SOLUTION 1 : INTERCEPTION du clic droit (menu reduit a la seule entree)
// =====================================================================
// Quand la cellule cliquee vaut "BlagueVisible", on ANNULE le menu contextuel
// par defaut d'Excel et on affiche un menu popup custom ne contenant QUE l'entree
// "Blague (async)". Sinon, le menu Excel habituel s'affiche.
//
// IMPORTANT : aucune dependance d'interop Office (le package interop traine une
// dependance 'office' introuvable au runtime sous .NET 8). On s'abonne donc a
// l'evenement Application.SheetBeforeRightClick via les POINTS DE CONNEXION COM
// (IConnectionPointContainer / IConnectionPoint, fournis par le BCL) avec un objet
// "sink" expose en IDispatch portant le bon DISPID (1560, verifie par reflexion).
// Lecture de la cellule + menu popup CommandBars : tout en late-binding (dynamic).
//
// Branche/debranche depuis le cycle de vie de l'add-in (AddIn.AutoOpen/AutoClose).
public static class CellRightClickInterceptor
{
    // Contenus de cellule qui declenchent l'interception (Solution 1).
    private const string TriggerWinforms = "BlagueMenuWinform"; // -> menu contextuel WinForms
    private const string TriggerWpf = "BlagueMenuWpf";          // -> menu contextuel WPF

    // Menus custom affiches a la place du menu Excel (references statiques anti-GC).
    private static System.Windows.Forms.ContextMenuStrip? _winMenu;
    private static System.Windows.Controls.ContextMenu? _wpfMenu;

    // GUID de l'interface source d'evenements "AppEvents" d'Excel.
    private static readonly Guid AppEventsIid = new("00024413-0000-0000-C000-000000000046");

    private static IConnectionPoint? _connectionPoint;
    private static int _cookie;
    private static ExcelAppEventsSink? _sink; // garde une reference (anti-GC)

    public static void Hook()
    {
        // Un echec ici ne doit JAMAIS empecher le chargement de l'add-in
        // (sinon xlAutoOpen echoue et Excel ne charge pas du tout l'add-in).
        try
        {
            Log.Info("Hook: cast IConnectionPointContainer");
            var cpc = (IConnectionPointContainer)ExcelDnaUtil.Application;
            Guid iid = AppEventsIid;
            Log.Info("Hook: FindConnectionPoint(AppEvents)");
            cpc.FindConnectionPoint(ref iid, out IConnectionPoint? cp);
            if (cp == null)
            {
                Log.Error("Hook: point de connexion introuvable");
                return;
            }

            _connectionPoint = cp;
            _sink = new ExcelAppEventsSink(OnRightClick);
            Log.Info("Hook: Advise");
            cp.Advise(_sink, out _cookie);
            Log.Info($"Hook: OK (cookie={_cookie})");
        }
        catch (Exception ex)
        {
            Log.Error("Hook a echoue", ex);
            _connectionPoint = null;
            _sink = null;
            _cookie = 0;
        }
    }

    public static void Unhook()
    {
        if (_connectionPoint != null && _cookie != 0)
            _connectionPoint.Unadvise(_cookie);
        _cookie = 0;
        _connectionPoint = null;
        _sink = null;
    }

    // Appele par le sink a chaque clic droit. Selon le contenu de la cellule, on ANNULE
    // tout le menu contextuel d'Excel (Cancel=true supprime meme les elements injectes
    // par Excel : barre de recherche, options de collage...) puis on affiche NOTRE
    // propre menu (WinForms ou WPF) ne contenant que l'entree "Blague (async)".
    private static void OnRightClick(object target, ref bool cancel)
    {
        try
        {
            object? value = ((dynamic)target).Cells[1, 1].Value2;
            string text = (value as string)?.Trim() ?? string.Empty;
            Log.Info($"OnRightClick: cellule = '{value}'");

            if (string.Equals(text, TriggerWinforms, StringComparison.OrdinalIgnoreCase))
            {
                cancel = true; // supprime l'integralite du menu Excel
                ExcelAsyncUtil.QueueAsMacro(ShowBlagueMenuWinforms);
                Log.Info("OnRightClick: menu Excel annule -> menu WinForms differe");
            }
            else if (string.Equals(text, TriggerWpf, StringComparison.OrdinalIgnoreCase))
            {
                cancel = true;
                ExcelAsyncUtil.QueueAsMacro(ShowBlagueMenuWpf);
                Log.Info("OnRightClick: menu Excel annule -> menu WPF differe");
            }
        }
        catch (Exception ex)
        {
            Log.Error("OnRightClick a echoue", ex); // ne jamais casser le clic droit
        }
    }

    // Menu contextuel WINFORMS ne contenant QUE l'entree "Blague (async)".
    private static void ShowBlagueMenuWinforms()
    {
        try
        {
            _winMenu?.Dispose();

            var menu = new System.Windows.Forms.ContextMenuStrip();
            var item = menu.Items.Add("Blague (async)");
            item.Click += (_, _) => ChuckTrigger.Run();
            _winMenu = menu; // anti-GC

            menu.Show(System.Windows.Forms.Cursor.Position);
            Log.Info("ShowBlagueMenuWinforms: affiche");
        }
        catch (Exception ex)
        {
            Log.Error("ShowBlagueMenuWinforms a echoue", ex);
        }
    }

    // Menu contextuel WPF ne contenant QUE l'entree "Blague (async)".
    private static void ShowBlagueMenuWpf()
    {
        try
        {
            var menu = new System.Windows.Controls.ContextMenu
            {
                Placement = System.Windows.Controls.Primitives.PlacementMode.MousePoint,
            };
            var item = new System.Windows.Controls.MenuItem { Header = "Blague (async)" };
            item.Click += (_, _) => ChuckTrigger.Run();
            menu.Items.Add(item);
            _wpfMenu = menu; // anti-GC

            menu.IsOpen = true; // affichage standalone au curseur
            Log.Info("ShowBlagueMenuWpf: affiche");
        }
        catch (Exception ex)
        {
            Log.Error("ShowBlagueMenuWpf a echoue", ex);
        }
    }
}

// Interface d'evenements exposee en IDispatch. Seul SheetBeforeRightClick (DISPID 1560)
// est declare : Excel appellera Invoke pour les autres evenements AppEvents, ils seront
// simplement ignores (DISP_E_MEMBERNOTFOUND), ce qui est sans consequence.
[ComVisible(true)]
[Guid("00024413-0000-0000-C000-000000000046")] // = AppEvents : le point de connexion d'Excel
                                               // QI le sink sur cet IID lors de l'Advise.
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IExcelAppEvents
{
    [DispId(1560)]
    void SheetBeforeRightClick(object Sh, object Target, ref bool Cancel);
}

// Objet "sink" : Excel appelle SheetBeforeRightClick via son point de connexion AppEvents.
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public class ExcelAppEventsSink : IExcelAppEvents
{
    public delegate void RightClickHandler(object target, ref bool cancel);

    private readonly RightClickHandler _onRightClick;

    public ExcelAppEventsSink(RightClickHandler onRightClick) => _onRightClick = onRightClick;

    public void SheetBeforeRightClick(object Sh, object Target, ref bool Cancel)
        => _onRightClick(Target, ref Cancel);
}
