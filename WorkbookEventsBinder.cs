using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ExcelDnaPoc;

// =====================================================================
//  Binding d'evenements Excel niveau WORKBOOK sur des callbacks C#
// =====================================================================
// Pour chaque classeur ouvert au demarrage (StartupLoader), on s'abonne a ses
// evenements (SheetChange, SheetSelectionChange, Activate, BeforeClose...) qui
// declenchent des methodes C#. Meme technique que l'interception du clic droit :
// points de connexion COM (sans interop Office) + sink IDispatch portant les bons
// DISPID (verifies par reflexion). Extensible : ajouter un [DispId(...)] a l'interface.
public static class WorkbookEventsBinder
{
    // GUID de l'interface source d'evenements "WorkbookEvents" d'Excel.
    private static readonly Guid WorkbookEventsIid = new("00024412-0000-0000-C000-000000000046");

    private sealed class Binding
    {
        public required IConnectionPoint ConnectionPoint;
        public required int Cookie;
        public required WorkbookEventsSink Sink; // garde la reference (anti-GC)
    }

    private static readonly List<Binding> _bindings = new();

    // Abonne les evenements C# sur un classeur (object COM Workbook).
    public static void Bind(object workbook)
    {
        try
        {
            string name = (string)((dynamic)workbook).Name;

            var cpc = (IConnectionPointContainer)workbook;
            Guid iid = WorkbookEventsIid;
            cpc.FindConnectionPoint(ref iid, out IConnectionPoint? cp);
            if (cp is null)
            {
                Log.Error($"WorkbookEvents: point de connexion introuvable ({name})");
                return;
            }

            var sink = new WorkbookEventsSink(name);
            cp.Advise(sink, out int cookie);
            _bindings.Add(new Binding { ConnectionPoint = cp, Cookie = cookie, Sink = sink });
            Log.Info($"WorkbookEvents: bindes sur '{name}' (cookie={cookie})");
        }
        catch (Exception ex)
        {
            Log.Error("WorkbookEvents: bind echoue", ex);
        }
    }

    public static void UnbindAll()
    {
        foreach (Binding b in _bindings)
        {
            try { b.ConnectionPoint.Unadvise(b.Cookie); } catch { }
        }
        _bindings.Clear();
    }
}

// Interface d'evenements WorkbookEvents exposee en IDispatch. On ne declare que les
// evenements qui nous interessent ; les autres DISPID -> DISP_E_MEMBERNOTFOUND (ignores).
[ComVisible(true)]
[Guid("00024412-0000-0000-C000-000000000046")] // = WorkbookEvents (QI lors de l'Advise)
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IWorkbookEvents
{
    [DispId(1558)] void SheetSelectionChange(object Sh, object Target);
    [DispId(1564)] void SheetChange(object Sh, object Target);
    [DispId(304)]  void Activate();
    [DispId(1530)] void Deactivate();
    [DispId(1546)] void BeforeClose(ref bool Cancel);
}

// Sink : Excel appelle ces methodes (callbacks C#) via le point de connexion du classeur.
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public class WorkbookEventsSink : IWorkbookEvents
{
    private readonly string _workbook;

    public WorkbookEventsSink(string workbook) => _workbook = workbook;

    public void SheetSelectionChange(object Sh, object Target)
    {
        try
        {
            string sheet = (string)((dynamic)Sh).Name;
            string addr = (string)((dynamic)Target).Address;
            Log.Info($"[{_workbook}] SelectionChange : {sheet}!{addr}");
        }
        catch (Exception ex) { Log.Error("SheetSelectionChange", ex); }
    }

    public void SheetChange(object Sh, object Target)
    {
        try
        {
            string sheet = (string)((dynamic)Sh).Name;
            string addr = (string)((dynamic)Target).Address;
            object? val = ((dynamic)Target).Value2;
            Log.Info($"[{_workbook}] Change : {sheet}!{addr} = '{val}'");
        }
        catch (Exception ex) { Log.Error("SheetChange", ex); }
    }

    public void Activate() => Log.Info($"[{_workbook}] Activate");

    public void Deactivate() => Log.Info($"[{_workbook}] Deactivate");

    public void BeforeClose(ref bool Cancel)
        => Log.Info($"[{_workbook}] BeforeClose (Cancel={Cancel})");
}
