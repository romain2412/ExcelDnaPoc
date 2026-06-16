using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ExcelDnaPoc;

// =====================================================================
//  Binding d'evenements Excel niveau WORKSHEET sur des callbacks C#
// =====================================================================
// Pour chaque FEUILLE des classeurs ouverts au demarrage, on s'abonne a ses
// evenements (Change, SelectionChange, Activate, BeforeDoubleClick, Calculate...).
// Meme technique que le clic droit / les events Workbook : points de connexion COM
// (sans interop Office) + sink IDispatch portant les bons DISPID (verifies par
// reflexion). Interface source = "DocEvents" (GUID 00024411-...).
public static class WorksheetEventsBinder
{
    private static readonly Guid DocEventsIid = new("00024411-0000-0000-C000-000000000046");

    private sealed class Binding
    {
        public required IConnectionPoint ConnectionPoint;
        public required int Cookie;
        public required WorksheetEventsSink Sink; // anti-GC
    }

    private static readonly List<Binding> _bindings = new();

    // Abonne les evenements C# sur une feuille (object COM Worksheet).
    public static void Bind(object worksheet, string workbookName)
    {
        try
        {
            string sheetName = (string)((dynamic)worksheet).Name;

            var cpc = (IConnectionPointContainer)worksheet;
            Guid iid = DocEventsIid;
            cpc.FindConnectionPoint(ref iid, out IConnectionPoint? cp);
            if (cp is null)
            {
                Log.Error($"WorksheetEvents: point de connexion introuvable ({workbookName}!{sheetName})");
                return;
            }

            var sink = new WorksheetEventsSink($"{workbookName}!{sheetName}");
            cp.Advise(sink, out int cookie);
            _bindings.Add(new Binding { ConnectionPoint = cp, Cookie = cookie, Sink = sink });
            Log.Info($"WorksheetEvents: bindes sur '{workbookName}!{sheetName}' (cookie={cookie})");
        }
        catch (Exception ex)
        {
            Log.Error("WorksheetEvents: bind echoue", ex);
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

// Interface des evenements Worksheet ("DocEvents") exposee en IDispatch.
[ComVisible(true)]
[Guid("00024411-0000-0000-C000-000000000046")] // = DocEvents (QI lors de l'Advise)
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IWorksheetEvents
{
    [DispId(1543)] void SelectionChange(object Target);
    [DispId(1545)] void Change(object Target);
    [DispId(304)]  void Activate();
    [DispId(1530)] void Deactivate();
    [DispId(1537)] void BeforeDoubleClick(object Target, ref bool Cancel);
    [DispId(279)]  void Calculate();
}

// Sink : Excel appelle ces callbacks C# via le point de connexion de la feuille.
[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
public class WorksheetEventsSink : IWorksheetEvents
{
    private readonly string _sheet; // "Classeur!Feuille"

    public WorksheetEventsSink(string sheet) => _sheet = sheet;

    public void SelectionChange(object Target)
    {
        try { Log.Info($"[{_sheet}] (ws) SelectionChange : {((dynamic)Target).Address}"); }
        catch (Exception ex) { Log.Error("ws SelectionChange", ex); }
    }

    public void Change(object Target)
    {
        try
        {
            string addr = (string)((dynamic)Target).Address;
            object? val = ((dynamic)Target).Value2;
            Log.Info($"[{_sheet}] (ws) Change : {addr} = '{val}'");
        }
        catch (Exception ex) { Log.Error("ws Change", ex); }
    }

    public void Activate() => Log.Info($"[{_sheet}] (ws) Activate");

    public void Deactivate() => Log.Info($"[{_sheet}] (ws) Deactivate");

    public void BeforeDoubleClick(object Target, ref bool Cancel)
    {
        try { Log.Info($"[{_sheet}] (ws) BeforeDoubleClick : {((dynamic)Target).Address} (Cancel={Cancel})"); }
        catch (Exception ex) { Log.Error("ws BeforeDoubleClick", ex); }
    }

    public void Calculate() => Log.Info($"[{_sheet}] (ws) Calculate");
}
