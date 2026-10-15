using System;
using System.Collections.Generic;
using System.Threading;
using OpenSlideCharts;

internal static class PowerPointEventsTests
{
    private static int checks;
    private static void Check(bool value, string reason)
    {
        ++checks;
        if (!value) throw new Exception(reason);
    }

    private sealed class Registrar
    {
        internal object App;
        internal readonly Dictionary<int, Delegate> Handlers = new Dictionary<int, Delegate>();
        internal readonly List<int> Removed = new List<int>();
        internal int FailAttach;
        internal int FailRemove;
        internal void Add(object app, Guid iid, int dispid, Delegate handler)
        {
            Check(Object.ReferenceEquals(app, App), "Application identity changed.");
            Check(iid == new Guid("914934C2-5A91-11CF-8700-00AA0060263B"), "Event IID changed.");
            if (dispid == FailAttach) throw new InvalidOperationException("attach failure");
            Handlers.Add(dispid, handler);
        }
        internal void Remove(object app, Guid iid, int dispid, Delegate handler)
        {
            Check(Object.ReferenceEquals(Handlers[dispid], handler), "Delegate identity changed during removal.");
            Removed.Add(dispid);
            if (dispid == FailRemove) throw new InvalidOperationException("remove failure");
            Handlers.Remove(dispid);
        }
    }

    private static PowerPointEvents Attach(Registrar registrar, Action<object> selected, BeforeRightClick right,
        BeforePresentationAction save, BeforePresentationAction close, Action<object, object> deactivate)
    {
        return new PowerPointEvents(registrar.App, selected, right, save, close, deactivate, registrar.Add, registrar.Remove);
    }

    public static int Main()
    {
        var registrar = new Registrar { App = new object() };
        object selection = new object(), presentation = new object(), window = new object();
        int selections = 0, clicks = 0, saves = 0, closes = 0, deactivations = 0;
        var bridge = Attach(registrar,
            delegate(object value) { Check(Object.ReferenceEquals(value, selection), "Selection identity lost."); selections++; },
            delegate(object value, ref bool cancel) { Check(Object.ReferenceEquals(value, selection), "Right-click identity lost."); clicks++; cancel = true; },
            delegate(object value, ref bool cancel) { Check(Object.ReferenceEquals(value, presentation), "Save identity lost."); saves++; cancel = true; },
            delegate(object value, ref bool cancel) { Check(cancel, "Incoming cancel lost."); closes++; cancel = false; },
            delegate(object p, object w) { Check(Object.ReferenceEquals(p, presentation) && Object.ReferenceEquals(w, window), "Window context lost."); deactivations++; });
        Check(registrar.Handlers.Count == 5, "Expected all requested events.");
        registrar.Handlers[2001].DynamicInvoke(selection);
        object[] rightArgs = { selection, false };
        registrar.Handlers[2002].DynamicInvoke(rightArgs);
        Check((bool)rightArgs[1], "Right-click ref cancel was not written back.");
        object[] saveArgs = { presentation, false };
        registrar.Handlers[2018].DynamicInvoke(saveArgs);
        Check((bool)saveArgs[1], "BeforeSave ref cancel was not written back.");
        object[] closeArgs = { presentation, true };
        registrar.Handlers[2025].DynamicInvoke(closeArgs);
        Check(!(bool)closeArgs[1], "BeforeClose ref cancel was not written back.");
        registrar.Handlers[2010].DynamicInvoke(presentation, window);
        Check(selections == 1 && clicks == 1 && saves == 1 && closes == 1 && deactivations == 1, "Callbacks not dispatched once.");
        Delegate staleRight = registrar.Handlers[2002];
        Exception wrongThreadError = null;
        Thread thread = new Thread(delegate() { try { bridge.Dispose(); } catch (Exception e) { wrongThreadError = e; } });
        thread.Start(); thread.Join();
        Check(wrongThreadError is InvalidOperationException, "Wrong-thread cleanup was accepted.");
        Check(registrar.Handlers.Count == 5, "Wrong-thread cleanup removed handlers.");
        registrar.FailRemove = 2002;
        try { bridge.Dispose(); throw new Exception("Expected detach failure."); }
        catch (AggregateException e) { Check(e.InnerExceptions.Count == 1, "Unexpected detach error count."); }
        Check(registrar.Handlers.Count == 1 && registrar.Handlers.ContainsKey(2002), "Dispose must try every removal.");
        object[] staleArgs = { selection, false };
        staleRight.DynamicInvoke(staleArgs);
        Check(clicks == 1 && !(bool)staleArgs[1], "Stale callback ran during failed disposal.");
        registrar.FailRemove = 0;
        bridge.Dispose(); bridge.Dispose();
        Check(registrar.Handlers.Count == 0, "Retry did not detach failed subscription.");

        registrar = new Registrar { App = new object(), FailAttach = 2018 };
        try
        {
            Attach(registrar, delegate(object x) { }, delegate(object x, ref bool c) { }, delegate(object x, ref bool c) { }, null, null);
            throw new Exception("Expected attachment failure.");
        }
        catch (InvalidOperationException e) { Check(e.Message == "attach failure", "Original attachment error lost."); }
        Check(registrar.Handlers.Count == 0 && registrar.Removed.Count == 2, "Partial attach not rolled back.");

        int partialCalls = 0;
        registrar = new Registrar { App = new object(), FailAttach = 2018, FailRemove = 2001 };
        try
        {
            Attach(registrar, delegate(object x) { partialCalls++; }, delegate(object x, ref bool c) { }, delegate(object x, ref bool c) { }, null, null);
            throw new Exception("Expected aggregate attachment failure.");
        }
        catch (AggregateException e) { Check(e.InnerExceptions.Count == 2, "Rollback failure swallowed original error."); }
        Check(registrar.Handlers.Count == 1, "Rollback did not attempt every handler.");
        registrar.Handlers[2001].DynamicInvoke(selection);
        Check(partialCalls == 0, "Partial attachment unexpectedly activated callbacks.");
        try { new PowerPointEvents(null, delegate(object x) { }, null); throw new Exception("Expected null guard."); }
        catch (ArgumentNullException) { checks++; }
        try { new PowerPointEvents(new object(), delegate(object x) { }, null); throw new Exception("Expected COM guard."); }
        catch (ArgumentException) { checks++; }
        Console.WriteLine("PowerPointEvents managed lifecycle/ref-argument tests: " + checks + " checks passed. No Office instance launched; COM transport remains a live-host check.");
        return 0;
    }
}
