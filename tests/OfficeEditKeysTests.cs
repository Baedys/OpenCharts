using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using OpenSlideCharts;

public static class OfficeEditKeysTests
{
    private static int scenarios, assertions;
    private const long Down = 1, Repeat = 0x40000001L, Up = 0xC0000001L;
    [STAThread]
    public static int Main()
    {
        try
        {
            Installation(); Filtering(); Modifiers(); OwnedPress(); Preview(); DeclinedPress();
            TwoKeys(); SelfDispose(); LateDelivery(); WrongThread(); PredicateFailure(); HandlerFailure();
            PostingFailure(); InstallFailure(); UnhookFailure(); DelegateRoot(); WrongCookie(); ReentrantPump();
            InvalidArguments(); NativeMenus(); GuiQueryFailure(); RealThreadHook();
            Console.WriteLine("PASS: " + scenarios + " Office edit-key scenarios, " + assertions + " assertions");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }

    private sealed class FakeNative : OfficeEditKeys.NativeApi
    {
        internal WeakReference Callback;
        internal int Kind, InstallCalls, UnhookCalls, NextCalls, StateCalls, PostCalls;
        internal int Modifiers, Error = 1234, PostFailures, UnhookFailures;
        internal uint MenuFlags;
        internal bool FailInstall, ZeroThread, FailGui;
        internal uint InstalledThread;
        internal IntPtr Module, InstalledWindow, Unhooked, LastWindow, LastCookie;
        internal int LastMessage;
        internal readonly IntPtr Handle = new IntPtr(4567);
        internal override uint CurrentThreadId() { return ZeroThread ? 0 : base.CurrentThreadId(); }
        internal override IntPtr InstallHook(int kind, OfficeEditKeys.HookProc callback, IntPtr module, uint threadId)
        {
            InstallCalls++; Kind = kind; Callback = new WeakReference(callback); Module = module; InstalledThread = threadId;
            InstalledWindow = ((OfficeEditKeys)callback.Target).MessageHandle;
            return FailInstall ? IntPtr.Zero : Handle;
        }
        internal override bool Unhook(IntPtr handle)
        { UnhookCalls++; Unhooked = handle; if (UnhookFailures > 0) { UnhookFailures--; return false; } return true; }
        internal override IntPtr CallNext(IntPtr handle, int code, IntPtr key, IntPtr flags)
        { NextCalls++; return new IntPtr(77); }
        internal override short KeyState(int key)
        { StateCalls++; return (Modifiers & (key == 16 ? 1 : key == 17 ? 2 : 4)) != 0 ? unchecked((short)0x8000) : (short)0; }
        internal override uint GuiFlags(uint threadId) { Assert(threadId == InstalledThread, "GUI query uses creation thread"); if (FailGui) throw new Win32Exception(Error); return MenuFlags; }
        internal override bool Post(IntPtr handle, int message, IntPtr wParam, IntPtr lParam)
        {
            PostCalls++; LastWindow = handle; LastMessage = message; LastCookie = wParam;
            if (PostFailures > 0) { PostFailures--; return false; }
            return base.Post(handle, message, wParam, lParam);
        }
        internal override int LastError() { return Error; }
        internal IntPtr Fire(int key, long flags, int code = 0)
        { return ((OfficeEditKeys.HookProc)Callback.Target)(code, new IntPtr(key), new IntPtr(unchecked((int)flags))); }
    }
    private static void Assert(bool condition, string label)
    { assertions++; if (!condition) throw new Exception(label); }
    private static T Throws<T>(Action action, string label) where T : Exception
    { try { action(); } catch (T error) { assertions++; return error; } throw new Exception(label); }
    private static void Consumed(IntPtr result) { Assert(result == new IntPtr(1), "owned event consumed"); }
    private static void Passed(IntPtr result) { Assert(result == new IntPtr(77), "unowned event forwarded"); }
    private static void Pump() { Application.DoEvents(); }
    private static void Installation()
    {
        scenarios++; FakeNative api = new FakeNative();
        using (OfficeEditKeys guard = new OfficeEditKeys(k => true, k => { }, null, api))
        {
            Assert(api.Kind == 2 && api.Module == IntPtr.Zero, "WH_KEYBOARD module NULL");
            Assert(api.InstalledThread == api.CurrentThreadId() && api.InstalledThread != 0, "only current native thread");
            Assert(IsWindow(guard.MessageHandle), "message-only window exists");
            Assert(FindWindowExW(new IntPtr(-3), IntPtr.Zero, null, "OpenCharts edit key dispatch") == guard.MessageHandle, "HWND_MESSAGE discovery");
        }
        Assert(api.UnhookCalls == 1 && api.Unhooked == api.Handle, "exact installed hook removed");
        Assert(!IsWindow(api.InstalledWindow), "window destroyed");
    }
    private static void Filtering()
    {
        scenarios++; FakeNative api = new FakeNative(); int predicates = 0, actions = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => { predicates++; return true; }, k => actions++, null, api))
        {
            Passed(api.Fire(65, Down)); Passed(api.Fire(27, Down, -1)); Passed(api.Fire(27, Down, 2));
            Passed(api.Fire(27, Up)); Passed(api.Fire(46, Repeat)); Pump();
            Assert(predicates == 0 && actions == 0 && api.PostCalls == 0 && api.StateCalls == 0, "filter before predicate/modifier lookup");
        }
    }
    private static void Modifiers()
    {
        scenarios++; FakeNative api = new FakeNative(); int predicates = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => { predicates++; return true; }, k => { }, null, api))
        {
            foreach (int modifier in new int[] { 1, 2, 4, 7 }) { api.Modifiers = modifier; Passed(api.Fire(27, Down)); Passed(api.Fire(46, Down)); }
            api.Modifiers = 0; Passed(api.Fire(27, Down | 0x20000000L));
            Assert(predicates == 0 && api.PostCalls == 0, "modified keys never claimed");
        }
    }
    private static void OwnedPress()
    {
        scenarios++; FakeNative api = new FakeNative(); int predicates = 0, actions = 0, thread = Thread.CurrentThread.ManagedThreadId;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => { predicates++; Assert(Thread.CurrentThread.ManagedThreadId == thread, "predicate on creator"); return true; }, k => { actions++; Assert(k == 27 && Thread.CurrentThread.ManagedThreadId == thread, "action on creator"); }, null, api))
        {
            Consumed(api.Fire(27, Down)); Assert(actions == 0, "hook never calls action");
            Consumed(api.Fire(27, Repeat)); Consumed(api.Fire(27, Down));
            api.Modifiers = 7; Consumed(api.Fire(27, Repeat)); Consumed(api.Fire(27, Up));
            Assert(api.PostCalls == 1 && predicates == 1, "one queued action per owned press"); Pump(); Assert(actions == 1, "release before dispatch preserves action");
            api.Modifiers = 0; Consumed(api.Fire(27, Down)); Consumed(api.Fire(27, Up)); Pump(); Assert(actions == 2 && predicates == 2, "next press independent");
        }
    }
    private static void Preview()
    {
        scenarios++; FakeNative api = new FakeNative(); int actions = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => true, k => actions++, null, api))
        {
            Passed(api.Fire(27, Down, 3)); Assert(api.PostCalls == 0, "PM_NOREMOVE preview does not dispatch");
            Consumed(api.Fire(27, Down)); Consumed(api.Fire(27, Up, 3)); Consumed(api.Fire(27, Repeat));
            Consumed(api.Fire(27, Up)); Passed(api.Fire(27, Up)); Pump(); Assert(actions == 1, "preview cannot release ownership or duplicate action");
        }
    }
    private static void DeclinedPress()
    {
        scenarios++; FakeNative api = new FakeNative(); bool accept = false; int actions = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => accept, k => actions++, null, api))
        {
            Passed(api.Fire(46, Down)); accept = true; Passed(api.Fire(46, Repeat)); Passed(api.Fire(46, Up));
            Consumed(api.Fire(46, Down)); Pump(); Assert(actions == 1, "never acquire previously declined press during repeat");
        }
    }
    private static void TwoKeys()
    {
        scenarios++; FakeNative api = new FakeNative(); List<int> keys = new List<int>();
        using (OfficeEditKeys guard = new OfficeEditKeys(k => true, k => keys.Add(k), null, api))
        {
            Consumed(api.Fire(27, Down)); Consumed(api.Fire(46, Down)); Consumed(api.Fire(27, Up)); Consumed(api.Fire(46, Up));
            Pump(); Assert(keys.Count == 2 && keys[0] == 27 && keys[1] == 46, "distinct presses preserve order");
        }
    }
    private static void SelfDispose()
    {
        scenarios++; FakeNative api = new FakeNative(); int actions = 0; OfficeEditKeys guard = null;
        guard = new OfficeEditKeys(k => true, k => { actions++; guard.Dispose(); Pump(); }, null, api);
        Consumed(api.Fire(27, Down)); Consumed(api.Fire(46, Down)); Pump();
        Assert(actions == 1 && api.UnhookCalls == 1 && !IsWindow(api.InstalledWindow), "handler disposal cancels remaining queue safely"); guard.Dispose();
    }
    private static void LateDelivery()
    {
        scenarios++; FakeNative api = new FakeNative(); int actions = 0;
        OfficeEditKeys guard = new OfficeEditKeys(k => true, k => actions++, null, api);
        OfficeEditKeys.HookProc late = (OfficeEditKeys.HookProc)api.Callback.Target;
        Consumed(api.Fire(27, Down)); guard.Dispose(); guard.Dispose(); Pump();
        Passed(late(0, new IntPtr(27), new IntPtr(1))); Assert(actions == 0 && api.UnhookCalls == 1, "late native and posted work ignored");
    }
    private static void WrongThread()
    {
        scenarios++; FakeNative api = new FakeNative();
        using (OfficeEditKeys guard = new OfficeEditKeys(k => true, k => { }, null, api))
        {
            Exception failure = null; IntPtr result = IntPtr.Zero;
            Thread thread = new Thread(() => { try { guard.Dispose(); } catch (Exception e) { failure = e; } result = api.Fire(27, Down); });
            thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
            Assert(failure is InvalidOperationException && api.UnhookCalls == 0, "cross-thread dispose rejected before native mutation"); Passed(result);
        }
        Exception mta = null; Thread other = new Thread(() => { try { new OfficeEditKeys(k => true, k => { }); } catch (Exception e) { mta = e; } });
        other.SetApartmentState(ApartmentState.MTA); other.Start(); other.Join(); Assert(mta is InvalidOperationException, "MTA creation rejected");
    }
    private static void PredicateFailure()
    {
        scenarios++; FakeNative api = new FakeNative(); Exception expected = new Exception("predicate"); int errors = 0, actions = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => { throw expected; }, k => actions++, e => { Assert(e == expected, "original predicate error"); errors++; throw new Exception("reporter"); }, api))
        {
            Passed(api.Fire(27, Down)); Assert(errors == 0, "onError deferred outside hook"); Pump(); Assert(errors == 1 && actions == 0, "predicate and error-handler exceptions contained");
        }
    }
    private static void HandlerFailure()
    {
        scenarios++; FakeNative api = new FakeNative(); Exception expected = new Exception("handler"); int errors = 0, owner = Thread.CurrentThread.ManagedThreadId;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => true, k => { throw expected; }, e => { Assert(e == expected && Thread.CurrentThread.ManagedThreadId == owner, "onKey error on STA"); errors++; }, api))
        { Consumed(api.Fire(46, Down)); Assert(errors == 0, "no callback in hook"); Pump(); Assert(errors == 1, "onKey exception contained"); }
    }
    private static void PostingFailure()
    {
        scenarios++; FakeNative api = new FakeNative { PostFailures = 2 }; int errors = 0, actions = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => true, k => actions++, e => { Assert(e is Win32Exception && ((Win32Exception)e).NativeErrorCode == 1234, "post original Win32 code"); errors++; }, api))
        {
            Consumed(api.Fire(27, Down)); Consumed(api.Fire(27, Repeat)); Assert(actions == 0 && errors == 0, "failed posting never invokes user code inside hook");
            Consumed(api.Fire(27, Up)); Pump(); Assert(actions == 1 && errors == 1 && api.PostCalls == 3, "retry preserves action and bounds error queue");
        }
    }
    private static void InstallFailure()
    {
        scenarios++; FakeNative api = new FakeNative { FailInstall = true, Error = 5 };
        Win32Exception error = Throws<Win32Exception>(() => new OfficeEditKeys(k => true, k => { }, null, api), "install must fail");
        Assert(error.NativeErrorCode == 5 && api.UnhookCalls == 0 && !IsWindow(api.InstalledWindow), "failed install preserves native code and destroys window");
    }
    private static void UnhookFailure()
    {
        scenarios++; FakeNative api = new FakeNative { UnhookFailures = 1, Error = 6 }; int actions = 0;
        OfficeEditKeys guard = new OfficeEditKeys(k => true, k => actions++, null, api);
        Consumed(api.Fire(27, Down)); Win32Exception error = Throws<Win32Exception>(() => guard.Dispose(), "unhook must fail");
        Assert(error.NativeErrorCode == 6 && IsWindow(api.InstalledWindow), "failed unhook retains owner/window for retry");
        Pump(); Passed(api.Fire(27, Down)); Assert(actions == 0, "failed disposal still stops application delivery");
        guard.Dispose(); Assert(api.UnhookCalls == 2 && api.Unhooked == api.Handle && !IsWindow(api.InstalledWindow), "retry removes exact hook");
    }
    private static WeakReference CreateUnreferenced(FakeNative api)
    { return new WeakReference(new OfficeEditKeys(k => true, k => { }, null, api)); }
    private static void DelegateRoot()
    {
        scenarios++; FakeNative api = new FakeNative(); WeakReference owner = CreateUnreferenced(api);
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
        Assert(owner.IsAlive && api.Callback.IsAlive, "native hook delegate remains strongly rooted");
        ((OfficeEditKeys)owner.Target).Dispose();
    }
    private static void WrongCookie()
    {
        scenarios++; FakeNative api = new FakeNative(); int actions = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => true, k => actions++, null, api))
        {
            Consumed(api.Fire(27, Down)); SendMessageW(api.LastWindow, api.LastMessage, new IntPtr(api.LastCookie.ToInt64() + 1), IntPtr.Zero);
            Assert(actions == 0, "stale window-message cookie ignored"); Pump(); Assert(actions == 1, "authentic queued dispatch retained");
        }
    }
    private static void ReentrantPump()
    {
        scenarios++; FakeNative api = new FakeNative(); int depth = 0, maximum = 0, actions = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => true, k => { depth++; maximum = Math.Max(maximum, depth); actions++; if (k == 27) { Consumed(api.Fire(46, Down)); Pump(); } depth--; }, null, api))
        { Consumed(api.Fire(27, Down)); Pump(); Assert(actions == 2 && maximum == 1, "nested pump never reenters onKey"); }
    }
    private static void InvalidArguments()
    {
        scenarios++;
        Throws<ArgumentNullException>(() => new OfficeEditKeys(null, k => { }), "null predicate");
        Throws<ArgumentNullException>(() => new OfficeEditKeys(k => true, null), "null action");
        FakeNative api = new FakeNative { ZeroThread = true };
        Throws<InvalidOperationException>(() => new OfficeEditKeys(k => true, k => { }, null, api), "zero thread rejected");
        Assert(api.InstallCalls == 0, "never fall back to global hook");
    }
    private static void NativeMenus()
    {
        scenarios++; FakeNative api = new FakeNative(); int predicates = 0, actions = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => { predicates++; return true; }, k => actions++, null, api))
        {
            foreach (uint flag in new uint[] { 4, 8, 16, 28 })
            { api.MenuFlags = flag; Passed(api.Fire(27, Down)); Passed(api.Fire(46, Down)); }
            Assert(predicates == 0 && api.PostCalls == 0, "native menus retain new Escape/Delete presses");
            api.MenuFlags = 3; Consumed(api.Fire(27, Down)); api.MenuFlags = 28;
            Consumed(api.Fire(27, Repeat)); Consumed(api.Fire(27, Up)); Passed(api.Fire(27, Down));
            Pump(); Assert(actions == 1 && predicates == 1, "menu opened during owned press cannot steal release");
        }
    }
    private static void GuiQueryFailure()
    {
        scenarios++; FakeNative api = new FakeNative { FailGui = true }; int predicates = 0, errors = 0;
        using (OfficeEditKeys guard = new OfficeEditKeys(k => { predicates++; return true; }, k => { }, e => { Assert(e is Win32Exception && ((Win32Exception)e).NativeErrorCode == 1234, "GUI query error preserved"); errors++; }, api))
        {
            Passed(api.Fire(27, Down)); Assert(predicates == 0 && errors == 0, "unknown native menu state never claims key");
            Pump(); Assert(errors == 1, "GUI query failure reported asynchronously");
        }
    }
    private static void RealThreadHook()
    {
        scenarios++; IntPtr handle;
        Assert((new OfficeEditKeys.NativeApi().GuiFlags(new OfficeEditKeys.NativeApi().CurrentThreadId()) & 0x1C) == 0, "real GUI thread structure marshals without menu");
        using (OfficeEditKeys guard = new OfficeEditKeys(k => false, k => { throw new Exception("no keys injected"); }))
        { handle = guard.MessageHandle; Assert(IsWindow(handle), "real thread-local hook installed"); }
        Assert(!IsWindow(handle), "real hook and window disposed");
    }
    [DllImport("user32.dll")] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Unicode)] private static extern IntPtr FindWindowExW(IntPtr parent, IntPtr after, string className, string title);
    [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr SendMessageW(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
}
