using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    /// <summary>
    /// Captures only owned, unmodified Escape/Delete presses on its creation
    /// thread. shouldConsume may inspect managed state and cheap read-only native
    /// focus state. It runs inside WH_KEYBOARD: no COM, message pumping or UI changes.
    /// onKey/onError run later on a private message-only window on the STA.
    /// Dispose on the creation thread; if unhook fails, retry Dispose there.
    /// </summary>
    [ComVisible(false)]
    public sealed class OfficeEditKeys : IDisposable
    {
        public const int Escape = 27, Delete = 46;
        private const int KeyboardHook = 2, ActionCode = 0, NoRemoveCode = 3;
        private const int DispatchMessage = 0x8000 + 1211;
        private static readonly object RootsLock = new object();
        private static readonly HashSet<OfficeEditKeys> Roots = new HashSet<OfficeEditKeys>();
        private static int nextCookie;
        private readonly Func<int, bool> shouldConsume;
        private readonly Action<int> onKey;
        private readonly Action<Exception> onError;
        private readonly NativeApi native;
        private readonly uint nativeThread;
        private readonly int managedThread, cookie;
        private readonly HookProc hookCallback;
        private readonly Queue<Notice> pending = new Queue<Notice>();
        private MessageWindow window;
        private IntPtr hook;
        private int ownedKeys;
        private bool stopping, completed, posted, dispatching, postErrorQueued;

        public OfficeEditKeys(Func<int, bool> shouldConsume, Action<int> onKey, Action<Exception> onError = null)
            : this(shouldConsume, onKey, onError, new NativeApi()) { }

        internal OfficeEditKeys(Func<int, bool> shouldConsume, Action<int> onKey, Action<Exception> onError, NativeApi native)
        {
            if (shouldConsume == null) throw new ArgumentNullException("shouldConsume");
            if (onKey == null) throw new ArgumentNullException("onKey");
            if (native == null) throw new ArgumentNullException("native");
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("OfficeEditKeys requires an STA creation thread with a message loop.");
            this.shouldConsume = shouldConsume; this.onKey = onKey; this.onError = onError; this.native = native;
            managedThread = Thread.CurrentThread.ManagedThreadId;
            nativeThread = native.CurrentThreadId();
            if (nativeThread == 0) throw new InvalidOperationException("A nonzero current thread ID is required; global keyboard hooks are not supported.");
            int serial = Interlocked.Increment(ref nextCookie);
            if (serial == 0) serial = Interlocked.Increment(ref nextCookie);
            cookie = serial;
            hookCallback = KeyboardCallback;
            window = new MessageWindow(this);
            try
            {
                // WH_KEYBOARD, current process/current thread only, module NULL.
                hook = native.InstallHook(KeyboardHook, hookCallback, IntPtr.Zero, nativeThread);
                if (hook == IntPtr.Zero) throw new Win32Exception(native.LastError());
                // A native callback pointer alone does not keep its delegate
                // alive. Retain the owner until its exact hook is removed.
                lock (RootsLock) Roots.Add(this);
            }
            catch (Exception failure)
            {
                stopping = true;
                if (hook != IntPtr.Zero)
                {
                    if (native.Unhook(hook)) hook = IntPtr.Zero;
                    else { lock (RootsLock) Roots.Add(this); }
                }
                if (hook == IntPtr.Zero)
                {
                    try { window.DestroyHandle(); }
                    catch (Exception cleanup) { failure.Data["OfficeEditKeysCleanup"] = cleanup; }
                }
                throw;
            }
        }

        public void Dispose()
        {
            RequireThread();
            if (completed) return;
            // Disable first: queued delivery and late hook callbacks cannot
            // invoke user code while native teardown is underway or retried.
            stopping = true; ownedKeys = 0; pending.Clear(); posted = false;
            if (hook != IntPtr.Zero)
            {
                if (!native.Unhook(hook)) throw new Win32Exception(native.LastError());
                hook = IntPtr.Zero;
            }
            if (window != null && window.Handle != IntPtr.Zero) window.DestroyHandle();
            completed = true;
            lock (RootsLock) Roots.Remove(this);
            GC.KeepAlive(hookCallback);
        }

        internal IntPtr MessageHandle { get { return window == null ? IntPtr.Zero : window.Handle; } }

        private bool IsCreationThread()
        {
            return Thread.CurrentThread.ManagedThreadId == managedThread && native.CurrentThreadId() == nativeThread;
        }

        private void RequireThread()
        {
            if (!IsCreationThread()) throw new InvalidOperationException("OfficeEditKeys must be used and disposed on its creation thread.");
        }

        private IntPtr KeyboardCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (code < 0 || stopping || !IsCreationThread()) return Next(code, wParam, lParam);
                long keyValue = wParam.ToInt64();
                int mask = keyValue == Escape ? 1 : keyValue == Delete ? 2 : 0;
                bool consumed = false;
                if (mask != 0)
                {
                    // PeekMessage(PM_NOREMOVE) is only a preview. It must not
                    // clear press ownership or dispatch a new action twice.
                    if (code == NoRemoveCode) consumed = (ownedKeys & mask) != 0;
                    else if (code == ActionCode)
                    {
                        long bits = lParam.ToInt64();
                        bool keyUp = (bits & 0x80000000L) != 0;
                        bool wasDown = (bits & 0x40000000L) != 0;
                        if ((ownedKeys & mask) != 0)
                        {
                            if (keyUp) ownedKeys &= ~mask;
                            consumed = true; // Owned repeats/up remain ours even if modifiers change.
                        }
                        // Native Office popup/system menus retain new keys.
                        // A failed GUI query follows the deferred-error/pass path.
                        else if (!keyUp && !wasDown && !Modified(bits) && (native.GuiFlags(nativeThread) & 0x1C) == 0)
                        {
                            int key = (int)keyValue;
                            if (shouldConsume(key))
                            {
                                ownedKeys |= mask;
                                pending.Enqueue(new Notice { Key = key });
                                consumed = true;
                            }
                        }
                    }
                }
                EnsurePosted();
                return consumed ? new IntPtr(1) : Next(code, wParam, lParam);
            }
            catch (Exception error)
            {
                // No application callback, including onError, runs in a hook.
                QueueError(error); EnsurePosted();
                return Next(code, wParam, lParam);
            }
        }

        private bool Modified(long bits)
        {
            return (bits & 0x20000000L) != 0 || (native.KeyState(0x10) & 0x8000) != 0 ||
                (native.KeyState(0x11) & 0x8000) != 0 || (native.KeyState(0x12) & 0x8000) != 0;
        }

        private IntPtr Next(int code, IntPtr wParam, IntPtr lParam)
        {
            try { return native.CallNext(hook, code, wParam, lParam); }
            catch (Exception error) { QueueError(error); EnsurePosted(); return IntPtr.Zero; }
        }

        private void QueueError(Exception error)
        {
            if (!stopping && IsCreationThread()) pending.Enqueue(new Notice { Error = error });
        }

        private void EnsurePosted()
        {
            if (stopping || posted || dispatching || pending.Count == 0 || window == null || window.Handle == IntPtr.Zero) return;
            posted = true;
            try
            {
                if (!native.Post(window.Handle, DispatchMessage, new IntPtr(cookie), IntPtr.Zero))
                    throw new Win32Exception(native.LastError());
                postErrorQueued = false;
            }
            catch (Exception error)
            {
                posted = false;
                // Keep the consumed key pending and retry on a later hook
                // notification. Avoid an error flood when the queue is full.
                if (!postErrorQueued) { pending.Enqueue(new Notice { Error = error }); postErrorQueued = true; }
            }
        }

        private void DispatchPending()
        {
            if (stopping || !IsCreationThread() || dispatching) return;
            posted = false; dispatching = true;
            try
            {
                while (!stopping && pending.Count > 0)
                {
                    Notice notice = pending.Dequeue(); // Clear before onKey can dispose or pump messages.
                    if (notice.Error != null) ReportError(notice.Error);
                    else
                    {
                        try { onKey(notice.Key); }
                        catch (Exception error) { ReportError(error); }
                    }
                }
            }
            finally { dispatching = false; }
        }

        private void ReportError(Exception error)
        {
            if (onError == null) return;
            try { onError(error); } catch { /* Never cross a native callback boundary. */ }
        }

        private sealed class Notice { internal int Key; internal Exception Error; }

        private sealed class MessageWindow : NativeWindow
        {
            private readonly OfficeEditKeys owner;
            internal MessageWindow(OfficeEditKeys owner)
            {
                this.owner = owner;
                CreateHandle(new CreateParams { Parent = new IntPtr(-3), Caption = "OpenCharts edit key dispatch" });
            }
            protected override void WndProc(ref Message message)
            {
                if (message.Msg == DispatchMessage)
                {
                    message.Result = IntPtr.Zero;
                    if (message.WParam.ToInt64() != owner.cookie) return;
                    try { owner.DispatchPending(); }
                    catch (Exception error) { owner.ReportError(error); }
                    return;
                }
                base.WndProc(ref message);
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        internal delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        // Injectable native boundary: tests exercise routing and real STA
        // message delivery without installing a keyboard hook in Office.
        internal class NativeApi
        {
            internal virtual uint CurrentThreadId() { return GetCurrentThreadId(); }
            internal virtual IntPtr InstallHook(int kind, HookProc callback, IntPtr module, uint threadId)
            { return SetWindowsHookExW(kind, callback, module, threadId); }
            internal virtual bool Unhook(IntPtr handle) { return UnhookWindowsHookEx(handle); }
            internal virtual IntPtr CallNext(IntPtr handle, int code, IntPtr key, IntPtr flags)
            { return CallNextHookEx(handle, code, key, flags); }
            internal virtual short KeyState(int key) { return GetKeyState(key); }
            internal virtual uint GuiFlags(uint threadId)
            {
                GuiThreadInfo info = new GuiThreadInfo();
                info.Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo));
                if (!GetGUIThreadInfo(threadId, ref info)) throw new Win32Exception(Marshal.GetLastWin32Error());
                // GUI_INMENUMODE=4, GUI_SYSTEMMENUMODE=8, GUI_POPUPMENUMODE=16.
                // https://learn.microsoft.com/windows/win32/api/winuser/ns-winuser-guithreadinfo
                return info.Flags;
            }
            internal virtual bool Post(IntPtr handle, int message, IntPtr wParam, IntPtr lParam)
            { return PostMessageW(handle, message, wParam, lParam); }
            internal virtual int LastError() { return Marshal.GetLastWin32Error(); }
            [StructLayout(LayoutKind.Sequential)]
            private struct GuiThreadInfo
            {
                internal uint Size, Flags;
                internal IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
                internal int Left, Top, Right, Bottom;
            }
            [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool GetGUIThreadInfo(uint threadId, ref GuiThreadInfo info);
            [DllImport("kernel32.dll", ExactSpelling = true)] private static extern uint GetCurrentThreadId();
            [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] private static extern IntPtr SetWindowsHookExW(int kind, HookProc callback, IntPtr module, uint threadId);
            [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool UnhookWindowsHookEx(IntPtr hook);
            [DllImport("user32.dll", ExactSpelling = true)] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr key, IntPtr flags);
            [DllImport("user32.dll", ExactSpelling = true)] private static extern short GetKeyState(int key);
            [DllImport("user32.dll", ExactSpelling = true, SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] private static extern bool PostMessageW(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
        }
    }
}
