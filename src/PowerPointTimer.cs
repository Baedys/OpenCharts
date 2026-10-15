using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;

namespace OpenSlideCharts
{
    /// <summary>
    /// A 200 ms thread timer. Requires Office's pumped STA (the add-in uses
    /// StandardOleMarshalObject). Does not access or create a document window.
    /// </summary>
    public sealed class PowerPointTimer : IDisposable
    {
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate void TimerProc(IntPtr window, uint message, UIntPtr id, uint time);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern UIntPtr SetTimer(IntPtr window, UIntPtr id, uint milliseconds, TimerProc callback);
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool KillTimer(IntPtr window, UIntPtr id);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        private static readonly object Gate = new object();
        private static readonly Dictionary<Tuple<uint, UIntPtr>, PowerPointTimer> Active = new Dictionary<Tuple<uint, UIntPtr>, PowerPointTimer>();
        // Static lifetime is intentional: KillTimer leaves already queued messages.
        private static readonly TimerProc Callback = Dispatch;
        private readonly uint ownerThread;
        private readonly Action tick;
        private readonly Action<Exception> error;
        private UIntPtr id;
        private bool enabled, disposed, invoking;

        public PowerPointTimer(object app, Action tick, Action<Exception> error = null)
        {
            if (app == null) throw new ArgumentNullException("app");
            if (tick == null) throw new ArgumentNullException("tick");
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Create PowerPointTimer on Office's pumped STA thread.");
            ownerThread = GetCurrentThreadId(); this.tick = tick; this.error = error;
        }
        public void Start()
        {
            CheckThread();
            if (disposed) throw new ObjectDisposedException("PowerPointTimer");
            if (enabled) return;
            if (id != UIntPtr.Zero) Stop();
            UIntPtr created = SetTimer(IntPtr.Zero, UIntPtr.Zero, 200, Callback);
            if (created == UIntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not start Office's thread timer.");
            id = created;
            lock (Gate) Active.Add(Tuple.Create(ownerThread, id), this);
            enabled = true;
        }
        public void Stop()
        {
            CheckThread(); enabled = false;
            if (id == UIntPtr.Zero) return;
            if (!KillTimer(IntPtr.Zero, id)) throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not stop Office's thread timer.");
            lock (Gate) Active.Remove(Tuple.Create(ownerThread, id));
            id = UIntPtr.Zero;
        }
        public void Dispose()
        {
            CheckThread(); if (disposed) return;
            Stop(); disposed = true;
        }
        private void CheckThread()
        {
            if (GetCurrentThreadId() != ownerThread)
                throw new InvalidOperationException("Use PowerPointTimer on its original Office UI thread. Created=" + ownerThread + "; current=" + GetCurrentThreadId() + ".");
        }
        private static void Dispatch(IntPtr window, uint message, UIntPtr id, uint time)
        {
            // Windows can reuse system IDs after Stop. A queued old tick can then
            // reach the CURRENT registration early, never a retired delegate.
            // Consumers must observe current state, not treat ticks as elapsed time.
            try
            {
                PowerPointTimer timer;
                lock (Gate) if (!Active.TryGetValue(Tuple.Create(GetCurrentThreadId(), id), out timer)) return;
                if (window != IntPtr.Zero || message != 0x0113 || !timer.enabled || timer.disposed || timer.invoking) return;
                timer.invoking = true;
                try { timer.tick(); }
                catch (Exception failure)
                {
                    try { timer.Stop(); }
                    catch (Exception stopFailure) { failure = new AggregateException(failure, stopFailure); }
                    if (timer.error != null) try { timer.error(failure); } catch (Exception) { }
                }
                finally { timer.invoking = false; }
            }
            catch (Exception) { /* Never unwind a managed exception through user32. */ }
        }
    }
}
