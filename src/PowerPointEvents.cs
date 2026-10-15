using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;

namespace OpenSlideCharts
{
    public delegate void BeforeRightClick(object selection, ref bool cancel);
    public delegate void BeforePresentationAction(object presentation, ref bool cancel);

    /// <summary>
    /// Subscribes to PowerPoint's EApplication connection point without a PIA dependency.
    /// Construct and dispose on the same Office thread. Callback exceptions are not hidden.
    /// </summary>
    public sealed class PowerPointEvents : IDisposable
    {
        public static readonly Guid EventInterfaceId = new Guid("914934C2-5A91-11CF-8700-00AA0060263B");
        public const int SelectionChangedDispid = 2001;
        public const int BeforeRightClickDispid = 2002;
        public const int BeforeSaveDispid = 2018;
        public const int BeforeCloseDispid = 2025;
        public const int WindowDeactivateDispid = 2010;

        private sealed class Subscription
        {
            internal int Dispid;
            internal Delegate Handler;
        }

        private object application;
        private readonly int ownerThread;
        private readonly Action<object, Guid, int, Delegate> remove;
        private readonly List<Subscription> subscriptions = new List<Subscription>();
        private bool callbacksEnabled;

        public PowerPointEvents(object app, Action<object> selectionChanged, BeforeRightClick beforeRightClick,
            BeforePresentationAction beforeSave = null, BeforePresentationAction beforeClose = null,
            Action<object, object> windowDeactivate = null)
            : this(ValidateApplication(app), selectionChanged, beforeRightClick, beforeSave, beforeClose,
                windowDeactivate, ComEventsHelper.Combine, RemoveComEvent)
        {
        }

        // Injectable registration operations allow lifecycle and ref-argument probes without Office.
        internal PowerPointEvents(object app, Action<object> selectionChanged, BeforeRightClick beforeRightClick,
            BeforePresentationAction beforeSave, BeforePresentationAction beforeClose,
            Action<object, object> windowDeactivate,
            Action<object, Guid, int, Delegate> combine, Action<object, Guid, int, Delegate> remove)
        {
            if (app == null) throw new ArgumentNullException("app");
            if (combine == null) throw new ArgumentNullException("combine");
            if (remove == null) throw new ArgumentNullException("remove");
            if (selectionChanged == null && beforeRightClick == null && beforeSave == null &&
                beforeClose == null && windowDeactivate == null)
                throw new ArgumentException("At least one PowerPoint event callback is required.");

            application = app;
            ownerThread = Thread.CurrentThread.ManagedThreadId;
            this.remove = remove;
            try
            {
                if (selectionChanged != null)
                    Add(combine, SelectionChangedDispid, new Action<object>(delegate(object selection)
                    {
                        if (callbacksEnabled) selectionChanged(selection);
                    }));
                if (beforeRightClick != null)
                    Add(combine, BeforeRightClickDispid, new BeforeRightClick(delegate(object selection, ref bool cancel)
                    {
                        if (callbacksEnabled) beforeRightClick(selection, ref cancel);
                    }));
                if (beforeSave != null)
                    Add(combine, BeforeSaveDispid, new BeforePresentationAction(delegate(object presentation, ref bool cancel)
                    {
                        if (callbacksEnabled) beforeSave(presentation, ref cancel);
                    }));
                if (beforeClose != null)
                    Add(combine, BeforeCloseDispid, new BeforePresentationAction(delegate(object presentation, ref bool cancel)
                    {
                        if (callbacksEnabled) beforeClose(presentation, ref cancel);
                    }));
                if (windowDeactivate != null)
                    Add(combine, WindowDeactivateDispid, new Action<object, object>(delegate(object presentation, object window)
                    {
                        if (callbacksEnabled) windowDeactivate(presentation, window);
                    }));
                callbacksEnabled = true;
            }
            catch (Exception attachError)
            {
                try { Dispose(); }
                catch (Exception cleanupError)
                {
                    throw new AggregateException("PowerPoint event attachment and rollback both failed.", attachError, cleanupError);
                }
                throw;
            }
        }

        private static object ValidateApplication(object app)
        {
            if (app == null) throw new ArgumentNullException("app");
            if (!Marshal.IsComObject(app)) throw new ArgumentException("A PowerPoint application COM object is required.", "app");
            return app;
        }

        private void Add(Action<object, Guid, int, Delegate> combine, int dispid, Delegate handler)
        {
            combine(application, EventInterfaceId, dispid, handler);
            subscriptions.Add(new Subscription { Dispid = dispid, Handler = handler });
        }

        private static void RemoveComEvent(object app, Guid iid, int dispid, Delegate handler)
        {
            ComEventsHelper.Remove(app, iid, dispid, handler);
        }

        public void Dispose()
        {
            if (application == null) return;
            if (Thread.CurrentThread.ManagedThreadId != ownerThread)
                throw new InvalidOperationException("Dispose PowerPointEvents on its original Office thread.");

            // Suppress even reentrant callbacks during Unadvise. Keep failed subscriptions
            // available for a later retry, while still attempting every other removal.
            callbacksEnabled = false;
            List<Exception> errors = new List<Exception>();
            for (int i = subscriptions.Count - 1; i >= 0; --i)
            {
                Subscription subscription = subscriptions[i];
                try
                {
                    remove(application, EventInterfaceId, subscription.Dispid, subscription.Handler);
                    subscriptions.RemoveAt(i);
                }
                catch (Exception error) { errors.Add(error); }
            }
            if (subscriptions.Count == 0) application = null;
            if (errors.Count != 0)
                throw new AggregateException("One or more PowerPoint event subscriptions could not be removed.", errors);
            // The caller owns the application RCW; releasing it here could break other users.
        }
    }
}
