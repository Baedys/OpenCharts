using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    /// <summary>
    /// Window-only connector handles. Coordinates use the host's Windows Forms screen units.
    /// Call on the owning UI thread. Callbacks run synchronously and may dispose this form.
    /// Callback exceptions propagate; the caller must provide its own error boundary.
    /// A drop ends capture before invoking dropped; disposal never invokes cancelled.
    /// </summary>
    [ComVisible(false)]
    public sealed class ConnectorHandleOverlay : Form
    {
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW", SetLastError = true)]
        private static extern IntPtr SetOwner64(IntPtr window, int index, IntPtr value);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongW", SetLastError = true)]
        private static extern int SetOwner32(IntPtr window, int index, int value);
        [DllImport("kernel32.dll", EntryPoint = "SetLastError")]
        private static extern void ClearLastError(uint error);
        private sealed class NativeOwner : IWin32Window
        {
            internal IntPtr Value;
            public IntPtr Handle { get { return Value; } }
        }

        private readonly Action<bool, Point> dropped;
        private readonly Action cancelled;
        private readonly Action started;
        private Point first, second, current;
        private Rectangle geometry;
        private IntPtr owner;
        private bool hasGeometry;
        private int dragging = -1;
        private bool disposing;
        private const int Radius = 6;

        public ConnectorHandleOverlay(Action<bool, Point> dropped, Action cancelled, Action started = null)
        {
            if (dropped == null) throw new ArgumentNullException("dropped");
            if (cancelled == null) throw new ArgumentNullException("cancelled");
            this.dropped = dropped;
            this.cancelled = cancelled;
            this.started = started;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            Text = "OpenCharts connector handles";
            AutoScaleMode = AutoScaleMode.None;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;
            DoubleBuffered = true;
        }

        public bool IsDragging { get { return dragging >= 0; } }
        public Point DragStartScreen { get; private set; }
        protected override bool ShowWithoutActivation { get { return true; } }
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams result = base.CreateParams;
                result.ExStyle |= 0x08000000 | 0x00000080; // NOACTIVATE | TOOLWINDOW
                return result;
            }
        }

        public void SetGeometry(Rectangle screenBounds, Point firstScreen, Point secondScreen, IntPtr ownerHandle)
        {
            if (IsDisposed || disposing) throw new ObjectDisposedException("ConnectorHandleOverlay");
            if (screenBounds.Width <= 0 || screenBounds.Height <= 0)
                throw new ArgumentOutOfRangeException("screenBounds");
            bool changed = !hasGeometry || geometry != screenBounds || first != firstScreen ||
                second != secondScreen || owner != ownerHandle;
            bool notify = changed && EndDrag();
            bool changeOwner = hasGeometry && owner != ownerHandle;
            geometry = screenBounds;
            first = firstScreen;
            second = secondScreen;
            owner = ownerHandle;
            hasGeometry = true;
            if (changeOwner && Visible) Hide();
            Bounds = screenBounds;
            if (!Visible)
            {
                Show(new NativeOwner { Value = ownerHandle });
                // WinForms retains its previous unmanaged owner when Show receives zero.
                if (ownerHandle == IntPtr.Zero)
                {
                    ClearLastError(0);
                    IntPtr previous = IntPtr.Size == 8 ? SetOwner64(Handle, -8, IntPtr.Zero) :
                        new IntPtr(SetOwner32(Handle, -8, 0));
                    int error = Marshal.GetLastWin32Error();
                    if (previous == IntPtr.Zero && error != 0) throw new Win32Exception(error);
                }
            }
            Invalidate();
            if (notify) cancelled(); // Terminal callback: may dispose this form.
        }

        public void CancelDrag()
        {
            if (EndDrag()) cancelled();
        }

        public void HideHandles()
        {
            if (IsDisposed || disposing) return;
            bool notify = EndDrag();
            Hide();
            if (notify) cancelled();
        }

        // Clear state first: releasing capture raises OnMouseCaptureChanged synchronously.
        private bool EndDrag()
        {
            if (dragging < 0) return false;
            dragging = -1;
            Capture = false;
            if (!IsDisposed && !disposing) Invalidate();
            return true;
        }

        private Point Client(Point screen)
        {
            return new Point(screen.X - geometry.X, screen.Y - geometry.Y);
        }

        private static bool Near(Point a, Point b)
        {
            long dx = (long)a.X - b.X, dy = (long)a.Y - b.Y;
            if (Math.Abs(dx) > Radius || Math.Abs(dy) > Radius) return false;
            return dx * dx + dy * dy <= Radius * Radius;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (IsDisposed || !hasGeometry) return;
            Point a = Client(first), b = Client(second);
            if (dragging == 0) a = current;
            if (dragging == 1) b = current;
            using (Pen pen = new Pen(Color.DodgerBlue, 2))
            {
                if (IsDragging) e.Graphics.DrawLine(pen, a, b);
                DrawHandle(e.Graphics, pen, a);
                DrawHandle(e.Graphics, pen, b);
            }
        }

        private static void DrawHandle(Graphics graphics, Pen pen, Point center)
        {
            Rectangle circle = new Rectangle(center.X - 5, center.Y - 5, 10, 10);
            graphics.FillEllipse(Brushes.White, circle);
            graphics.DrawEllipse(pen, circle);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (IsDisposed || disposing || !hasGeometry || !Visible || IsDragging || e.Button != MouseButtons.Left) return;
            dragging = Near(e.Location, Client(first)) ? 0 : Near(e.Location, Client(second)) ? 1 : -1;
            if (!IsDragging) return;
            DragStartScreen = PointToScreen(e.Location);
            current = e.Location;
            Capture = true;
            if (!IsDragging) return; // Capture can be refused or stolen synchronously.
            if (!Capture) { CancelDrag(); return; }
            Invalidate();
            if (started != null) started();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (IsDisposed || !IsDragging) return;
            current = e.Location;
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (IsDisposed || !IsDragging || e.Button != MouseButtons.Left) return;
            bool isFirst = dragging == 0;
            Point screen = PointToScreen(e.Location);
            EndDrag();
            dropped(isFirst, screen);
        }

        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            base.OnMouseCaptureChanged(e);
            if (!IsDisposed && !disposing && !Capture && EndDrag()) cancelled();
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (!IsDisposed && !disposing && !Visible && EndDrag()) cancelled();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x21) { m.Result = new IntPtr(3); return; } // MA_NOACTIVATE
            if (m.Msg == 0x84 && hasGeometry) // WM_NCHITTEST: signed screen coordinates.
            {
                long packed = m.LParam.ToInt64();
                Point point = new Point((short)(packed & 0xffff), (short)((packed >> 16) & 0xffff));
                m.Result = new IntPtr(IsDragging || Near(point, first) || Near(point, second) ? 1 : -1);
                return;
            }
            if (m.Msg == 0x1f) // WM_CANCELMODE
            {
                bool notify = EndDrag();
                base.WndProc(ref m);
                if (notify) cancelled();
                return;
            }
            base.WndProc(ref m);
        }

        protected override void Dispose(bool disposingManaged)
        {
            disposing = true;
            EndDrag();
            base.Dispose(disposingManaged);
        }
    }
}
