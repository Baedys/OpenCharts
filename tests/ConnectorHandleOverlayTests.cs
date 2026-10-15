using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OpenSlideCharts;

// Local WinForms lifecycle tests, not HID/Office integration tests.
internal static class ConnectorHandleOverlayTests
{
    private static int checks;
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
    private static readonly Rectangle Box = new Rectangle(-20000, -20000, 180, 100);
    private static readonly Point First = new Point(-19980, -19970);
    private static readonly Point Second = new Point(-19860, -19970);

    private static void Check(bool value, string label)
    {
        if (!value) throw new Exception(label);
        checks++;
    }

    private static void Call(ConnectorHandleOverlay form, string name, params object[] args)
    {
        try { typeof(ConnectorHandleOverlay).GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, args); }
        catch (TargetInvocationException error) { throw error.InnerException; }
    }

    private static void Down(ConnectorHandleOverlay f, int x, int y)
    {
        Call(f, "OnMouseDown", new MouseEventArgs(MouseButtons.Left, 1, x, y, 0));
    }

    private static Message Send(ConnectorHandleOverlay f, int message, IntPtr data)
    {
        object[] args = { Message.Create(f.Handle, message, IntPtr.Zero, data) };
        Call(f, "WndProc", args);
        return (Message)args[0];
    }

    private static IntPtr Packed(Point p)
    {
        return new IntPtr(unchecked((int)((uint)(ushort)p.X | ((uint)(ushort)p.Y << 16))));
    }

    [STAThread]
    public static int Main()
    {
        try
        {
            int drops = 0, cancels = 0, starts = 0;
            bool firstDrop = false, settled = false;
            Point droppedPoint = Point.Empty;
            ConnectorHandleOverlay f = null;
            f = new ConnectorHandleOverlay(delegate(bool first, Point p)
            {
                drops++; firstDrop = first; droppedPoint = p;
                settled = !f.IsDragging && !f.Capture;
            }, delegate { cancels++; }, delegate { starts++; });
            using (f)
            using (Form owner = new Form())
            {
                IntPtr ownerHandle = owner.Handle;
                f.SetGeometry(Box, First, Second, ownerHandle);
                Check(GetWindow(f.Handle, 4) == ownerHandle, "Native owner assigned");
                Check(f.Visible && !f.IsDragging, "Geometry shows idle handles");
                Check(!f.ShowInTaskbar && f.FormBorderStyle == FormBorderStyle.None, "Tool window appearance");
                Check(f.TransparencyKey == Color.Magenta, "Transparent background");
                CreateParams cp = (CreateParams)typeof(ConnectorHandleOverlay).GetProperty("CreateParams", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(f, null);
                Check((cp.ExStyle & 0x08000080) == 0x08000080, "Nonactivating tool styles");
                Check(Send(f, 0x21, IntPtr.Zero).Result.ToInt32() == 3, "Mouse activation suppressed");
                Check(Send(f, 0x84, Packed(First)).Result.ToInt32() == 1, "First handle accepts hit");
                Check(Send(f, 0x84, Packed(Second)).Result.ToInt32() == 1, "Second handle accepts hit");
                Check(Send(f, 0x84, Packed(new Point(-19920, -19970))).Result.ToInt32() == -1, "Idle line passes hits through");
                Down(f, 80, 30);
                Check(!f.IsDragging && starts == 0, "Empty space cannot begin drag");
                Call(f, "OnMouseDown", new MouseEventArgs(MouseButtons.Right, 1, 20, 30, 0));
                Check(!f.IsDragging, "Right mouse cannot begin drag");
                Down(f, 20, 30);
                Check(f.IsDragging && f.Capture && starts == 1, "First handle takes native mouse capture");
                Check(f.DragStartScreen == First, "Drag origin comes from MouseDown coordinates, not the later cursor position");
                f.SetGeometry(Box, First, Second, ownerHandle);
                Check(f.IsDragging && cancels == 0, "Identical geometry preserves drag");
                Call(f, "OnMouseMove", new MouseEventArgs(MouseButtons.Left, 0, 80, 60, 0));
                Call(f, "OnMouseUp", new MouseEventArgs(MouseButtons.Right, 1, 80, 60, 0));
                Check(f.IsDragging, "Other button up preserves drag");
                Call(f, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, 80, 60, 0));
                Check(drops == 1 && firstDrop && settled, "Drop sees released capture and idle state");
                Check(droppedPoint == new Point(Box.X + 80, Box.Y + 60), "Drop uses screen coordinates");
                Check(cancels == 0, "Successful drop does not cancel");
                Down(f, 140, 30);
                Call(f, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, 145, 35, 0));
                Check(drops == 2 && !firstDrop, "Second handle identity");
                Down(f, 20, 30); f.CancelDrag(); f.CancelDrag();
                Check(cancels == 1 && !f.Capture, "Explicit cancellation once");
                Down(f, 20, 30); f.Capture = false;
                Check(cancels == 2 && !f.IsDragging, "Native capture loss cancels");
                Down(f, 20, 30); Send(f, 0x1f, IntPtr.Zero);
                Check(cancels == 3 && !f.Capture, "Cancel mode cancels");
                Down(f, 20, 30);
                f.SetGeometry(Box, new Point(First.X + 1, First.Y), Second, ownerHandle);
                Check(cancels == 4 && !f.IsDragging, "Endpoint change cancels");
                f.SetGeometry(Box, First, Second, ownerHandle); Down(f, 20, 30);
                using (Form other = new Form())
                {
                    f.SetGeometry(Box, First, Second, other.Handle);
                    Check(cancels == 5 && !f.IsDragging && f.Visible, "Owner change cancels and re-shows");
                    Check(GetWindow(f.Handle, 4) == other.Handle, "Native owner changes");
                    f.SetGeometry(Box, First, Second, IntPtr.Zero);
                    Check(GetWindow(f.Handle, 4) == IntPtr.Zero, "Native owner can be cleared");
                    f.SetGeometry(Box, First, Second, ownerHandle);
                }
                Down(f, 20, 30); f.HideHandles();
                Check(cancels == 6 && !f.Visible && !f.Capture, "HideHandles cancels and hides");
                f.SetGeometry(Box, First, Second, ownerHandle); Down(f, 20, 30); f.Hide();
                Check(cancels == 7 && !f.IsDragging, "Inherited Hide also cancels");
                f.SetGeometry(Box, First, Second, ownerHandle); Down(f, 20, 30); f.Dispose();
                Check(cancels == 7 && f.IsDisposed && !f.IsDragging, "Dispose cleans without callback");
            }
            ConnectorHandleOverlay disposeOnDrop = null;
            disposeOnDrop = new ConnectorHandleOverlay(delegate { disposeOnDrop.Dispose(); }, delegate { throw new Exception("Unexpected cancel"); });
            disposeOnDrop.SetGeometry(Box, First, Second, IntPtr.Zero);
            Down(disposeOnDrop, 20, 30);
            Call(disposeOnDrop, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, 30, 35, 0));
            Check(disposeOnDrop.IsDisposed, "Drop callback may dispose");
            ConnectorHandleOverlay disposeOnCancel = null;
            disposeOnCancel = new ConnectorHandleOverlay(delegate { }, delegate { disposeOnCancel.Dispose(); });
            disposeOnCancel.SetGeometry(Box, First, Second, IntPtr.Zero); Down(disposeOnCancel, 20, 30);
            disposeOnCancel.SetGeometry(Box, new Point(First.X + 1, First.Y), Second, IntPtr.Zero);
            Check(disposeOnCancel.IsDisposed, "Geometry cancellation callback may dispose");
            ConnectorHandleOverlay disposeOnStart = null;
            disposeOnStart = new ConnectorHandleOverlay(delegate { }, delegate { throw new Exception("Unexpected dispose cancellation"); }, delegate { disposeOnStart.Dispose(); });
            disposeOnStart.SetGeometry(Box, First, Second, IntPtr.Zero); Down(disposeOnStart, 20, 30);
            Check(disposeOnStart.IsDisposed && !disposeOnStart.IsDragging, "Started callback may dispose");
            int movedCancels = 0;
            using (ConnectorHandleOverlay moving = new ConnectorHandleOverlay(delegate { }, delegate { movedCancels++; }))
            {
                moving.SetGeometry(Box, First, Second, IntPtr.Zero); Down(moving, 20, 30);
                Rectangle shifted = Box; shifted.Offset(1, 0);
                moving.SetGeometry(shifted, First, Second, IntPtr.Zero);
                Check(movedCancels == 1 && !moving.IsDragging && !moving.Capture, "Bounds change cancels");
            }
            using (ConnectorHandleOverlay throwing = new ConnectorHandleOverlay(delegate { throw new InvalidOperationException("callback"); }, delegate { }))
            {
                throwing.SetGeometry(Box, First, Second, IntPtr.Zero); Down(throwing, 20, 30);
                bool propagated = false;
                try { Call(throwing, "OnMouseUp", new MouseEventArgs(MouseButtons.Left, 1, 30, 35, 0)); }
                catch (InvalidOperationException e) { propagated = e.Message == "callback"; }
                Check(propagated && !throwing.IsDragging && !throwing.Capture, "Drop exception propagates after cleanup");
            }
            Console.WriteLine("ConnectorHandleOverlay: " + checks + " local WinForms checks passed; no HID or Office coverage claimed.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
