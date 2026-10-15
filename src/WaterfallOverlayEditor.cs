using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    public sealed partial class WaterfallSlideEditor
    {
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetFocus();
        [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool ScreenToClient(IntPtr window, ref Point point);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string name, string title);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref Point point);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out NativeRectangle rectangle);
        [StructLayout(LayoutKind.Sequential)] private struct NativeRectangle { internal int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int count);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
        private ConnectorHandleOverlay interfaceHandles;
        private OfficeEditKeys editKeys;
        private object editWindow;
        private bool interfaceInputEnabled;
        private ConnectorScreenProjection projection, dragProjection;
        private string dragToken, dragConnector;
        private Point firstHandle, secondHandle, dragPress;

        public bool InterfaceHandlesVisible { get { return interfaceHandles != null && interfaceHandles.Visible; } }
        private bool InterfaceDragging { get { return interfaceHandles != null && interfaceHandles.IsDragging; } }

        private static IntPtr ForegroundOfficeFrame()
        {
            IntPtr foreground = GetForegroundWindow(); uint process;
            GetWindowThreadProcessId(foreground, out process);
            // External COM test controllers must not install hooks on their own UI thread.
            if (process != GetCurrentProcessId()) return IntPtr.Zero;
            var name = new StringBuilder(256); GetClassName(foreground, name, name.Capacity);
            if (name.ToString() != "PPTFrameClass") return IntPtr.Zero;
            return GetAncestor(foreground, 3);
        }
        private static bool SlideHasKeyboardFocus()
        {
            var name = new StringBuilder(256); GetClassName(GetFocus(), name, name.Capacity);
            // Fail closed on other Office generations and on ribbon/task-pane fields.
            return name.ToString() == "mdiClass" && ForegroundOfficeFrame() != IntPtr.Zero;
        }
        private bool ChartOwnsScreenPoint(Point point)
        {
            try
            {
                object hit = (object)((dynamic)editWindow).RangeFromPoint(point.X, point.Y);
                if (hit == null) return false;
                if (SameObject(hit, chart)) return true;
                dynamic shape = hit;
                return (int)shape.Child != 0 && SameObject((object)shape.ParentGroup, chart);
            }
            catch (COMException)
            {
                // Office may fail this advisory visibility query during layout updates.
                // It is advisory: hide now and retry on the next timer tick.
                return false;
            }
        }
        private bool ChartFitsViewport(IntPtr owner, Rectangle bounds)
        {
            // Supported Windows profile: the document client excludes ribbon and docked
            // task panes. Normal view additionally reserves outline and notes dividers.
            IntPtr client = FindWindowEx(owner, IntPtr.Zero, "MDIClient", null);
            NativeRectangle rectangle; Point origin = Point.Empty;
            if (client == IntPtr.Zero || !GetClientRect(client, out rectangle) || !ClientToScreen(client, ref origin)) return false;
            Rectangle viewport = new Rectangle(origin.X, origin.Y, rectangle.Right - rectangle.Left, rectangle.Bottom - rectangle.Top);
            dynamic window = editWindow; int view = window.ViewType;
            if (view == 9)
            {
                int horizontal = window.SplitHorizontal, vertical = window.SplitVertical;
                if (horizontal < 0 || horizontal >= 100 || vertical <= 0 || vertical > 100) return false;
                int reserved = (int)Math.Ceiling(viewport.Width * horizontal / 100.0);
                viewport = Rectangle.FromLTRB(viewport.Left + reserved, viewport.Top, viewport.Right,
                    viewport.Top + (int)Math.Floor(viewport.Height * vertical / 100.0));
            }
            else if (view != 1) return false;
            viewport.Inflate(-16, -16);
            return viewport.Width > 0 && viewport.Height > 0 && viewport.Contains(bounds);
        }
        private void HideInterfaceHandles()
        {
            interfaceInputEnabled = false;
            if (interfaceHandles != null) interfaceHandles.HideHandles();
        }
        private void UpdateInterfaceHandles()
        {
            interfaceInputEnabled = false;
            if (!Active || busy || choosing || AddingConnector || currentId == null || leaveRequested)
            { HideInterfaceHandles(); return; }
            IntPtr owner = ForegroundOfficeFrame();
            if (owner == IntPtr.Zero) { HideInterfaceHandles(); return; }
            if (!SameObject((object)app.ActiveWindow, editWindow) || !IsActiveSlide() || !ChartExists() || !FrameUnchanged())
            { Stop(false); return; }
            OfficeChartSelection selection = OfficeChartSelection.Resolve((object)app.ActiveWindow.Selection);
            if (selection == null || !SameObject(selection.Chart, chart) || selection.Primitive == null ||
                selection.Primitive.Role != "WATERFALL_CONNECTOR" || selection.Primitive.EntityId != currentId)
            { HideInterfaceHandles(); return; }
            EditLine line = lines[currentId];
            object live = FindLine(line.ShapeId);
            if (live == null || ConnectorEndpoints.Read(live).Difference(line.Original) > .2)
            { HideInterfaceHandles(); return; }
            dynamic window = editWindow;
            var current = new ConnectorScreenProjection(
                new Point((int)window.PointsToScreenPixelsX((float)left), (int)window.PointsToScreenPixelsY((float)top)),
                new Point((int)window.PointsToScreenPixelsX((float)(left + width)), (int)window.PointsToScreenPixelsY((float)(top + height))), width, height);
            // RangeFromPoint asks Office which document shape is actually visible.
            // Hide the interface when the chart is clipped by a pane, ribbon or zoom.
            // This avoids painting handles over unrelated Office controls.
            if (!ChartFitsViewport(owner, current.Bounds) || !ChartOwnsScreenPoint(current.ToScreen(1, 1)) || !ChartOwnsScreenPoint(current.ToScreen(width - 1, 1)) ||
                !ChartOwnsScreenPoint(current.ToScreen(1, height - 1)) || !ChartOwnsScreenPoint(current.ToScreen(width - 1, height - 1)))
            { HideInterfaceHandles(); return; }
            if (InterfaceDragging && (dragProjection == null || !dragProjection.SameGeometry(current))) interfaceHandles.CancelDrag();
            projection = current;
            if (interfaceHandles == null)
            {
                interfaceHandles = new ConnectorHandleOverlay(
                    delegate(bool first, Point point) { Run(delegate { DropInterfaceHandle(first, point); }); },
                    CancelInterfaceDrag,
                    delegate { dragProjection = projection; dragToken = token; dragConnector = currentId; dragPress = interfaceHandles.DragStartScreen; });
                interfaceHandles.MouseUp += delegate(object sender, MouseEventArgs args)
                {
                    if (args.Button != MouseButtons.Right || !SlideHasKeyboardFocus()) return;
                    Point point = interfaceHandles.PointToScreen(args.Location); IntPtr focus = GetFocus();
                    // PowerPoint opens its contextual menu from the button messages,
                    // rather than WM_CONTEXTMENU alone. Coordinates belong to its pane.
                    if (!ScreenToClient(focus, ref point)) return;
                    IntPtr position = new IntPtr(unchecked((point.Y << 16) | (point.X & 0xffff)));
                    PostMessage(focus, 0x204, new IntPtr(2), position);
                    PostMessage(focus, 0x205, IntPtr.Zero, position);
                };
            }
            if (editKeys == null)
                editKeys = new OfficeEditKeys(delegate(int key) { return interfaceInputEnabled && Active && !busy && !choosing && !AddingConnector && SlideHasKeyboardFocus(); },
                    delegate(int key) { Run(delegate { InterfaceKey(key); }); },
                    delegate(Exception error) { LastMessage = Describe(error); });
            firstHandle = current.ToScreen(line.Original.X1 - left, line.Original.Y1 - top);
            secondHandle = current.ToScreen(line.Original.X2 - left, line.Original.Y2 - top);
            interfaceHandles.SetGeometry(current.Bounds, firstHandle, secondHandle, owner);
            interfaceInputEnabled = true;
        }
        private void CancelInterfaceDrag()
        {
            dragToken = null; dragConnector = null; dragProjection = null;
        }
        private void InterfaceKey(int key)
        {
            if (!Active || busy || choosing || AddingConnector) return;
            if (key == OfficeEditKeys.Escape && InterfaceDragging) { interfaceHandles.CancelDrag(); return; }
            // Dispatch is deferred out of the keyboard hook; validate selection again.
            RequireCommandChart();
            OfficeChartSelection selection = OfficeChartSelection.Resolve((object)app.ActiveWindow.Selection);
            if (selection == null || !SameObject(selection.Chart, chart) || selection.Primitive == null ||
                selection.Primitive.Role != "WATERFALL_CONNECTOR" || selection.Primitive.EntityId != currentId) return;
            if (InterfaceDragging) return;
            if (key == OfficeEditKeys.Escape) Stop(false, true);
            else if (key == OfficeEditKeys.Delete) DeleteConnector(slide, chart, currentId);
            UpdateInterfaceHandles();
        }
        private void DropInterfaceHandle(bool first, Point point)
        {
            ConnectorScreenProjection dropped = dragProjection; string session = dragToken, connector = dragConnector;
            Point start = dragPress;
            CancelInterfaceDrag();
            if (!Active || token != session || currentId != connector || dropped == null) return;
            if (Math.Abs((long)point.X - start.X) <= 3 && Math.Abs((long)point.Y - start.Y) <= 3) return;
            RequireCommandChart();
            UpdateInterfaceHandles();
            if (!InterfaceHandlesVisible || !dropped.SameGeometry(projection)) return;
            PointF local = dropped.ToChart(point);
            WaterfallHit hit = scene.FindNearest(local.X, local.Y, 14);
            if (hit.Candidates.Count == 0) { LastMessage = "Déposez l'extrémité près d'un niveau de colonne."; return; }
            if (hit.IsAmbiguous)
            {
                var choices = new CommandMenu();
                foreach (WaterfallTarget item in hit.Candidates)
                {
                    WaterfallTarget target = item;
                    choices.Items.Add(TargetLabel(target), null, delegate
                    {
                        if (!Active || token != session || currentId != connector) return;
                        Run(delegate { RequireCommandChart(); CommitInterfaceTarget(target, first); });
                    });
                }
                choices.Items.Add("Annuler le déplacement");
                HideInterfaceHandles(); ShowCommandMenu(choices);
                LastMessage = "Plusieurs niveaux coïncident. Choisissez l'ancrage.";
                return;
            }
            CommitInterfaceTarget(hit.Target, first);
        }
        private void CommitInterfaceTarget(WaterfallTarget target, bool first)
        {
            WaterfallConnector connector = spec.Waterfall.Connectors.Find(delegate(WaterfallConnector c) { return c.Id == currentId; });
            if (connector == null || WaterfallInteraction.SameAnchor(first ? connector.Left : connector.Right, target.Anchor)) return;
            bool before = busy; busy = true;
            try { CommitTarget(target, first); }
            finally { busy = before; }
            UpdateInterfaceHandles();
        }
        private void DisposeInterface()
        {
            interfaceInputEnabled = false; CancelInterfaceDrag(); projection = null; editWindow = null;
            if (interfaceHandles != null) { interfaceHandles.Dispose(); interfaceHandles = null; }
            if (editKeys != null) { editKeys.Dispose(); editKeys = null; }
        }
    }
}
