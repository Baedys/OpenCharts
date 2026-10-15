using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    // Edits the chart's real native PowerPoint connectors. No helper shapes are
    // inserted into the document; valid endpoint changes update its model.
    public sealed class WaterfallSlideEditor : IDisposable
    {
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int key);
        private sealed class EditLine
        {
            internal string Name, ConnectorId;
            internal object Shape;
            internal int ShapeId = -1;
            internal ConnectorEndpoints Original, Observed;
            internal int StableTicks;
        }
        private readonly dynamic app;
        private readonly Action editData;
        private PowerPointTimer timer;
        private readonly Dictionary<string, EditLine> lines = new Dictionary<string, EditLine>(StringComparer.Ordinal);
        private object slide, chart;
        private ChartSpec spec;
        private WaterfallScene scene;
        private double left, top, width, height;
        private string token, currentId;
        private int chartId;
        private bool busy, choosing, leaveRequested, suppressSelection;
        private ContextMenuStrip menu;
        public bool Active { get { return chart != null; } }
        public int AppliedEdits { get; private set; }
        public int PollCount { get; private set; }
        public string LastMessage { get; private set; }

        public WaterfallSlideEditor(object application, Action editDataAction)
        {
            app = application; editData = editDataAction;
        }
        private void EnsureTimer()
        {
            if (timer != null) return;
            timer = new PowerPointTimer((object)app, delegate
            {
                try { PollCount++; if (leaveRequested && Active) Stop(true); else ProcessPendingEdits(false); }
                catch (Exception error) { LastMessage = Describe(error); if (timer != null) timer.Stop(); }
            }, delegate(Exception error) { LastMessage = Describe(error); });
        }
        public void Start(object slideObject, object chartObject, string connectorId = null)
        {
            bool reuse = Active && SameObject(chart, chartObject);
            object finished = Stop(true); if (reuse) chartObject = finished;
            dynamic group = chartObject;
            ChartSpec candidate = ChartStorage.Read(chartObject);
            if (candidate.Waterfall == null) throw new ArgumentException("Sélectionnez une cascade à connecteurs.");
            if ((int)group.Child != 0 || Math.Abs((double)group.Rotation) > .01)
                throw new NotSupportedException("L'édition des connecteurs exige un groupe sans rotation, directement sur la diapositive.");
            scene = WaterfallScene.Build(candidate.Waterfall, (double)group.Width, (double)group.Height);
            chart = chartObject; slide = slideObject; spec = candidate;
            left = group.Left; top = group.Top; width = group.Width; height = group.Height;
            token = Guid.NewGuid().ToString("N").ToUpperInvariant(); currentId = connectorId; leaveRequested = false;
            bool previousBusy = busy; busy = true;
            try
            {
                EnsureTimer();
                bool tagged = false;
                for (int i = 1; i <= (int)group.GroupItems.Count; i++) if ((string)group.GroupItems.Item(i).Tags.Item("OSC_ROLE") == "WATERFALL_CONNECTOR") tagged = true;
                if (!tagged && scene.Connectors.Count > 0) chart = OfficeRenderer.Replace(slide, chart, spec);
                chartId = (int)((dynamic)chart).Id;
                BuildHandles();
                leaveRequested = false; timer.Start(); LastMessage = "Déplacez une extrémité du connecteur vers un niveau de colonne.";
            }
            catch { Stop(false); throw; }
            finally { busy = previousBusy; }
        }
        private void BuildHandles()
        {
            lines.Clear();
            if (!Active) return;
            dynamic group = chart;
            Layout layout = ChartEngine.Build(spec, width, height);
            for (int i = 1; i <= (int)group.GroupItems.Count; i++)
            {
                dynamic shape = group.GroupItems.Item(i); int index;
                if ((string)shape.Tags.Item("OSC_ROLE") != "WATERFALL_CONNECTOR" ||
                    !Int32.TryParse((string)shape.Tags.Item("OSC_ITEM"), out index) || index < 0 || index >= layout.Items.Count) continue;
                Primitive item = layout.Items[index];
                if (item.Role != "WATERFALL_CONNECTOR" || lines.ContainsKey(item.EntityId)) continue;
                lines.Add(item.EntityId, new EditLine { Name = shape.Name, ConnectorId = item.EntityId, Shape = shape,
                    ShapeId = shape.Id, Original = ConnectorEndpoints.Read((object)shape) });
            }
            if (lines.Count != scene.Connectors.Count) throw new InvalidOperationException("Les connecteurs natifs ne correspondent plus aux données enregistrées. Rééditez les données pour régénérer le graphique.");
            if (currentId == null || !lines.ContainsKey(currentId)) currentId = scene.Connectors.Count == 0 ? null : scene.Connectors[0].Id;
            if (currentId != null && !suppressSelection) ((dynamic)lines[currentId].Shape).Select();
        }
        private object FindLine(int id)
        {
            dynamic group = chart;
            for (int i = 1; i <= (int)group.GroupItems.Count; i++) { dynamic item = group.GroupItems.Item(i); if ((int)item.Id == id) return item; }
            return null;
        }
        private void RestoreChart()
        {
            chart = OfficeRenderer.Replace(slide, chart, spec, new ChartRectangle { X = left, Y = top, Width = width, Height = height });
            chartId = (int)((dynamic)chart).Id;
            scene = WaterfallScene.Build(spec.Waterfall, width, height); BuildHandles();
        }
        public object Stop(bool applyPending = true, bool selectChart = false)
        {
            if (applyPending && Active && !busy && !choosing) ProcessPendingEdits(true);
            if (timer != null) timer.Stop(); bool previous = busy; busy = true;
            try
            {
                if (Active && ChartExists() && FrameUnchanged() && HasNativeChanges()) RestoreChart();
                object selected = chart;
                chart = null; slide = null; spec = null; scene = null; token = null; currentId = null;
                lines.Clear(); choosing = false; leaveRequested = false;
                if (menu != null) { menu.Dispose(); menu = null; }
                if (selectChart && selected != null) ((dynamic)selected).Select();
                return selected;
            }
            finally { busy = previous; }
        }
        public void SelectionChanged(object selectionObject)
        {
            if (!Active || busy || choosing) return;
            try
            {
                dynamic selection = selectionObject;
                if ((int)selection.Type != 2) { leaveRequested = true; return; }
                if ((int)selection.Type == 2 && (int)selection.ShapeRange.Count == 1)
                {
                    OfficeChartSelection chosen = OfficeChartSelection.Resolve(selectionObject);
                    if (chosen != null && SameObject(chosen.Chart, chart))
                    {
                        if (chosen.Primitive != null && chosen.Primitive.Role == "WATERFALL_CONNECTOR" && currentId != chosen.Primitive.EntityId)
                        {
                            string next = chosen.Primitive.EntityId;
                            ProcessPendingEdits(true);
                            if (Active && !choosing && lines.ContainsKey(next))
                            {
                                currentId = next; busy = true; try { ((dynamic)lines[next].Shape).Select(); } finally { busy = false; }
                            }
                        }
                        leaveRequested = false; return;
                    }
                }
                suppressSelection = true; try { Stop(true); } finally { suppressSelection = false; }
            }
            catch (COMException error) { if (!IsBusy(error)) SafeStop(); }
            catch { SafeStop(); }
        }
        public bool ProcessPendingEdits(bool force = true)
        {
            if (!Active || busy || choosing || currentId == null) return false;
            bool beforeSuppress = suppressSelection;
            busy = true;
            try
            {
                if (!ChartExists()) { Stop(false); return false; }
                dynamic group = chart;
                if (!IsActiveSlide()) { force = true; suppressSelection = true; leaveRequested = true; }
                if (!FrameUnchanged() || Math.Abs((double)group.Rotation) > .01)
                { LastMessage = "Session terminée : cadre modifié."; Stop(false); return false; }
                EditLine line = lines[currentId]; dynamic live = FindLine(line.ShapeId);
                // A stationary pause with the mouse button held is still a drag.
                // Replacing its group would destroy PowerPoint's live drag target.
                if (!force && (GetAsyncKeyState(1) & 0x8000) != 0)
                { line.Observed = null; line.StableTicks = 0; return false; }
                if ((object)live == null)
                {
                    ChartSpec candidate = ChartStorage.Clone(spec); WaterfallInteraction.DeleteConnector(candidate.Waterfall, currentId);
                    Apply(candidate); LastMessage = "Connecteur supprimé ; sommes recalculées."; return true;
                }
                ConnectorEndpoints moved = ConnectorEndpoints.Read((object)live);
                if (moved.Difference(line.Original) <= .2) { line.Observed = null; line.StableTicks = 0; return false; }
                if (!force)
                {
                    if (line.Observed == null || moved.Difference(line.Observed) > .1) { line.Observed = moved; line.StableTicks = 0; return false; }
                    if (++line.StableTicks < 2) return false;
                }
                bool first = Math.Max(Math.Abs(moved.X1 - line.Original.X1), Math.Abs(moved.Y1 - line.Original.Y1)) > .2;
                bool second = Math.Max(Math.Abs(moved.X2 - line.Original.X2), Math.Abs(moved.Y2 - line.Original.Y2)) > .2;
                if (first && second) throw new ArgumentException("Déplacez une seule extrémité du connecteur à la fois.");
                WaterfallHit hit = scene.FindNearest((first ? moved.X1 : moved.X2) - left, (first ? moved.Y1 : moved.Y2) - top, 14);
                if (hit.Candidates.Count == 0) throw new ArgumentException("Déposez l'extrémité près d'une base, d'un cumul ou du bord d'un segment.");
                if (hit.IsAmbiguous) { ChooseTarget(hit, first); return false; }
                CommitTarget(hit.Target, first); return true;
            }
            catch (COMException error) { if (!IsBusy(error)) { LastMessage = error.Message; SafeStop(); } return false; }
            catch (Exception error)
            {
                LastMessage = Describe(error);
                if (Active) RestoreChart();
                MessageBox.Show(LastMessage, "OpenSlideCharts", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            finally { busy = false; suppressSelection = beforeSuppress; }
        }
        private void CommitTarget(WaterfallTarget target, bool first)
        {
            ChartSpec candidate = ChartStorage.Clone(spec);
            WaterfallInteraction.Reattach(candidate.Waterfall, currentId, first, target.Anchor);
            Apply(candidate); LastMessage = "Connecteur rattaché ; valeurs e recalculées.";
        }
        private void Apply(ChartSpec candidate)
        {
            // Validate both geometry and metadata before disturbing the source chart.
            ChartEngine.Build(candidate, width, height);
            if (IsActiveSlide()) app.StartNewUndoEntry();
            chart = OfficeRenderer.Replace(slide, chart, candidate);
            chartId = (int)((dynamic)chart).Id;
            spec = candidate; scene = WaterfallScene.Build(candidate.Waterfall, width, height);
            AppliedEdits++; BuildHandles();
        }
        private void ChooseTarget(WaterfallHit hit, bool first)
        {
            choosing = true; bool selected = false;
            var choices = new ContextMenuStrip(); menu = choices;
            foreach (WaterfallTarget item in hit.Candidates)
            {
                WaterfallTarget target = item;
                choices.Items.Add(TargetLabel(target), null, delegate
                {
                    selected = true; choosing = false; busy = true;
                    try { CommitTarget(target, first); } catch (Exception error) { LastMessage = Describe(error); RestoreChart(); }
                    finally { busy = false; }
                });
            }
            choices.Items.Add("Annuler le déplacement");
            choices.Closed += delegate { choosing = false; if (!selected && Active) { bool before = busy; busy = true; suppressSelection = true; try { RestoreChart(); } finally { busy = before; suppressSelection = false; } } };
            choices.Show(Cursor.Position);
            LastMessage = "Plusieurs niveaux coïncident. Choisissez l'ancrage.";
        }
        private string TargetLabel(WaterfallTarget target)
        {
            WaterfallColumn column = spec.Waterfall.Columns.Find(delegate(WaterfallColumn c) { return c.Id == target.ColumnId; });
            string edge = target.Edge == "Base" ? "Base" : target.Edge == "Output" ? "Cumul" : target.Edge == "Top" ? "Sommet" : "Pied";
            int index = target.SegmentId == null ? -1 : column.Segments.FindIndex(delegate(WaterfallSegment s) { return s.Id == target.SegmentId; });
            int row = index < 0 ? -1 : column.Segments[index].SheetRow >= 0 ? column.Segments[index].SheetRow : index;
            return column.Label + " — " + edge + (row < 0 ? "" : " du segment " + (row + 1));
        }
        public void ShowContextMenu(object selectionObject, ref bool cancel)
        {
            if (busy || choosing) return;
            OfficeChartSelection selection = null; string connector = null;
            try
            {
                selection = OfficeChartSelection.Resolve(selectionObject);
                if (selection == null && connector == null) return;
                if (selection != null && selection.Spec.Waterfall == null) return;
                object selectedChart = selection == null ? chart : selection.Chart;
                if (selection != null && selection.Primitive != null && selection.Primitive.Role == "WATERFALL_CONNECTOR") connector = selection.Primitive.EntityId;
                if (menu != null) menu.Dispose(); menu = new ContextMenuStrip();
                string selectedId = connector;
                string selectedSession = Active && selection != null && SameObject(selection.Chart, chart) ? token : null;
                menu.Items.Add("Modifier les données", null, delegate { Run(delegate { object current = Active && token == selectedSession ? Stop(true) : selectedChart; ((dynamic)current).Select(); editData(); }); });
                menu.Items.Add("Modifier les connecteurs", null, delegate { Run(delegate { object current = Active && token == selectedSession ? Stop(true) : selectedChart; Start((object)app.ActiveWindow.View.Slide, current, selectedId); }); });
                if (connector != null)
                    menu.Items.Add("Supprimer ce connecteur", null, delegate { Run(delegate { if (!Active || token != selectedSession) Start((object)app.ActiveWindow.View.Slide, selectedChart, selectedId); ChartSpec candidate = ChartStorage.Clone(spec); WaterfallInteraction.DeleteConnector(candidate.Waterfall, selectedId); busy = true; try { Apply(candidate); } finally { busy = false; } }); });
                if (Active) menu.Items.Add("Terminer l'édition des connecteurs", null, delegate { Run(delegate { Stop(true, true); }); });
                cancel = true; menu.Show(Cursor.Position);
            }
            catch (Exception error) { LastMessage = Describe(error); }
        }
        private void Run(Action action)
        { try { action(); } catch (Exception error) { LastMessage = Describe(error); MessageBox.Show(LastMessage, "OpenSlideCharts", MessageBoxButtons.OK, MessageBoxIcon.Warning); } }
        private static string Describe(Exception error)
        {
            var solve = error as WaterfallSolveException;
            if (solve != null) return solve.Failure == WaterfallSolveFailure.Inconsistent ? "Ancrages incompatibles : les contraintes manuelles sont conservées." : "La modification exige un ancrage supplémentaire. Utilisez l'éditeur des données.";
            return error.Message;
        }
        private static bool IsBusy(COMException error)
        { return error.ErrorCode == unchecked((int)0x80010001) || error.ErrorCode == unchecked((int)0x8001010A); }
        private bool FrameUnchanged()
        {
            dynamic bounds = ((dynamic)chart).GroupItems.Item(1);
            return Math.Abs((double)bounds.Left - left) <= .1 && Math.Abs((double)bounds.Top - top) <= .1 && Math.Abs((double)bounds.Width - width) <= .1 && Math.Abs((double)bounds.Height - height) <= .1;
        }
        private bool IsActiveSlide()
        {
            try { return SameObject((object)app.ActiveWindow.View.Slide, slide); }
            catch (COMException error) { if (IsBusy(error)) throw; return false; }
        }
        private bool ChartExists()
        {
            if (!Active || slide == null) return false;
            try
            {
                dynamic sld = slide;
                for (int i = 1; i <= (int)sld.Shapes.Count; i++) if ((int)sld.Shapes.Item(i).Id == chartId) return true;
                return false;
            }
            catch (COMException error) { if (IsBusy(error)) throw; return false; }
        }
        private bool HasNativeChanges()
        {
            foreach (EditLine line in lines.Values)
            {
                object live = FindLine(line.ShapeId);
                if (live == null) return true;
                try { if (ConnectorEndpoints.Read(live).Difference(line.Original) > .2) return true; }
                catch (NotSupportedException) { return true; }
            }
            return false;
        }
        public void SafeStop()
        { try { Stop(true); } catch (Exception error) { LastMessage = Describe(error); if (timer != null) timer.Stop(); } }
        private static bool SameObject(object first, object second)
        {
            if (Object.ReferenceEquals(first, second)) return true;
            if (first == null || second == null || !Marshal.IsComObject(first) || !Marshal.IsComObject(second)) return false;
            IntPtr a = Marshal.GetIUnknownForObject(first), b = IntPtr.Zero;
            try { b = Marshal.GetIUnknownForObject(second); return a == b; }
            finally { if (b != IntPtr.Zero) Marshal.Release(b); Marshal.Release(a); }
        }
        public void Dispose() { Stop(false); if (timer != null) timer.Dispose(); }
    }
}
