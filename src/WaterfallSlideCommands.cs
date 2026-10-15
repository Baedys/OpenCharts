using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    public sealed partial class WaterfallSlideEditor
    {
        private WaterfallTarget pendingSource;
        public bool AddingConnector { get { return Active && pendingSource != null; } }
        private sealed class CommandMenu : ContextMenuStrip
        {
            internal Action EscapeAction;
            protected override bool ProcessDialogKey(Keys keyData)
            {
                if ((keyData & Keys.KeyCode) == Keys.Escape && EscapeAction != null) { EscapeAction(); return true; }
                return base.ProcessDialogKey(keyData);
            }
        }
        private sealed class CommandSubmenu : ToolStripDropDownMenu
        {
            internal Action EscapeAction;
            protected override bool ProcessDialogKey(Keys keyData)
            {
                if ((keyData & Keys.KeyCode) == Keys.Escape && EscapeAction != null) { EscapeAction(); return true; }
                return base.ProcessDialogKey(keyData);
            }
        }

        public void BeginSegmentConnector(object slideObject, object chartObject, string columnId, string segmentId, string edge)
        {
            StartForSegment(slideObject, chartObject);
            try
            {
                pendingSource = SegmentTarget(columnId, segmentId, edge);
                SelectSegment(columnId, segmentId);
                LastMessage = "Sélectionnez le segment d'arrivée du connecteur.";
            }
            catch { Stop(false); throw; }
        }
        public void CancelConnectorCreation() { Stop(false, true); }
        public void CompleteSegmentConnector(object chartObject, string columnId, string segmentId, string edge, bool constrain = true)
        {
            RequireCommandChart();
            if (!SameObject(chartObject, chart)) throw new InvalidOperationException("Choisissez une destination dans le même graphique.");
            if (!AddingConnector) throw new InvalidOperationException("Choisissez d'abord l'origine du connecteur.");
            WaterfallTarget source = pendingSource, target = SegmentTarget(columnId, segmentId, edge);
            if (WaterfallInteraction.SameAnchor(source.Anchor, target.Anchor)) throw new ArgumentException("Choisissez deux niveaux distincts.");
            WaterfallCommandAnalysis analysis = WaterfallCommands.AnalyzeAddConnector(spec.Waterfall, source.Anchor, target.Anchor, constrain);
            ExecuteWithRepair(analysis, delegate(string release)
            {
                CloseCommandMenu();
                RequireCommandChart();
                ChartSpec candidate = ChartStorage.Clone(spec);
                WaterfallCommandResult result = WaterfallCommands.AddConnector(candidate.Waterfall, source.Anchor, target.Anchor, constrain, release);
                string previousId = currentId;
                pendingSource = null; currentId = result.ConnectorId;
                bool before = busy; busy = true;
                try { Apply(candidate); LastMessage = "Connecteur ajouté ; valeurs e recalculées."; }
                catch { pendingSource = source; currentId = previousId; throw; }
                finally { busy = before; }
            });
        }
        public void PinSegmentZero(object slideObject, object chartObject, string columnId, string segmentId, string edge)
        {
            StartForSegment(slideObject, chartObject);
            WaterfallTarget target = SegmentTarget(columnId, segmentId, edge);
            WaterfallCommandAnalysis analysis = WaterfallCommands.AnalyzePinZero(spec.Waterfall, target.Anchor);
            ExecuteWithRepair(analysis, delegate(string release)
            {
                CloseCommandMenu();
                RequireCommandChart();
                ChartSpec candidate = ChartStorage.Clone(spec);
                WaterfallCommands.PinZero(candidate.Waterfall, target.Anchor, release);
                bool before = busy, beforeSuppress = suppressSelection;
                busy = true; suppressSelection = true;
                try { Apply(candidate); SelectSegment(columnId, segmentId); LastMessage = "Extrémité du segment ancrée à zéro."; }
                finally { busy = before; suppressSelection = beforeSuppress; }
            });
        }
        private void StartForSegment(object slideObject, object chartObject)
        {
            bool before = suppressSelection; suppressSelection = true;
            try { Start(slideObject, chartObject); }
            finally { suppressSelection = before; }
        }
        private WaterfallTarget SegmentTarget(string columnId, string segmentId, string edge)
        {
            RequireCommandChart();
            WaterfallTarget result = scene.Targets.Find(delegate(WaterfallTarget t)
            { return t.ColumnId == columnId && t.Edge == edge && t.Side == "Left" && ((edge == "Base" || edge == "Output") || t.SegmentId == segmentId); });
            if (result == null) throw new ArgumentException("Ce niveau de segment n'existe plus.");
            return result;
        }
        private void SelectSegment(string columnId, string segmentId)
        {
            bool before = busy; busy = true;
            try
            {
                Layout layout = ChartEngine.Build(spec, width, height); dynamic group = chart;
                for (int i = 1; i <= (int)group.GroupItems.Count; i++)
                {
                    dynamic shape = group.GroupItems.Item(i); int index;
                    if (!Int32.TryParse((string)shape.Tags.Item("OSC_ITEM"), out index) || index < 0 || index >= layout.Items.Count) continue;
                    Primitive item = layout.Items[index];
                    if (item.Role == "WATERFALL_SEGMENT" && item.ColumnId == columnId && item.SegmentId == segmentId) { shape.Select(); return; }
                }
                throw new InvalidOperationException("Le segment natif n'est plus disponible.");
            }
            finally { busy = before; }
        }
        private void RequireCommandChart()
        {
            if (!Active || !ChartExists() || !IsActiveSlide() || !FrameUnchanged() || !Upright(chart))
                throw new InvalidOperationException("La sélection ou le cadre a changé. Recommencez la commande.");
        }
        private void PollConnectorCreation()
        {
            if (!Active || busy) return;
            if (!ChartExists() || !IsActiveSlide() || !FrameUnchanged() || !Upright(chart) || (GetAsyncKeyState(27) & 0x8000) != 0)
            { bool before = suppressSelection; suppressSelection = true; try { Stop(false); } finally { suppressSelection = before; } return; }
            if (!choosing)
            {
                OfficeChartSelection selected = OfficeChartSelection.Resolve((object)app.ActiveWindow.Selection);
                if (selected == null || !SameObject(selected.Chart, chart))
                { bool before = suppressSelection; suppressSelection = true; try { Stop(false); } finally { suppressSelection = before; } }
            }
        }
        private void CreationSelectionChanged(object selectionObject)
        {
            OfficeChartSelection selected = OfficeChartSelection.Resolve(selectionObject);
            if (selected == null || !SameObject(selected.Chart, chart))
            { bool before = suppressSelection; suppressSelection = true; try { Stop(false); } finally { suppressSelection = before; } return; }
            if (selected.Primitive == null || selected.Primitive.Role != "WATERFALL_SEGMENT") return;
            Primitive destination = selected.Primitive; object destinationChart = selected.Chart; string session = token;
            var choices = new CommandMenu();
            AddDestination(choices, destinationChart, session, destination, "Top", "Sommet du segment");
            AddDestination(choices, destinationChart, session, destination, "Bottom", "Pied du segment");
            AddDestination(choices, destinationChart, session, destination, "Output", "Cumul de la colonne");
            choices.Items.Add("Annuler l'ajout", null, delegate { Run(delegate { CancelConnectorCreation(); }); });
            ShowCommandMenu(choices);
            LastMessage = "Choisissez le niveau d'arrivée et le type de connexion.";
        }
        private void AddDestination(ContextMenuStrip choices, object destinationChart, string session, Primitive destination, string edge, string label)
        {
            var entry = new ToolStripMenuItem(label); entry.DropDown = new CommandSubmenu();
            entry.DropDownItems.Add("Égaliser les niveaux", null, delegate { if (!Active || token != session) return; Run(delegate { choosing = false; CompleteSegmentConnector(destinationChart, destination.ColumnId, destination.SegmentId, edge, true); }); });
            entry.DropDownItems.Add("Connecteur libre", null, delegate { if (!Active || token != session) return; Run(delegate { choosing = false; CompleteSegmentConnector(destinationChart, destination.ColumnId, destination.SegmentId, edge, false); }); });
            choices.Items.Add(entry);
        }
        private void ExecuteWithRepair(WaterfallCommandAnalysis analysis, Action<string> commit)
        {
            if (analysis.StrictSucceeds) { commit(null); return; }
            if (analysis.AutoPinIds.Count == 1) { commit(analysis.AutoPinIds[0]); return; }
            if (analysis.AutoPinIds.Count == 0)
                throw new InvalidOperationException(analysis.StrictFailure == WaterfallSolveFailure.Inconsistent
                    ? "Ancrages incompatibles : modifiez les ancrages ou choisissez un connecteur libre."
                    : "La commande exige un ancrage supplémentaire dans l'éditeur des données.");
            var choices = new CommandMenu(); string session = token;
            foreach (string item in analysis.AutoPinIds)
            {
                string id = item; WaterfallPin pin = spec.Waterfall.Pins.Find(delegate(WaterfallPin p) { return p.Id == id; });
                choices.Items.Add("Déplacer l'ancrage automatique : " + AnchorLabel(pin.Anchor), null, delegate
                {
                    if (!Active || token != session) return;
                    Run(delegate { choosing = false; commit(id); });
                });
            }
            choices.Items.Add("Conserver les ancrages");
            ShowCommandMenu(choices);
            LastMessage = "Plusieurs ancrages automatiques peuvent être déplacés. Choisissez lequel.";
        }
        private string AnchorLabel(WaterfallAnchor anchor)
        {
            WaterfallTarget target = scene.Targets.Find(delegate(WaterfallTarget t) { return WaterfallInteraction.SameAnchor(t.Anchor, anchor); });
            if (target != null) return TargetLabel(target);
            WaterfallColumn column = spec.Waterfall.Columns.Find(delegate(WaterfallColumn c) { return c.Id == anchor.ColumnId; });
            return column.Label + " — niveau lié";
        }
        private void ShowCommandMenu(ContextMenuStrip choices)
        {
            CloseCommandMenu();
            menu = choices; choosing = true;
            BindCommandEscape(choices);
            choices.Closed += delegate(object sender, ToolStripDropDownClosedEventArgs args)
            {
                if (!Object.ReferenceEquals(menu, choices)) return;
                menu = null; choosing = false;
                if (args.CloseReason == ToolStripDropDownCloseReason.Keyboard && AddingConnector) Run(CancelConnectorCreation);
                DisposeMenuLater(choices);
            };
            choices.Show(Cursor.Position);
        }
        private void BindCommandEscape(ToolStripDropDown dropdown)
        {
            Action escape = delegate { Run(delegate { if (AddingConnector) CancelConnectorCreation(); else CloseCommandMenu(); }); };
            var root = dropdown as CommandMenu; if (root != null) root.EscapeAction = escape;
            var child = dropdown as CommandSubmenu; if (child != null) child.EscapeAction = escape;
            foreach (ToolStripItem item in dropdown.Items)
            {
                var entry = item as ToolStripDropDownItem;
                if (entry != null && entry.HasDropDownItems) BindCommandEscape(entry.DropDown);
            }
        }
        private void CloseCommandMenu()
        {
            ContextMenuStrip previous = menu; menu = null; choosing = false;
            if (previous != null) { previous.Close(); DisposeMenuLater(previous); }
        }
        private static void DisposeMenuLater(ContextMenuStrip choices)
        {
            // Closed precedes both SetVisibleCore's cleanup and the item's Click.
            // Disposing in that callback can make WinForms recreate a disposed handle.
            if (choices.IsDisposed) return;
            if (choices.IsHandleCreated) choices.BeginInvoke((MethodInvoker)delegate { choices.Dispose(); });
            else choices.Dispose();
        }
    }
}
