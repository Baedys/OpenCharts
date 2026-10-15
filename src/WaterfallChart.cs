using System;
using System.Collections.Generic;

namespace OpenSlideCharts
{
    // Explicit starting arrangements. The constraint graph, not this choice,
    // is the persisted source of truth after users edit connectors or pins.
    public static class WaterfallProfiles
    {
        public static void Arrange(WaterfallSpec model, string profile)
        {
            if (model == null || model.Columns == null || model.Columns.Count == 0) throw new ArgumentException("La cascade doit contenir des colonnes.");
            if (profile != "Forward" && profile != "Backward" && profile != "LeftTotal" && profile != "EndAtZero") throw new ArgumentException("Profil de cascade inconnu.");
            model.Connectors = new List<WaterfallConnector>(); model.Pins = new List<WaterfallPin>();
            foreach (WaterfallColumn column in model.Columns)
                foreach (WaterfallSegment segment in column.Segments)
                    if (segment.Calculated) model.Pins.Add(Pin(WaterfallAnchor.AtBase(column), 0));
            if (profile == "EndAtZero")
            {
                if (!HasEquals(model.Columns[0])) throw new ArgumentException("La fin à zéro exige une première colonne contenant e.");
                for (int i = 1; i < model.Columns.Count; i++) LinkNext(model, model.Columns[i - 1], model.Columns[i]);
                model.Pins.Add(Pin(WaterfallAnchor.AtOutput(model.Columns[model.Columns.Count - 1]), 0));
                return;
            }
            if (profile == "LeftTotal")
            {
                if (model.Columns.Count < 2 || !HasEquals(model.Columns[0])) throw new ArgumentException("Le total à gauche exige une première colonne contenant e et au moins une autre colonne.");
                model.Pins.Add(Pin(WaterfallAnchor.AtBase(model.Columns[1]), 0));
                for (int i = 2; i < model.Columns.Count; i++) LinkNext(model, model.Columns[i - 1], model.Columns[i]);
                model.Connectors.Add(Link(WaterfallAnchor.AtOutput(model.Columns[0]), WaterfallAnchor.AtOutput(model.Columns[model.Columns.Count - 1])));
                return;
            }
            bool reverse = profile == "Backward";
            WaterfallColumn first = model.Columns[reverse ? model.Columns.Count - 1 : 0];
            if (!HasEquals(first)) model.Pins.Add(Pin(WaterfallAnchor.AtBase(first), 0));
            else throw new ArgumentException("Le départ du cumul doit être numérique. Choisissez l'autre sens ou définissez les connecteurs et ancrages.");
            for (int i = 1; i < model.Columns.Count; i++)
            {
                int previous = reverse ? model.Columns.Count - i : i - 1;
                int next = reverse ? model.Columns.Count - i - 1 : i;
                LinkNext(model, model.Columns[previous], model.Columns[next]);
            }
        }
        private static bool HasEquals(WaterfallColumn column)
        { foreach (WaterfallSegment s in column.Segments) if (s.Calculated) return true; return false; }
        private static void LinkNext(WaterfallSpec model, WaterfallColumn previous, WaterfallColumn next)
        { model.Connectors.Add(Link(WaterfallAnchor.AtOutput(previous), HasEquals(next) ? WaterfallAnchor.AtOutput(next) : WaterfallAnchor.AtBase(next))); }
        public static WaterfallConnector Link(WaterfallAnchor left, WaterfallAnchor right)
        { return new WaterfallConnector { Id = "link_" + Guid.NewGuid().ToString("N"), Left = left, Right = right, Constrain = true }; }
        public static WaterfallPin Pin(WaterfallAnchor anchor, double level)
        { return new WaterfallPin { Id = "pin_" + Guid.NewGuid().ToString("N"), Anchor = anchor, Level = level }; }
        public static ChartSpec Sample()
        {
            var spec = new ChartSpec { Kind = "Waterfall", Title = "Cascade : variations et sous-totaux", Waterfall = new WaterfallSpec() };
            double?[][] values = { new double?[] { 100 }, new double?[] { 35 }, new double?[] { -12, -17 }, new double?[] { null }, new double?[] { -5 }, new double?[] { null } };
            string[] names = { "Départ", "Hausse", "Coûts", "Sous-total", "Autres", "Total" };
            for (int c = 0; c < values.Length; c++)
            {
                var column = new WaterfallColumn { Id = "column_" + c, Label = names[c] };
                for (int s = 0; s < values[c].Length; s++) column.Segments.Add(new WaterfallSegment { Id = "segment_" + c + "_" + s, SheetRow = s, Value = values[c][s], Calculated = !values[c][s].HasValue });
                spec.Waterfall.Columns.Add(column);
            }
            Arrange(spec.Waterfall, "Forward");
            return spec;
        }
    }

    public static partial class ChartEngine
    {
        private static void ConstraintWaterfall(Layout layout, ChartSpec spec, double width, double height)
        {
            WaterfallSolution solution = WaterfallSolver.Solve(spec.Waterfall);
            double low = 0, high = 0;
            foreach (WaterfallColumnSolution column in solution.Columns)
            {
                low = Math.Min(low, Math.Min(column.Base, column.Output)); high = Math.Max(high, Math.Max(column.Base, column.Output));
                foreach (WaterfallSegmentSolution s in column.Segments) { low = Math.Min(low, s.Bottom); high = Math.Max(high, s.Top); }
            }
            foreach (WaterfallConnectorSolution c in solution.Connectors)
            { low = Math.Min(low, Math.Min(c.LeftLevel, c.RightLevel)); high = Math.Max(high, Math.Max(c.LeftLevel, c.RightLevel)); }
            var scale = new Scale(low, high);
            double left = 58, top = 64, plotW = width - 82, plotH = height - 142;
            AxisY(layout, scale, left, top, plotW, plotH);
            Line(layout, left, top + plotH * (1 - scale.Fraction(0)), plotW, 0, Ink);
            double step = plotW / solution.Columns.Count, bar = step * .62;
            var indices = new Dictionary<string, int>();
            for (int i = 0; i < solution.Columns.Count; i++) indices.Add(solution.Columns[i].ColumnId, i);
            foreach (WaterfallConnectorSolution connector in solution.Connectors)
            {
                int a = indices[connector.LeftColumnId], b = indices[connector.RightColumnId];
                double x1 = left + step * (a + .5), x2 = left + step * (b + .5);
                double direction = a < b ? 1 : -1;
                x1 += direction * bar / 2; x2 -= direction * bar / 2;
                double y1 = top + plotH * (1 - scale.Fraction(connector.LeftLevel));
                double y2 = top + plotH * (1 - scale.Fraction(connector.RightLevel));
                Line(layout, x1, y1, x2 - x1, y2 - y1, connector.Skew ? Red : Ink);
                layout.Items[layout.Items.Count - 1].LineWidth = .8;
                layout.Items[layout.Items.Count - 1].DashStyle = 4;
            }
            for (int c = 0; c < solution.Columns.Count; c++)
            {
                WaterfallColumnSolution column = solution.Columns[c];
                double x = left + step * (c + .5) - bar / 2;
                foreach (WaterfallSegmentSolution s in column.Segments)
                {
                    double y = top + plotH * (1 - scale.Fraction(s.Top));
                    double bottom = top + plotH * (1 - scale.Fraction(s.Bottom)), h = bottom - y;
                    // Segment color follows its sheet row; sign is carried by geometry.
                    int row = s.SheetRow >= 0 ? s.SheetRow : column.Segments.IndexOf(s);
                    int color = row == 0 ? Green : (row == 1 ? 0x777777 : SeriesColors[(row - 2) % SeriesColors.Length]);
                    if (s.Value == 0) Line(layout, x, y, bar, 0, color);
                    else Shape(layout, "rect", x, y, bar, h, color);
                    string label = Number(spec.Waterfall.ShowSigns ? s.Value : s.Extent);
                    bool outside = column.Segments.Count == 1 && (s.Calculated || c == 0 || c == solution.Columns.Count - 1) || h < 20;
                    double labelY = outside ? (s.Value < 0 ? bottom + 2 : y - 20) : y + (h - 18) / 2;
                    Text(layout, label, x - 5, labelY, bar + 10, 18, Font(step), outside ? Ink : 0xFFFFFF);
                }
                Text(layout, column.Label ?? "", left + step * c, height - 42, step, 30, Font(step), Ink);
            }
        }
    }
}
