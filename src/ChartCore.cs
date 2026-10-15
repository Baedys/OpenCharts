using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenSlideCharts
{
    public sealed class ChartSpec
    {
        public string Kind, Title;
        public List<ChartRow> Rows = new List<ChartRow>();
        public string WorkbookPath, Worksheet, RangeAddress;
        public List<string> Series = new List<string>();
        public string LayoutMode = "Clustered";
        public ChartAppearance Appearance;
        public WaterfallSpec Waterfall;
        public List<string> ImportWarnings = new List<string>();
    }
    public sealed class ChartRow
    {
        public string Label;
        public double Value, Secondary;
        public bool Total;
        public List<double?> Values = new List<double?>();
    }
    public sealed class Primitive
    {
        public string Kind, Text, Alignment;
        public double X, Y, W, H;
        public int Color;
        public double FontSize;
        public string FontName;
        public double LineWidth;
        public int DashStyle;
        public string VerticalAlignment;
        public double? MarginLeft, MarginRight;
        public bool NoWrap;
    }
    public sealed class Layout
    {
        public List<Primitive> Items = new List<Primitive>();
    }

    // Original, dependency-free geometry. All coordinates use points.
    public static partial class ChartEngine
    {
        public const int Blue = 12936747;
        public const int Green = 8295780;
        public const int Red = 5924813;
        public const int Ink = 4734003;
        public const int Gray = 14473684;
        private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

        private sealed class Scale
        {
            public double Unit, Low, High;
            public Scale(double low, double high)
            {
                Unit = Math.Max(Math.Abs(low), Math.Abs(high));
                if (Unit == 0) Unit = 1;
                Low = low / Unit;
                High = high / Unit;
                if (Low == High) { Low = 0; High = 1; }
            }
            public double Fraction(double value) { return (value / Unit - Low) / (High - Low); }
            public double Tick(double fraction) { return (Low + fraction * (High - Low)) * Unit; }
        }

        public static Layout Build(ChartSpec spec, double width, double height)
        {
            Validate(spec, width, height);
            string kind = spec.Kind.ToLowerInvariant();
            Layout result = new Layout();
            if (spec.Appearance == null || spec.Appearance.ShowTitle)
                Text(result, spec.Title ?? spec.Kind, 14, 8, width - 28, 28, 20, Ink, "left");
            if (spec.Waterfall != null) ConstraintWaterfall(result, spec, width, height);
            else if (spec.Series != null && spec.Series.Count > 0) MultiSeries(result, spec, width, height);
            else if (kind == "mekko") Mekko(result, spec, width, height);
            else if (kind == "gantt" || kind == "bar") Horizontal(result, spec, width, height, kind == "gantt");
            else Vertical(result, spec, width, height, kind);
            foreach (Primitive item in result.Items)
            {
                if (!Finite(item.X) || !Finite(item.Y) || !Finite(item.W) || !Finite(item.H) ||
                    item.X < -.00001 || item.Y < -.00001 || item.X > width + .00001 || item.Y > height + .00001 ||
                    item.X + item.W < -.00001 || item.Y + item.H < -.00001 || item.X + item.W > width + .00001 || item.Y + item.H > height + .00001 ||
                    (item.Kind != "line" && (item.W < 0 || item.H < 0)))
                    throw new ArgumentException("Chart geometry exceeds its frame. Leave room around PlotRect for labels and ticks.");
            }
            return result;
        }

        private static void Validate(ChartSpec spec, double width, double height)
        {
            if (spec == null) throw new ArgumentNullException("spec");
            bool observedPlot = spec.Appearance != null && spec.Appearance.PlotRect != null;
            if (!Finite(width) || !Finite(height) || width < (observedPlot ? 1 : 320) || height < (observedPlot ? 1 : 240) || width > 100000 || height > 100000)
                throw new ArgumentOutOfRangeException("width", "Chart dimensions must be finite, from 320 x 240 to 100000 x 100000 points.");
            string kind = (spec.Kind ?? "").ToLowerInvariant();
            if (kind != "waterfall" && kind != "column" && kind != "bar" && kind != "line" && kind != "mekko" && kind != "gantt")
                throw new NotSupportedException("Unsupported chart kind: " + spec.Kind + ". Supported: Waterfall, Column, Bar, Line, Mekko, Gantt.");
            if (spec.Waterfall != null)
            {
                if (kind != "waterfall" || spec.Series != null && spec.Series.Count > 0 || spec.Appearance != null)
                    throw new NotSupportedException("Constraint waterfalls require Kind=Waterfall without a series matrix or imported appearance.");
                return;
            }
            if (spec.Rows == null || spec.Rows.Count == 0 || spec.Rows.Count > 80)
                throw new ArgumentException("Provide between 1 and 80 rows.", "spec");
            double running = 0;
            bool positiveWidth = false;
            foreach (ChartRow row in spec.Rows)
            {
                if (row == null || !Finite(row.Value) || !Finite(row.Secondary))
                    throw new ArgumentException("Every row must contain finite numeric values.", "spec");
                if (kind == "waterfall")
                {
                    // A Total is an absolute level, resets the running balance, and draws from zero.
                    running = row.Total ? row.Value : running + row.Value;
                    if (!Finite(running)) throw new ArgumentException("Waterfall cumulative value overflow.", "spec");
                }
                if (kind == "gantt" && (row.Secondary <= 0 || !Finite(row.Value + row.Secondary)))
                    throw new ArgumentException("Gantt duration must be positive; end must be finite.", "spec");
                if (kind == "mekko")
                {
                    if (row.Value < 0 || row.Secondary < 0 || row.Secondary > 100)
                        throw new ArgumentException("Mekko requires nonnegative widths and shares between 0 and 100.", "spec");
                    positiveWidth |= row.Value > 0;
                }
            }
            if (kind == "mekko" && !positiveWidth) throw new ArgumentException("Mekko requires at least one positive width.", "spec");
        }

        private static void Vertical(Layout layout, ChartSpec spec, double width, double height, string kind)
        {
            double left = 68, top = 60, plotW = width - 88, plotH = height - 112;
            double low = 0, high = 0, running = 0;
            List<double> starts = new List<double>(), ends = new List<double>();
            foreach (ChartRow row in spec.Rows)
            {
                double start = 0, end = row.Value;
                if (kind == "waterfall")
                {
                    start = row.Total ? 0 : running;
                    end = row.Total ? row.Value : running + row.Value;
                    running = end;
                }
                starts.Add(start); ends.Add(end);
                low = Math.Min(low, Math.Min(start, end)); high = Math.Max(high, Math.Max(start, end));
            }
            Scale scale = new Scale(low, high);
            AxisY(layout, scale, left, top, plotW, plotH);
            double zero = top + plotH * (1 - scale.Fraction(0));
            Line(layout, left, zero, plotW, 0, Ink);
            double step = plotW / spec.Rows.Count, barW = step * .64;
            double prevX = 0, prevY = 0;
            for (int i = 0; i < spec.Rows.Count; i++)
            {
                ChartRow row = spec.Rows[i];
                double x = left + step * (i + .5);
                double yStart = top + plotH * (1 - scale.Fraction(starts[i]));
                double yEnd = top + plotH * (1 - scale.Fraction(ends[i]));
                if (kind == "line")
                {
                    if (i > 0) Line(layout, prevX, prevY, x - prevX, yEnd - prevY, Blue);
                    Shape(layout, "ellipse", x - 3, yEnd - 3, 6, 6, Blue);
                }
                else
                {
                    if (kind == "waterfall" && i > 0 && !row.Total)
                        Line(layout, prevX + barW / 2, yStart, step - barW, 0, Gray);
                    int color = kind == "waterfall" ? (row.Total ? Blue : (row.Value < 0 ? Red : Green)) : Blue;
                    if (yStart == yEnd) Line(layout, x - barW / 2, yStart, barW, 0, color);
                    else Shape(layout, "rect", x - barW / 2, Math.Min(yStart, yEnd), barW, Math.Abs(yStart - yEnd), color);
                }
                double labelY = Math.Max(38, Math.Min(yStart, yEnd) - 19);
                if (kind == "line") labelY = Math.Max(38, yEnd - 19);
                Text(layout, Number(row.Value), left + step * i, labelY, step, 18, Font(step), Ink);
                Text(layout, row.Label ?? "", left + step * i, top + plotH + 8, step, 30, Font(step), Ink);
                prevX = x; prevY = yEnd;
            }
        }

        private static void Horizontal(Layout layout, ChartSpec spec, double width, double height, bool gantt)
        {
            double left = Math.Min(130, width * .28), top = 58, plotW = width - left - 40, plotH = height - 106;
            double low = 0, high = 0;
            foreach (ChartRow row in spec.Rows)
            {
                low = Math.Min(low, row.Value);
                high = Math.Max(high, gantt ? row.Value + row.Secondary : row.Value);
            }
            Scale scale = new Scale(low, high);
            for (int tick = 0; tick <= 4; tick++)
            {
                double fraction = tick / 4.0, x = left + fraction * plotW;
                Line(layout, x, top, 0, plotH, Gray);
                Text(layout, Number(scale.Tick(fraction)), x - 30, top + plotH + 8, 60, 20, 9, Ink);
            }
            double zero = left + scale.Fraction(0) * plotW;
            Line(layout, zero, top, 0, plotH, Ink);
            double step = plotH / spec.Rows.Count;
            for (int i = 0; i < spec.Rows.Count; i++)
            {
                ChartRow row = spec.Rows[i];
                double start = gantt ? row.Value : 0, end = gantt ? row.Value + row.Secondary : row.Value;
                double x1 = left + scale.Fraction(start) * plotW, x2 = left + scale.Fraction(end) * plotW;
                double y = top + (i + .2) * step, h = step * .6;
                if (x1 == x2) Line(layout, x1, y, 0, h, Blue);
                else Shape(layout, "rect", Math.Min(x1, x2), y, Math.Abs(x2 - x1), h, gantt || row.Value >= 0 ? Blue : Red);
                Text(layout, row.Label ?? "", 8, y, left - 14, h, Math.Min(11, Math.Max(6, step * .5)), Ink, "left");
                // Keep values inside plotting region; white text when the bar has room.
                string value = gantt ? Number(row.Secondary) + " d" : Number(row.Value);
                double available = Math.Abs(x2 - x1), boxW = Math.Min(65, plotW);
                bool inside = available >= 50 && h >= 12;
                double textX = inside ? Math.Min(x1, x2) + 3 : Math.Max(left, Math.Min(left + plotW - boxW, Math.Max(x1, x2) + 3));
                Text(layout, value, textX, y, inside ? Math.Min(available - 6, boxW) : boxW, h, Math.Min(10, Math.Max(6, step * .45)), inside ? 16777215 : Ink, "left");
            }
        }

        private static void Mekko(Layout layout, ChartSpec spec, double width, double height)
        {
            double left = 60, top = 58, plotW = width - 80, plotH = height - 110, max = 0;
            foreach (ChartRow row in spec.Rows) max = Math.Max(max, row.Value);
            double sum = 0;
            foreach (ChartRow row in spec.Rows) sum += row.Value / max;
            AxisY(layout, new Scale(0, 100), left, top, plotW, plotH);
            double x = left;
            foreach (ChartRow row in spec.Rows)
            {
                double w = plotW * (row.Value / max) / sum, share = row.Secondary / 100;
                if (w > 0)
                {
                    if (share < 1) Shape(layout, "rect", x, top, w, plotH * (1 - share), Gray);
                    if (share > 0) Shape(layout, "rect", x, top + plotH * (1 - share), w, plotH * share, Blue);
                    Line(layout, x, top, 0, plotH, 16777215);
                    if (w >= 24) Text(layout, Number(row.Secondary) + "%", x, top + plotH / 2 - 10, w, 20, Font(w), Ink);
                    Text(layout, row.Label ?? "", x, top + plotH + 8, w, 30, Font(w), Ink);
                }
                x += w;
            }
        }

        private static void AxisY(Layout layout, Scale scale, double left, double top, double width, double height)
        {
            for (int i = 0; i <= 4; i++)
            {
                double f = i / 4.0, y = top + height * (1 - f);
                Line(layout, left, y, width, 0, Gray);
                Text(layout, Number(scale.Tick(f)), 4, y - 9, left - 10, 18, 9, Ink, "right");
            }
        }
        private static bool Finite(double n) { return !Double.IsNaN(n) && !Double.IsInfinity(n); }
        private static double Font(double width) { return width < 35 ? 7 : (width < 60 ? 9 : 11); }
        private static string Number(double value) { return value.ToString("G5", Invariant); }
        private static void Shape(Layout l, string kind, double x, double y, double w, double h, int color)
        {
            l.Items.Add(new Primitive { Kind = kind, X = x, Y = y, W = w, H = h, Color = color });
        }
        private static void Line(Layout l, double x, double y, double dx, double dy, int color) { Shape(l, "line", x, y, dx, dy, color); }
        private static void Text(Layout l, string text, double x, double y, double w, double h, double size, int color, string alignment = "center")
        {
            l.Items.Add(new Primitive { Kind = "text", Text = text, X = x, Y = y, W = w, H = h, Color = color, FontSize = size, Alignment = alignment });
        }
    }

    public static class ChartSamples
    {
        public static ChartSpec Create(string kind)
        {
            ChartSpec spec = new ChartSpec { Kind = kind, Title = "OpenSlideCharts - " + kind };
            string key = (kind ?? "").ToLowerInvariant();
            if (key == "waterfall")
            {
                Add(spec, "Opening", 100, 0, true); Add(spec, "Sales", 65, 0, false);
                Add(spec, "Costs", -80, 0, false); Add(spec, "Closing", 85, 0, true);
            }
            else if (key == "gantt")
            {
                Add(spec, "Research", 0, 4, false); Add(spec, "Build", 3, 7, false); Add(spec, "Review", 9, 3, false);
            }
            else if (key == "mekko")
            {
                Add(spec, "North", 40, 60, false); Add(spec, "South", 35, 35, false); Add(spec, "West", 25, 80, false);
            }
            else
            {
                Add(spec, "Q1", 40, 0, false); Add(spec, "Q2", 65, 0, false);
                Add(spec, "Q3", -20, 0, false); Add(spec, "Q4", 90, 0, false);
            }
            return spec;
        }
        private static void Add(ChartSpec s, string label, double value, double second, bool total)
        { s.Rows.Add(new ChartRow { Label = label, Value = value, Secondary = second, Total = total }); }
    }
}
