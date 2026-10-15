using System;
using System.Collections.Generic;
using OpenSlideCharts;

public static class CoreTests
{
    private static int assertions;
    public static int Main()
    {
        try
        {
            foreach (string kind in new string[] { "Waterfall", "Column", "Bar", "Line", "Mekko", "Gantt" })
            {
                Bounds(ChartEngine.Build(ChartSamples.Create(kind), 640, 360), 640, 360);
                Bounds(ChartEngine.Build(ChartSamples.Create(kind), 320, 240), 320, 240);
            }
            Waterfall(); SignedColumn(); Mekko(); Gantt(); Extremes(); Validation(); Series(); Appearance(); ConstraintCascades();
            Console.WriteLine("PASS: " + assertions + " assertions"); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.ToString()); return 1; }
    }
    private static ChartSpec Spec(string kind, params double[] values)
    {
        ChartSpec s = new ChartSpec { Kind = kind, Title = "Test" };
        foreach (double v in values) s.Rows.Add(new ChartRow { Label = "Row", Value = v });
        return s;
    }
    private static List<Primitive> Rects(Layout l)
    {
        List<Primitive> result = new List<Primitive>();
        foreach (Primitive p in l.Items) if (p.Kind == "rect") result.Add(p);
        return result;
    }
    private static void Waterfall()
    {
        ChartSpec s = Spec("Waterfall", 100, -150, 20, 40, -10);
        s.Rows[0].Total = true; s.Rows[3].Total = true;
        List<Primitive> r = Rects(ChartEngine.Build(s, 640, 360));
        // Levels: 100, -50, -30, reset to 40, 30. Range [-50,100], height 248.
        Near(r[0].Y, 60); Near(r[0].H, 248 * 100.0 / 150);
        Near(r[1].Y, 60); Near(r[1].H, 248);
        Near(r[2].Y, 60 + 248 * 130.0 / 150); Near(r[2].H, 248 * 20.0 / 150);
        Near(r[3].Y, 60 + 248 * 60.0 / 150); Near(r[3].H, 248 * 40.0 / 150);
        Near(r[4].Y, r[3].Y); Near(r[4].H, 248 * 10.0 / 150);
        Check(r[1].Color == ChartEngine.Red && r[3].Color == ChartEngine.Blue, "Waterfall colors");
    }
    private static void SignedColumn()
    {
        List<Primitive> r = Rects(ChartEngine.Build(Spec("Column", -10, 30), 640, 360));
        Near(r[0].Y, r[1].Y + r[1].H); Near(r[1].H, 3 * r[0].H);
    }
    private static void Mekko()
    {
        ChartSpec s = Spec("Mekko", 1, 3); s.Rows[0].Secondary = 25; s.Rows[1].Secondary = 75;
        List<Primitive> r = Rects(ChartEngine.Build(s, 640, 360));
        Check(r.Count == 4, "Mekko blocks"); Near(r[2].W, r[0].W * 3);
        Near(r[0].H, r[1].H * 3); Near(r[3].H, r[2].H * 3);
        Near(r[0].W + r[2].W, 560);
    }
    private static void Gantt()
    {
        ChartSpec s = Spec("Gantt", -5, 0); s.Rows[0].Secondary = 5; s.Rows[1].Secondary = 10;
        List<Primitive> r = Rects(ChartEngine.Build(s, 640, 360));
        Near(r[0].X + r[0].W, r[1].X); Near(r[1].W, 2 * r[0].W);
    }
    private static void Extremes()
    {
        foreach (string kind in new string[] { "Column", "Bar", "Line", "Waterfall" })
        {
            Bounds(ChartEngine.Build(Spec(kind, 0, 0), 640, 360), 640, 360);
            Bounds(ChartEngine.Build(Spec(kind, 1e300, -1e300), 640, 360), 640, 360);
            Bounds(ChartEngine.Build(Spec(kind, 1e-300, -1e-300), 640, 360), 640, 360);
        }
        ChartSpec m = Spec("Mekko", Double.MaxValue, Double.MaxValue);
        m.Rows[0].Secondary = 100; Bounds(ChartEngine.Build(m, 640, 360), 640, 360);
    }
    private static void Validation()
    {
        Reject(delegate { ChartEngine.Build(null, 640, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Pie", 1), 640, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Column"), 640, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Column", Double.NaN), 640, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Column", Double.PositiveInfinity), 640, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Column", 1), Double.NaN, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Column", 1), 100, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Gantt", 1), 640, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Mekko", 0), 640, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Mekko", -1), 640, 360); });
        Reject(delegate { ChartEngine.Build(Spec("Waterfall", Double.MaxValue, Double.MaxValue), 640, 360); });
        ChartSpec badShare = Spec("Mekko", 1); badShare.Rows[0].Secondary = 101;
        Reject(delegate { ChartEngine.Build(badShare, 640, 360); });
        ChartSpec badDuration = Spec("Gantt", 1); badDuration.Rows[0].Secondary = -1;
        Reject(delegate { ChartEngine.Build(badDuration, 640, 360); });
    }
    private static ChartSpec Matrix(string kind, string mode)
    {
        var result = new ChartSpec { Kind = kind, LayoutMode = mode, Title = "Matrix" };
        result.Series.AddRange(new string[] { "A", "B", "C" });
        result.Rows.Add(new ChartRow { Label = "One", Values = new List<double?> { 10, 20, -5 } });
        result.Rows.Add(new ChartRow { Label = "Two", Values = new List<double?> { 30, null, -10 } });
        return result;
    }
    private static void Series()
    {
        foreach (string kind in new string[] { "Column", "Bar", "Line" })
        {
            var spec = Matrix(kind, "Clustered");
            Bounds(ChartEngine.Build(spec, 640, 360), 640, 360);
            Bounds(ChartEngine.Build(spec, 320, 240), 320, 240);
        }
        var stacked = Matrix("Column", "Stacked");
        var bars = Rects(ChartEngine.Build(stacked, 640, 360));
        // Three legend swatches, then five actual segments. Positive segments share
        // a stack; negative segments start at zero, independently of positives.
        Check(bars.Count == 8, "Missing observation creates no segment");
        Near(bars[4].Y + bars[4].H, bars[3].Y);
        Near(bars[5].Y, bars[3].Y + bars[3].H);
        Near(bars[4].H, bars[3].H * 2);
        foreach (string kind in new string[] { "Column", "Bar" })
        {
            stacked.Kind = kind;
            Bounds(ChartEngine.Build(stacked, 320, 240), 320, 240);
        }
        var pct = Matrix("Column", "Percent");
        pct.Rows[0].Values = new List<double?> { Double.MaxValue, Double.MaxValue, 0 };
        pct.Rows[1].Values = new List<double?> { 0, null, 0 };
        bars = Rects(ChartEngine.Build(pct, 640, 360));
        Near(bars[3].H, bars[4].H);
        Bounds(ChartEngine.Build(pct, 320, 240), 320, 240);
        var line = Matrix("Line", "Clustered");
        line.Rows.Add(new ChartRow { Label = "Three", Values = new List<double?> { 40, 50, null } });
        Layout l = ChartEngine.Build(line, 640, 360);
        int secondSeriesLines = 0;
        foreach (Primitive p in l.Items) if (p.Kind == "line" && p.Color == ChartEngine.Green) secondSeriesLines++;
        Check(secondSeriesLines == 0, "Line must not bridge a missing observation");
        Reject(delegate { ChartEngine.Build(Matrix("Column", "Percent"), 640, 360); });
        Reject(delegate { ChartEngine.Build(Matrix("Line", "Stacked"), 640, 360); });
        Reject(delegate { ChartEngine.Build(Matrix("Waterfall", "Stacked"), 640, 360); });
        var mismatch = Matrix("Column", "Clustered"); mismatch.Rows[0].Values.RemoveAt(0);
        Reject(delegate { ChartEngine.Build(mismatch, 640, 360); });
        var overflow = Matrix("Column", "Stacked"); overflow.Rows[0].Values = new List<double?> { Double.MaxValue, Double.MaxValue, null };
        Reject(delegate { ChartEngine.Build(overflow, 640, 360); });
    }
    private static void Appearance()
    {
        var spec = Matrix("Column", "Stacked");
        spec.Rows[0].Values = new List<double?> { 10, 20, 30 };
        spec.Rows[1].Values = new List<double?> { 15, 25, 35 };
        spec.Appearance = new ChartAppearance {
            ShowTitle = false, ShowLegend = false, ShowValueAxis = false, ShowGridlines = false,
            ShowTotals = true, SeriesLabelsAtFirstStack = true,
            SeriesColors = new List<string> { "#102030", "#405060", "#708090" },
            DisplayOrder = new List<int> { 2, 1, 0 },
            PlotRect = new ChartRectangle { X = .15, Y = .1, Width = .8, Height = .8 },
            Minimum = 0, Maximum = 75, GapRatio = .8, ValueFontSize = 14, CategoryFontSize = 14,
            ValueFormat = "#,##0;\"-\"#,##0", TotalFormat = "0", AxisLineWidth = .75
        };
        Layout layout = ChartEngine.Build(spec, 600, 300);
        Bounds(layout, 600, 300);
        List<Primitive> bars = Rects(layout);
        Check(bars.Count == 6 && bars[0].Color == 0x908070, "Observed order and RGB-to-OLE conversion");
        Near(bars[0].H, 96); Near(bars[0].Y + bars[0].H, 270);
        Near(bars[0].W, 240 / 1.8);
        int lines = 0, totals = 0, names = 0;
        foreach (Primitive p in layout.Items)
        {
            if (p.Kind == "line") { lines++; Near(p.LineWidth, .75); }
            if (p.Text == "60" || p.Text == "75") totals++;
            if (p.Text == "A" || p.Text == "B" || p.Text == "C") names++;
            Check(p.Text != "Matrix", "Hidden imported title must stay hidden");
        }
        Check(lines == 1 && totals == 2 && names == 3, "Observed baseline, totals and series labels");
        Check(ChartAppearance.Format(3.5, "0", 0) == "4" && ChartAppearance.Format(-2.5, "0", 0) == "-3", "Integer labels round away from zero");
        spec.Appearance.SourceFrame = new ChartRectangle { Width = 300, Height = 150 };
        spec.Appearance.ValueLabelOffsets.Add(new ChartLabelOffset { SeriesIndex = 0, CategoryIndex = 0, X = 2, Y = -1 });
        Layout shifted = ChartEngine.Build(spec, 600, 300);
        for (int i = 0; i < layout.Items.Count; i++)
            if (layout.Items[i].Kind == "text" && layout.Items[i].Text == "10")
            { Near(shifted.Items[i].X - layout.Items[i].X, 4); Near(shifted.Items[i].Y - layout.Items[i].Y, -2); }
        spec.Appearance.ValueLabelOffsets.Add(new ChartLabelOffset { SeriesIndex = 0, CategoryIndex = 0 });
        Reject(delegate { ChartEngine.Build(spec, 600, 300); });
        spec.Appearance.ValueLabelOffsets.Clear();
        spec.Appearance.DisplayOrder = new List<int> { 0, 0, 1 };
        Reject(delegate { ChartEngine.Build(spec, 600, 300); });
        spec.Appearance.DisplayOrder = new List<int> { 0, 1, 2 }; spec.Appearance.SeriesColors[0] = "blue";
        Reject(delegate { ChartEngine.Build(spec, 600, 300); });
        spec.Appearance.SeriesColors[0] = "#000000"; spec.Appearance.PlotRect.Width = 2;
        Reject(delegate { ChartEngine.Build(spec, 600, 300); });
        spec.Appearance.PlotRect = new ChartRectangle { X = 0, Y = 0, Width = 1, Height = 1 };
        Reject(delegate { ChartEngine.Build(spec, 600, 300); });
    }
    private static void Bounds(Layout l, double w, double h)
    {
        Check(l.Items.Count > 0, "Nonempty layout");
        foreach (Primitive p in l.Items)
        {
            Check(Finite(p.X) && Finite(p.Y) && Finite(p.W) && Finite(p.H), "Finite geometry");
            Check(p.X >= -.00001 && p.Y >= -.00001 && p.X <= w + .00001 && p.Y <= h + .00001, "Origin in bounds");
            Check(p.X + p.W >= -.00001 && p.Y + p.H >= -.00001 && p.X + p.W <= w + .00001 && p.Y + p.H <= h + .00001, "End in bounds");
            if (p.Kind != "line") Check(p.W >= 0 && p.H >= 0, "Nonnegative sizes");
        }
    }
    private static void ConstraintCascades()
    {
        var sample = WaterfallProfiles.Sample();
        Bounds(ChartEngine.Build(sample, 640, 360), 640, 360);
        List<Primitive> segments = Rects(ChartEngine.Build(sample, 640, 360));
        Check(segments[2].Color != segments[3].Color, "Different segment rows must remain visually distinguishable.");
        Bounds(ChartEngine.Build(sample, 320, 240), 320, 240);
        var solved = WaterfallSolver.Solve(sample.Waterfall);
        Near(solved.Columns[3].Output, 106); Near(solved.Columns[5].Output, 101);
        sample.Waterfall.Connectors.RemoveAt(0);
        sample.Waterfall.Pins.Add(WaterfallProfiles.Pin(WaterfallAnchor.AtBase(sample.Waterfall.Columns[1]), 0));
        solved = WaterfallSolver.Solve(sample.Waterfall);
        Near(solved.Columns[3].Output, 6); Near(solved.Columns[5].Output, 1);
        Bounds(ChartEngine.Build(sample, 640, 360), 640, 360);
        var left = new ChartSpec { Kind = "Waterfall", Waterfall = new WaterfallSpec() };
        for (int c = 0; c < 4; c++)
        {
            var column = new WaterfallColumn { Id = "C" + c, Label = "Category " + c };
            column.Segments.Add(new WaterfallSegment { Id = "S" + c, SheetRow = 0, Calculated = c == 0, Value = c == 0 ? (double?)null : c == 1 ? 3 : c == 2 ? 2 : 4 });
            left.Waterfall.Columns.Add(column);
        }
        WaterfallProfiles.Arrange(left.Waterfall, "EndAtZero");
        Near(WaterfallSolver.Solve(left.Waterfall).Columns[0].Output, -9);
        Bounds(ChartEngine.Build(left, 640, 360), 640, 360);
        WaterfallProfiles.Arrange(left.Waterfall, "LeftTotal");
        Near(WaterfallSolver.Solve(left.Waterfall).Columns[0].Output, 9);
        Layout drawn = ChartEngine.Build(left, 640, 360); Bounds(drawn, 640, 360);
        int dashed = 0; foreach (Primitive p in drawn.Items) if (p.Kind == "line" && p.DashStyle == 4) dashed++;
        Check(dashed == 3, "Explicit connector topology must render all dashed connectors.");
        WaterfallProfiles.Arrange(left.Waterfall, "Backward");
        Near(WaterfallSolver.Solve(left.Waterfall).Columns[0].Output, 9);
        left.Waterfall.Connectors.Add(new WaterfallConnector { Id = "free", Left = WaterfallAnchor.AtBase(left.Waterfall.Columns[0]), Right = WaterfallAnchor.AtOutput(left.Waterfall.Columns[3]), Constrain = false });
        drawn = ChartEngine.Build(left, 640, 360);
        bool skew = false; foreach (Primitive p in drawn.Items) if (p.Kind == "line" && p.DashStyle == 4 && p.Color == ChartEngine.Red && p.H != 0) skew = true;
        Check(skew, "Free connector residual must render as a skew line.");
        Bounds(drawn, 640, 360);
    }
    private static bool Finite(double n) { return !Double.IsNaN(n) && !Double.IsInfinity(n); }
    private static void Near(double a, double b) { Check(Math.Abs(a - b) < .00001, "Expected " + b + ", got " + a); }
    private static void Check(bool ok, string message) { assertions++; if (!ok) throw new Exception(message); }
    private static void Reject(Action action)
    {
        bool rejected = false;
        try { action(); } catch (ArgumentException) { rejected = true; } catch (NotSupportedException) { rejected = true; }
        Check(rejected, "Invalid input must be rejected");
    }
}
