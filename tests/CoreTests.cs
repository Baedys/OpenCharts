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
            Waterfall(); SignedColumn(); Mekko(); Gantt(); Extremes(); Validation(); Series();
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
