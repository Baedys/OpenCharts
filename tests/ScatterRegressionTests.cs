using System;
using System.Collections.Generic;
using System.Reflection;
using OpenSlideCharts;

public static class ScatterRegressionTests
{
    private static int scenarios, assertions;
    public static int Main()
    {
        try
        {
            KnownLine(); NoisyFit(); WeightedFit(); WeightScaling(); CoordinateScaling(); Filtering(); Groups();
            Degenerate(); ConstantY(); ZeroCorrelation(); HugeCoordinates(); SubnormalCoordinates(); NearlyConstantX(); ExtremeWeights();
            SlopeOverflow(); Permutations(); Predictions(); NoMutation();
            AdjacentAxisClipping(); VirtualBoundClipping(); WideAndTinyClipping(); MissingIntersections();
            WeightedClipping(); RawCoefficientOverflow(); ClipValidation();
            Console.WriteLine("PASS: " + scenarios + " scatter regression scenarios, " + assertions + " assertions"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
    private static ScatterPoint P(double? x, double? y, double? weight = null, string group = null)
    { return new ScatterPoint { X = x, Y = y, Size = weight, Group = group, Label = "synthetic" }; }
    private static List<ScatterPoint> Points(params ScatterPoint[] points) { return new List<ScatterPoint>(points); }
    private static void Check(bool value, string message) { assertions++; if (!value) throw new Exception(message); }
    private static void Near(double actual, double expected, string message, double tolerance = 1e-13)
    { Check(!Double.IsNaN(actual) && !Double.IsInfinity(actual) && Math.Abs(actual - expected) <= tolerance * Math.Max(1, Math.Abs(expected)), message + ": " + actual + " vs " + expected); }
    private static void Valid(ScatterFit fit)
    {
        Check(fit != null, "fit exists"); Check(fit.XUnit > 0 && fit.YUnit > 0, "positive normalizers");
        Check(!Double.IsNaN(fit.Slope) && !Double.IsInfinity(fit.Slope), "finite slope");
        Check(fit.RSquared >= 0 && fit.RSquared <= 1, "bounded R squared");
    }
    private static void KnownLine()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(0, 1), P(1, 3), P(2, 5), P(3, 7)), false); Valid(fit);
        Near(fit.XUnit, 3, "x unit"); Near(fit.YUnit, 7, "y unit"); Near(fit.MeanX, .5, "mean x");
        Near(fit.MeanY, 4.0 / 7, "mean y"); Near(fit.Slope, 6.0 / 7, "normalized slope"); Near(fit.RSquared, 1, "exact line");
        Near(fit.PredictNormalized(0) * fit.YUnit, 1, "raw intercept reconstructed"); Check(fit.Count == 4, "count");
    }
    private static void NoisyFit()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(0, 0), P(1, 2), P(2, 1)), false); Valid(fit);
        Near(fit.Slope, .5, "ordinary slope"); Near(fit.RSquared, .25, "ordinary R squared");
        Near(fit.MeanX, .5, "ordinary x mean"); Near(fit.MeanY, .5, "ordinary y mean");
    }
    private static void WeightedFit()
    {
        // Independent hand calculation: W=6, Sx=8, Sy=7, Sxx=14,
        // Sxy=10, Syy=11. Centered cross=4, xx=20, yy=17.
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(0, 0, 1), P(1, 2, 2), P(2, 1, 3)), true); Valid(fit);
        Near(fit.Slope, .2, "raw Size weighted slope"); Near(fit.MeanX, 2.0 / 3, "weighted x mean");
        Near(fit.MeanY, 7.0 / 12, "weighted y mean"); Near(fit.RSquared, 4.0 / 85, "weighted R squared");
        Check(Math.Abs(fit.Slope - .5) > .1, "weights materially differ from unweighted");
    }
    private static void WeightScaling()
    {
        scenarios++;
        foreach (double scale in new double[] { Double.Epsilon, 1, 1e200, Double.MaxValue / 4 })
        {
            ScatterFit fit = ScatterRegression.Fit(Points(P(0, 0, scale), P(1, 2, 2 * scale), P(2, 1, 3 * scale)), true); Valid(fit);
            Near(fit.Slope, .2, "weight scaling slope"); Near(fit.RSquared, 4.0 / 85, "weight scaling R squared");
        }
    }
    private static void CoordinateScaling()
    {
        scenarios++;
        foreach (double xscale in new double[] { 1e-200, 1, 1e200 })
        foreach (double yscale in new double[] { 1e-200, 1, 1e200 })
        {
            ScatterFit fit = ScatterRegression.Fit(Points(P(-xscale, 3 * yscale), P(0, yscale), P(xscale, -yscale)), false); Valid(fit);
            Near(fit.Slope, -2.0 / 3, "independent axis scales"); Near(fit.RSquared, 1, "scaled exact line");
        }
    }
    private static void Filtering()
    {
        scenarios++; List<ScatterPoint> points = Points(null, P(null, 2, 1), P(1, null, 1), P(Double.NaN, 1, 1),
            P(1, Double.PositiveInfinity, 1), P(0, 0, 1), P(1, 1, 2), P(2, 2, null), P(3, 3, 0), P(4, 4, -1), P(5, 5, Double.NaN), P(6, 6, Double.PositiveInfinity));
        ScatterFit ordinary = ScatterRegression.Fit(points, false), weighted = ScatterRegression.Fit(points, true);
        Valid(ordinary); Valid(weighted); Check(ordinary.Count == 7 && weighted.Count == 2, "coordinate and weight filtering independent");
        Near(ordinary.Slope, 1, "ordinary ignores invalid Size"); Near(weighted.Slope, 1, "weighted eligible points");
    }
    private static void Groups()
    {
        scenarios++; List<ScatterPoint> points = Points(P(0, 0, 1, null), P(1, 1, 1, ""), P(0, 1, 1, "A"), P(1, 3, 1, "A"), P(0, 4, 1, "a"), P(1, 1, 1, "a"));
        ScatterFit all = ScatterRegression.Fit(points, false), empty = ScatterRegression.Fit(points, false, ""), upper = ScatterRegression.Fit(points, false, "A"), lower = ScatterRegression.Fit(points, false, "a");
        Check(all.Count == 6 && empty.Count == 2 && upper.Count == 2 && lower.Count == 2, "null all / empty group / ordinal groups");
        Near(empty.Slope, 1, "null point group maps to empty"); Near(upper.Slope, 2.0 / 3, "upper group slope"); Near(lower.Slope, -.75, "lower group slope");
        Check(ScatterRegression.Fit(points, false, "missing") == null, "empty group result");
    }
    private static void Degenerate()
    {
        scenarios++; Check(ScatterRegression.Fit(null, false) == null, "null points"); Check(ScatterRegression.Fit(Points(), true) == null, "empty points");
        Check(ScatterRegression.Fit(Points(P(1, 1)), false) == null, "one point");
        Check(ScatterRegression.Fit(Points(P(2, 1), P(2, 3)), false) == null, "vertical line");
        Check(ScatterRegression.Fit(Points(P(0, 1), P(-0.0, 2)), false) == null, "signed zero not distinct");
        Check(ScatterRegression.Fit(Points(P(1, 1, 1), P(2, 2, 0)), true) == null, "only one eligible weight");
    }
    private static void ConstantY()
    {
        scenarios++;
        foreach (double y in new double[] { 0, Double.Epsilon, -7, Double.MaxValue })
        {
            ScatterFit fit = ScatterRegression.Fit(Points(P(-1, y, 1), P(1, y, 2)), true); Valid(fit);
            Near(fit.Slope, 0, "constant y slope"); Near(fit.RSquared, 1, "constant y exact-fit convention");
            Near(fit.PredictNormalized(.7), y == 0 ? 0 : y > 0 ? 1 : -1, "constant prediction");
        }
    }
    private static void ZeroCorrelation()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(-1, 1, 1), P(0, 0, 2), P(1, 1, 1)), true); Valid(fit);
        Near(fit.Slope, 0, "zero covariance slope"); Near(fit.RSquared, 0, "nonconstant y with zero explanatory power");
        Near(fit.MeanY, .5, "weighted constant prediction");
    }
    private static void HugeCoordinates()
    {
        scenarios++; double maximum = Double.MaxValue;
        ScatterFit fit = ScatterRegression.Fit(Points(P(-maximum, maximum, maximum), P(0, 0, maximum), P(maximum, -maximum, maximum)), true); Valid(fit);
        Near(fit.Slope, -1, "extreme finite slope"); Near(fit.MeanX, 0, "huge centered mean"); Near(fit.RSquared, 1, "no sums overflow");
    }
    private static void SubnormalCoordinates()
    {
        scenarios++; double tiny = Double.Epsilon;
        ScatterFit fit = ScatterRegression.Fit(Points(P(tiny, tiny), P(2 * tiny, 3 * tiny), P(3 * tiny, 5 * tiny)), false); Valid(fit);
        Near(fit.Slope, 1.2, "subnormal axes retain slope"); Near(fit.RSquared, 1, "subnormal exact line");
        Near(fit.PredictNormalized(1), 1, "subnormal endpoint prediction");
        Check(fit.XUnit == 3 * tiny && fit.YUnit == 5 * tiny, "subnormal scales retained");
    }
    private static void NearlyConstantX()
    {
        scenarios++; double u = Math.Pow(2, -52);
        ScatterFit fit = ScatterRegression.Fit(Points(P(1, 1), P(1 + u, 2), P(1 + 2 * u, 3)), false); Valid(fit);
        Near(fit.Slope, (1 + 2 * u) / (3 * u), "near-constant x avoids cancellation"); Near(fit.RSquared, 1, "near-constant x exact line");
        ScatterFit flat = ScatterRegression.Fit(Points(P(0, 1), P(1, 1 + u), P(2, 1 + 2 * u)), false); Valid(flat);
        Check(flat.Slope > 0, "tiny nonzero slope preserved"); Near(flat.RSquared, 1, "near-constant y exact line");
    }
    private static void ExtremeWeights()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(0, 0, Double.Epsilon), P(1, 1, Double.MaxValue)), true); Valid(fit);
        Check(fit.Count == 2, "underflowing normalized weight still contributes"); Near(fit.Slope, 1, "extreme weight two-point slope"); Near(fit.RSquared, 1, "extreme weight exact line");
        ScatterFit reverse = ScatterRegression.Fit(Points(P(0, 0, Double.MaxValue), P(1, 1, Double.Epsilon)), true); Valid(reverse);
        Near(reverse.Slope, 1, "tiny weight preserves distinct x even at rounded-zero mean"); Near(reverse.RSquared, 1, "tiny weighted moments retained");
    }
    private static void SlopeOverflow()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(-1e-160, -1, Double.MaxValue), P(1e-160, 1, Double.MaxValue), P(1e150, 0, Double.Epsilon)), true);
        Check(fit == null, "unrepresentable normalized slope returns null rather than infinity/NaN");
        ScatterFit large = ScatterRegression.Fit(Points(P(-1e-150, -1, Double.MaxValue), P(1e-150, 1, Double.MaxValue), P(1e150, 0, Double.Epsilon)), true); Valid(large);
        Check(large.Slope > 1e299 && large.Slope < 1.1e300, "large representable slope remains supported");
    }
    private static void Permutations()
    {
        scenarios++; List<ScatterPoint> points = Points(P(-7, 9, 1e-200), P(-1, 3, 2), P(3, -4, 8), P(9, 2, 1e200));
        ScatterFit original = ScatterRegression.Fit(points, true); points.Reverse(); ScatterFit reverse = ScatterRegression.Fit(points, true);
        Check(original.Slope == reverse.Slope && original.MeanX == reverse.MeanX && original.MeanY == reverse.MeanY && original.RSquared == reverse.RSquared, "exact accumulation independent of input order");
    }
    private static void Predictions()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(-1, 2), P(1, -2)), false);
        Near(fit.PredictNormalized(.5), -.5, "normalized prediction"); Near(fit.PredictNormalized(-2), 2, "finite extrapolation");
        foreach (double invalid in new double[] { Double.NaN, Double.PositiveInfinity, Double.NegativeInfinity })
        { bool rejected = false; try { fit.PredictNormalized(invalid); } catch (ArgumentOutOfRangeException) { rejected = true; } Check(rejected, "nonfinite prediction request rejected"); }
    }
    private static void NoMutation()
    {
        scenarios++; ScatterPoint a = P(2, 4, 3, "A"), b = P(1, 3, 2, "B"); a.Color = "#112233"; a.Marker = "Circle";
        List<ScatterPoint> points = Points(a, b); ScatterRegression.Fit(points, true);
        Check(points.Count == 2 && Object.ReferenceEquals(points[0], a) && Object.ReferenceEquals(points[1], b), "list order and identity unchanged");
        Check(a.X == 2 && a.Y == 4 && a.Size == 3 && a.Group == "A" && a.Label == "synthetic" && a.Color == "#112233" && a.Marker == "Circle", "point data and styling unchanged");
    }
    // Reflection keeps the product's clipping helpers internal while this suite
    // can still be compiled against the public product assembly by Build.ps1.
    private static double[] Clip(ScatterFit fit, params double[] axes)
    {
        object[] arguments = new object[axes.Length]; for (int i = 0; i < axes.Length; i++) arguments[i] = axes[i];
        try { return (double[])typeof(ScatterFit).GetMethod("ClipToAxes", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(fit, arguments); }
        catch (TargetInvocationException error) { throw error.InnerException; }
    }
    private static double Raw(ScatterFit fit, string name)
    { return (double)typeof(ScatterFit).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(fit); }
    private static void Segment(double[] actual, double x0, double y0, double x1, double y1, string message)
    {
        Check(actual != null && actual.Length == 4, message + " has segment");
        bool forward = Math.Abs(actual[0] - x0) < 1e-13 && Math.Abs(actual[1] - y0) < 1e-13;
        Near(actual[0], forward ? x0 : x1, message + " x0"); Near(actual[1], forward ? y0 : y1, message + " y0");
        Near(actual[2], forward ? x1 : x0, message + " x1"); Near(actual[3], forward ? y1 : y0, message + " y1");
        foreach (double coordinate in actual) Check(coordinate >= 0 && coordinate <= 1, message + " fraction inside axis");
    }
    private static void AdjacentAxisClipping()
    {
        scenarios++; ScatterFit x = ScatterRegression.Fit(Points(P(1e16, 0), P(1e16 + 2, 1)), false); Valid(x);
        Segment(Clip(x, 1e16, 2, 0, 1, 0, 1, 0, 1), 0, 0, 1, 1, "adjacent large x endpoints");
        Near(Raw(x, "RawSlope"), .5, "exact raw slope near large x"); Near(Raw(x, "RawIntercept"), -5e15, "exact raw intercept near large x");
        ScatterFit y = ScatterRegression.Fit(Points(P(0, 1e16), P(1, 1e16 + 2)), false); Valid(y);
        Segment(Clip(y, 0, 1, 0, 1, 1e16, 2, 0, 1), 0, 0, 1, 1, "adjacent large y endpoints");
        Near(Raw(y, "RawSlope"), 2, "exact raw slope near large y"); Near(Raw(y, "RawIntercept"), 1e16, "exact raw intercept near large y");
        x.MeanX = 0; x.MeanY = 17; x.Slope = 999; x.XUnit = 7; x.YUnit = 9;
        Segment(Clip(x, 1e16, 2, 0, 1, 0, 1, 0, 1), 0, 0, 1, 1, "clip never reuses rounded public coefficients");
    }
    private static void VirtualBoundClipping()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(1e16, 0), P(1e16 + 2, 1)), false);
        // Intersections lie at raw x=1e16+0.5 and 1e16+1.5, neither binary64-representable.
        Segment(Clip(fit, 1e16, 1, 0, 2, 0, 1, .25, .75), .25, 0, .75, 1, "virtual raw bounds retain fractional intersections");
    }
    private static void WideAndTinyClipping()
    {
        scenarios++; double max = Double.MaxValue, tiny = Double.Epsilon;
        ScatterFit wide = ScatterRegression.Fit(Points(P(-max, -max), P(max, max)), false);
        Segment(Clip(wide, max, max, -2, 0, 0, max, -1, 1), 0, 0, 1, 1, "wide raw axis span does not overflow");
        ScatterFit subnormal = ScatterRegression.Fit(Points(P(tiny, 2 * tiny), P(2 * tiny, 4 * tiny)), false);
        Segment(Clip(subnormal, tiny, tiny, 0, 1, 2 * tiny, 2 * tiny, 0, 1), 0, 0, 1, 1, "subnormal centered axes");
    }
    private static void MissingIntersections()
    {
        scenarios++; ScatterFit outside = ScatterRegression.Fit(Points(P(0, 3), P(1, 4)), false);
        Check(Clip(outside, 0, 1, 0, 1, 0, 1, 0, 1) == null, "line outside rectangle");
        ScatterFit tangent = ScatterRegression.Fit(Points(P(0, 1), P(1, 2)), false);
        Check(Clip(tangent, 0, 1, 0, 1, 0, 1, 0, 1) == null, "single corner tangent is degenerate");
        ScatterFit horizontal = ScatterRegression.Fit(Points(P(0, 2), P(1, 2)), false);
        Segment(Clip(horizontal, 0, 1, 0, 1, 0, 1, 0, 4), 0, .5, 1, .5, "constant y crossing");
        Segment(Clip(horizontal, 0, 1, 0, 1, 2, 1, 0, 4), 0, 0, 1, 0, "constant y on rectangle edge");
        Check(Clip(horizontal, 0, 1, 0, 1, 0, 1, 3, 4) == null, "constant y outside rectangle");
    }
    private static void WeightedClipping()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(0, 0, 1), P(1, 2, 2), P(2, 1, 3)), true);
        Segment(Clip(fit, 0, 1, 0, 2, 0, 1, 0, 2), 0, .45, 1, .65, "weighted y=0.2x+0.9");
        Near(Raw(fit, "RawSlope"), .2, "weighted raw slope"); Near(Raw(fit, "RawIntercept"), .9, "weighted raw intercept");
        Near(fit.RSquared, 4.0 / 85, "weighted R squared unchanged");
    }
    private static void RawCoefficientOverflow()
    {
        scenarios++; double tiny = Double.Epsilon, max = Double.MaxValue;
        ScatterFit slope = ScatterRegression.Fit(Points(P(tiny, 0), P(2 * tiny, max)), false); Valid(slope);
        Check(Double.IsPositiveInfinity(Raw(slope, "RawSlope")), "raw slope overflow retained for display fallback");
        Segment(Clip(slope, tiny, tiny, 0, 1, 0, max, 0, 1), 0, 0, 1, 1, "exact clip works despite raw slope infinity");
        ScatterFit intercept = ScatterRegression.Fit(Points(P(1, max), P(2, 0)), false); Valid(intercept);
        Check(Double.IsPositiveInfinity(Raw(intercept, "RawIntercept")), "raw intercept overflow retained for display fallback");
        Segment(Clip(intercept, 1, 1, 0, 1, 0, max, 0, 1), 0, 1, 1, 0, "exact clip works despite raw intercept infinity");
    }
    private static void ClipValidation()
    {
        scenarios++; ScatterFit fit = ScatterRegression.Fit(Points(P(0, 0), P(1, 1)), false);
        double[] valid = { 0, 1, 0, 1, 0, 1, 0, 1 };
        for (int i = 0; i < valid.Length; i++)
        {
            double[] invalid = (double[])valid.Clone(); invalid[i] = i % 2 == 0 ? Double.NaN : Double.PositiveInfinity;
            bool rejected = false; try { Clip(fit, invalid); } catch (ArgumentException) { rejected = true; } Check(rejected, "nonfinite axis transform rejected");
        }
        foreach (int index in new int[] { 1, 5, 3, 7 })
        {
            double[] invalid = (double[])valid.Clone(); invalid[index] = 0;
            bool rejected = false; try { Clip(fit, invalid); } catch (ArgumentException) { rejected = true; } Check(rejected, "zero unit or degenerate bounds rejected");
        }
        bool wrongOrder = false; try { Clip(fit, 0, 1, 2, 1, 0, 1, 0, 1); } catch (ArgumentException) { wrongOrder = true; }
        Check(wrongOrder, "reversal belongs to renderer, not inverted axis bounds");
    }
}
