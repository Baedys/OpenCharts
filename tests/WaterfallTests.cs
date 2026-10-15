using System;
using System.Collections.Generic;
using OpenSlideCharts;

public static class WaterfallTests
{
    private static int assertions, scenarios;
    public static int Main()
    {
        try
        {
            ManualForward(); ManualRemoved(); ManualMixed(); ManualReverse(); ManualLeftBefore(); ManualLeftAfter();
            EqualsWithSiblings(); Recalculate(); FreeGraph(); AllEquals(); Conflict(); ForceBaseline(); Decimal();
            Validation(); Reorder(); SparseSheetRows(); Extremes(); ManyColumns(); ExactConflict(); SoftConnectors(); NoMutation();
            Console.WriteLine("PASS: " + scenarios + " waterfall scenarios, " + assertions + " assertions");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.ToString()); return 1; }
    }

    private static WaterfallSpec Spec(params double?[][] data)
    {
        WaterfallSpec spec = new WaterfallSpec();
        for (int i = 0; i < data.Length; i++)
        {
            WaterfallColumn column = new WaterfallColumn { Id = "c" + i, Label = "Category " + i };
            for (int j = 0; j < data[i].Length; j++)
                column.Segments.Add(new WaterfallSegment { Id = "s" + i + "_" + j, Value = data[i][j], Calculated = !data[i][j].HasValue });
            spec.Columns.Add(column);
        }
        return spec;
    }
    private static double?[] Row(params double?[] values) { return values; }
    private static WaterfallAnchor B(WaterfallSpec s, int i) { return WaterfallAnchor.AtBase(s.Columns[i]); }
    private static WaterfallAnchor O(WaterfallSpec s, int i) { return WaterfallAnchor.AtOutput(s.Columns[i]); }
    private static void Pin(WaterfallSpec s, WaterfallAnchor a, double level)
    {
        s.Pins.Add(new WaterfallPin { Id = "p" + s.Pins.Count, Anchor = a, Level = level });
    }
    private static void Connect(WaterfallSpec s, WaterfallAnchor a, WaterfallAnchor b)
    {
        s.Connectors.Add(new WaterfallConnector { Id = "edge" + s.Connectors.Count, Left = a, Right = b });
    }
    private static void Forward(WaterfallSpec s, int breakAt)
    {
        for (int i = 0; i < s.Columns.Count; i++)
        {
            bool calculated = s.Columns[i].Segments.Exists(delegate(WaterfallSegment v) { return v.Calculated; });
            if (i == 0 || i == breakAt || calculated) Pin(s, B(s, i), 0);
            if (i > 0 && i != breakAt) Connect(s, O(s, i - 1), calculated ? O(s, i) : B(s, i));
        }
    }
    private static WaterfallSpec Original()
    {
        return Spec(Row(41), Row(35), Row(-12,-17), Row((double?)null), Row(-5), Row((double?)null));
    }
    private static void Outputs(WaterfallSolution result, params double[] expected)
    {
        Check(result.Columns.Count == expected.Length, "Column count");
        for (int i = 0; i < expected.Length; i++) Near(result.Columns[i].Output, expected[i]);
    }
    private static void ManualForward()
    {
        scenarios++;
        WaterfallSpec s = Original(); Forward(s, -1);
        WaterfallSolution r = WaterfallSolver.Solve(s);
        Outputs(r,41,76,47,47,42,42);
        Near(r.Columns[3].Segments[0].Value,47); Near(r.Columns[5].Segments[0].Value,42);
        Near(r.Columns[2].Segments[0].Start,76); Near(r.Columns[2].Segments[0].End,64);
        Near(r.Columns[2].Segments[1].Start,64); Near(r.Columns[2].Segments[1].End,47);
        foreach (WaterfallConnectorSolution e in r.Connectors) Check(!e.Skew && e.Residual == 0 && e.Constrained, "Horizontal constraints");
    }
    private static void ManualRemoved()
    {
        scenarios++;
        WaterfallSpec s = Original(); Forward(s, 1);
        Outputs(WaterfallSolver.Solve(s),41,35,6,6,1,1);
    }
    private static void ManualMixed()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(3),Row(5,-2),Row((double?)null)); Forward(s, -1);
        WaterfallSolution r = WaterfallSolver.Solve(s);
        Outputs(r,3,6,6);
        Near(r.Columns[1].Segments[0].Start,3); Near(r.Columns[1].Segments[0].End,8);
        Near(r.Columns[1].Segments[1].Start,3); Near(r.Columns[1].Segments[1].End,1);
        Near(r.Columns[1].Segments[1].Extent,2);
        Check(r.Columns[1].Output > 3 && r.Columns[1].Output < 8, "Sum anchor lies inside positive segment");
    }
    private static void ManualReverse()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row((double?)null),Row(-35),Row(12,17),Row((double?)null),Row(5),Row(42));
        Pin(s,B(s,0),0); Pin(s,B(s,3),0); Pin(s,B(s,5),0);
        Connect(s,O(s,0),O(s,1)); Connect(s,B(s,1),O(s,2)); Connect(s,B(s,2),O(s,3));
        Connect(s,O(s,3),O(s,4)); Connect(s,B(s,4),O(s,5));
        WaterfallSolution r = WaterfallSolver.Solve(s);
        Near(r.Columns[0].Segments[0].Value,41); Near(r.Columns[3].Segments[0].Value,47);
        double[] bases = {0,76,47,0,42,0};
        for (int i = 0; i < bases.Length; i++) Near(r.Columns[i].Base,bases[i]);
        Near(r.Columns[2].Segments[0].Bottom,64); Near(r.Columns[2].Segments[0].Top,76);
        Near(r.Columns[2].Segments[1].Bottom,47); Near(r.Columns[2].Segments[1].Top,64);
    }
    private static void ManualLeftBefore()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row((double?)null),Row(3),Row(2),Row(4));
        Pin(s,B(s,0),0); Pin(s,O(s,3),0);
        for (int i = 1; i < 4; i++) Connect(s,O(s,i-1),B(s,i));
        Outputs(WaterfallSolver.Solve(s),-9,-6,-4,0);
    }
    private static void ManualLeftAfter()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row((double?)null),Row(3),Row(2),Row(4));
        Pin(s,B(s,0),0); Pin(s,B(s,1),0);
        Connect(s,O(s,0),O(s,3)); Connect(s,O(s,1),B(s,2)); Connect(s,O(s,2),B(s,3));
        Outputs(WaterfallSolver.Solve(s),9,3,5,9);
    }
    private static void EqualsWithSiblings()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(15),Row(2,null,3)); Forward(s,-1);
        Near(WaterfallSolver.Solve(s).Columns[1].Segments[1].Value,10);
        s.Columns[0].Segments[0].Value = 1;
        WaterfallColumnSolution r = WaterfallSolver.Solve(s).Columns[1];
        Near(r.Segments[1].Value,-4); Near(r.Output,1); Near(r.Segments[1].Bottom,-4);
        Near(r.Segments[0].Bottom,3); Near(r.Segments[0].Top,5);
    }
    private static void Recalculate()
    {
        scenarios++;
        WaterfallSpec s = Original(); Forward(s,-1); s.Columns[1].Segments[0].Value = 40;
        WaterfallSolution r = WaterfallSolver.Solve(s);
        Near(r.Columns[3].Segments[0].Value,52); Near(r.Columns[5].Segments[0].Value,47);
    }
    private static void FreeGraph()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(3));
        WaterfallSolveException ex = Failure(s,WaterfallSolveFailure.Underdetermined);
        Check(ex.FreeVariables.Count == 1 && ex.FreeVariables[0] == "c0:base", "Free variable diagnostic");
    }
    private static void AllEquals()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row((double?)null),Row((double?)null));
        Pin(s,B(s,0),0); Pin(s,B(s,1),0); Connect(s,O(s,0),O(s,1));
        Failure(s,WaterfallSolveFailure.Underdetermined);
    }
    private static void Conflict()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(3),Row(4)); Pin(s,B(s,0),0); Pin(s,B(s,1),0); Connect(s,O(s,0),O(s,1));
        Failure(s,WaterfallSolveFailure.Inconsistent);
        s.Connectors[0].Constrain = false;
        WaterfallConnectorSolution c = WaterfallSolver.Solve(s).Connectors[0];
        Check(c.Skew && !c.Constrained, "Explicit unconstrained connector skew");
        Near(c.LeftLevel,3); Near(c.RightLevel,4); Near(c.Residual,1);
    }
    private static void ForceBaseline()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(5,-2));
        Pin(s,WaterfallAnchor.ForWeights("c0",0,1),0);
        WaterfallColumnSolution c = WaterfallSolver.Solve(s).Columns[0];
        Near(c.Base,2); Near(c.Segments[1].Bottom,0);
    }
    private static void Decimal()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(0.1),Row(0.2),Row((double?)null)); Forward(s,-1);
        Pin(s,O(s,2),0.3); // Exact decimal constraints, no tolerance required.
        Near(WaterfallSolver.Solve(s).Columns[2].Segments[0].Value,0.3);
    }
    private static void Validation()
    {
        scenarios++;
        Invalid(null); Invalid(new WaterfallSpec());
        WaterfallSpec s = Spec(Row()); Invalid(s);
        s = Spec(Row(null,null)); Invalid(s);
        foreach (double invalid in new double[] { Double.NaN, Double.PositiveInfinity, Double.NegativeInfinity })
        {
            Invalid(Spec(Row(invalid)));
            s = Spec(Row(1)); Pin(s,B(s,0),invalid); Invalid(s);
            s = Spec(Row(1)); Pin(s,WaterfallAnchor.ForWeights("c0",invalid),0); Invalid(s);
        }
        s = Spec(Row(1),Row(2)); s.Columns[1].Id = s.Columns[0].Id; Invalid(s);
        s = Spec(Row(1),Row(2)); s.Columns[1].Segments[0].Id = s.Columns[0].Segments[0].Id; Invalid(s);
        s = Spec(Row(1)); s.Columns[0].Id = " "; Invalid(s);
        s = Spec(Row(1)); s.Columns[0].Segments[0].Id = null; Invalid(s);
        s = Spec(Row(1)); s.Columns[0].Segments[0].SheetRow = -2; Invalid(s);
        s = Spec(Row(1)); s.Columns[0].Segments[0].SheetRow = 8; Invalid(s);
        s = Spec(Row(1,2)); s.Columns[0].Segments[0].SheetRow = 3; s.Columns[0].Segments[1].SheetRow = 3; Invalid(s);
        s = Spec(Row(10,20)); s.Columns[0].Segments[0].SheetRow = 1; Invalid(s);
        s = Spec(Row(10,20)); s.Columns[0].Segments[1].SheetRow = 0; Invalid(s);
        s = Spec(Row(1)); s.Columns[0].Segments[0].Calculated = true; Invalid(s);
        s = Spec(Row((double?)null)); s.Columns[0].Segments[0].Calculated = false; Invalid(s);
        s = Spec(Row(1)); Pin(s,WaterfallAnchor.ForWeights("missing",0),0); Invalid(s);
        s = Spec(Row(1)); Pin(s,WaterfallAnchor.ForWeights("c0",0,1),0); Invalid(s);
        s = Spec(Row(1)); Pin(s,B(s,0),0); s.Pins.Add(s.Pins[0]); Invalid(s);
        s = Spec(Row(1)); Connect(s,O(s,0),O(s,0)); s.Connectors.Add(s.Connectors[0]); Invalid(s);
        s = Spec(Row(1)); s.SegmentOrder = "automatic"; Invalid(s);
        s = Spec(Row(1)); s.Columns[0].Segments = null; Invalid(s);
        s = Spec(Row(1)); s.Columns[0].Segments[0] = null; Invalid(s);
        s = Spec(Row(1)); s.Columns[0] = null; Invalid(s);
        s = Spec(Row(1)); s.Connectors = null; Invalid(s);
        s = Spec(Row(1)); s.Pins = null; Invalid(s);
        s = Spec(Row(1)); s.Connectors.Add(null); Invalid(s);
        s = Spec(Row(1)); s.Pins.Add(null); Invalid(s);
        s = Spec(Row(1)); Pin(s,null,0); Invalid(s);
        s = Spec(Row(1)); Pin(s,B(s,0),0); s.Pins[0].Anchor.Weights = null; Invalid(s);
        s = Spec(Row(1,2,3,4,5,6,7,8,9)); Invalid(s);
        s = new WaterfallSpec(); for (int i=0;i<81;i++) s.Columns.Add(new WaterfallColumn { Id="c"+i }); Invalid(s);
    }
    private static void Reorder()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(12,17,-2,-3)); Pin(s,B(s,0),0);
        WaterfallColumnSolution c = WaterfallSolver.Solve(s).Columns[0];
        Near(c.Segments[0].Bottom,17); Near(c.Segments[1].Bottom,0);
        Near(c.Segments[2].Top,0); Near(c.Segments[3].Top,-2);
        s.SegmentOrder = "SheetBottomToTop";
        c = WaterfallSolver.Solve(s).Columns[0];
        Near(c.Segments[0].Bottom,0); Near(c.Segments[1].Bottom,12);
        Near(c.Segments[2].Top,-3); Near(c.Segments[3].Top,0); Near(c.Output,24);
    }
    private static void SparseSheetRows()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(4,3,-5,-2),Row(7));
        int[] rows = {7,1,6,2};
        for (int i = 0; i < rows.Length; i++) s.Columns[0].Segments[i].SheetRow = rows[i];
        s.Columns[1].Segments[0].SheetRow = 7;
        Pin(s,B(s,0),0); Pin(s,B(s,1),0);
        WaterfallSolution result = WaterfallSolver.Solve(s);
        WaterfallColumnSolution c = result.Columns[0];
        Near(c.Segments[0].Bottom,0); Near(c.Segments[1].Bottom,4);
        Near(c.Segments[2].Top,-2); Near(c.Segments[3].Top,0);
        for (int i = 0; i < rows.Length; i++) Check(c.Segments[i].SheetRow == rows[i], "Sheet row provenance retained");
        Check(result.Columns[1].Segments[0].SheetRow == 7,"Sparse row identity retained across columns");
        Check(c.Segments[0].Id == "s0_0" && s.Columns[0].Segments[0].SheetRow == 7,"Solver does not reorder source or result identity");
        s.SegmentOrder = "SheetBottomToTop";
        c = WaterfallSolver.Solve(s).Columns[0];
        Near(c.Segments[0].Bottom,3); Near(c.Segments[1].Bottom,0);
        Near(c.Segments[2].Top,0); Near(c.Segments[3].Top,-5);
        // Legacy entries fall back to their list index; no implicit row creation.
        s = Spec(Row(2,3)); s.Columns[0].Segments[0].SheetRow = 5; Pin(s,B(s,0),0);
        c = WaterfallSolver.Solve(s).Columns[0];
        Near(c.Segments[0].Bottom,0); Near(c.Segments[1].Bottom,2);
        Check(c.Segments[1].SheetRow == -1,"Legacy provenance preserved");
    }
    private static void Extremes()
    {
        scenarios++;
            foreach (double value in new double[] { 1e300, -1e300, 1e-300, -1e-300, Double.Epsilon, Double.MaxValue, -Double.MaxValue })
        {
            WaterfallSpec s = Spec(Row(value),Row((double?)null)); Forward(s,-1);
            double actual = WaterfallSolver.Solve(s).Columns[1].Segments[0].Value;
            Check(!Double.IsInfinity(actual) && !Double.IsNaN(actual) && actual != 0, "Extreme finite nonzero");
            Check(Math.Abs(actual/value-1) < 1e-12, "Extreme ratio");
        }
        // Numerator and denominator individually overflow double, but their
        // quotient is ordinary. Exact elimination must retain the quotient.
        WaterfallSpec division = Spec(Row((double?)null));
        Pin(division,B(division,0),0);
        Pin(division,WaterfallAnchor.ForWeights("c0",1e-300),1);
        Near(WaterfallSolver.Solve(division).Columns[0].Segments[0].Value,1e300);
        WaterfallSpec rational = Spec(Row(null,1),Row(null,1));
        Pin(rational,B(rational,0),0); Pin(rational,B(rational,1),0);
        Connect(rational,WaterfallAnchor.ForWeights("c0",1,0),WaterfallAnchor.ForWeights("c1",1e-300,1));
        Connect(rational,WaterfallAnchor.ForWeights("c1",1,0),WaterfallAnchor.ForWeights("c0",-1e-300,1));
        WaterfallSolution solved = WaterfallSolver.Solve(rational);
        Near(solved.Columns[0].Segments[0].Value,1); Near(solved.Columns[1].Segments[0].Value,1);
        WaterfallSpec overflow = Spec(Row(Double.MaxValue),Row(Double.MaxValue)); Forward(overflow,-1); Invalid(overflow);
        WaterfallSpec underflow = Spec(Row((double?)null)); Pin(underflow,B(underflow,0),0);
        Pin(underflow,WaterfallAnchor.ForWeights("c0",1e300),1e-300); Invalid(underflow);
    }
    private static void ManyColumns()
    {
        scenarios++;
        double?[][] data = new double?[80][];
        for (int i=0;i<80;i++) data[i] = i==79 ? Row(null,1,2,3,4,5,6,7) : Row(1,-2,3,-4,5,-6,7,-8);
        WaterfallSpec s = Spec(data); Forward(s,-1);
        WaterfallSolution result = WaterfallSolver.Solve(s);
        Check(result.Columns.Count==80,"80 columns supported");
        Near(result.Columns[79].Output,-316); Near(result.Columns[79].Segments[0].Value,-344);
    }
    private static void ExactConflict()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(1)); Pin(s,B(s,0),0); Pin(s,B(s,0),1e-300);
        Failure(s,WaterfallSolveFailure.Inconsistent);
    }
    private static void SoftConnectors()
    {
        scenarios++;
        WaterfallSpec s = Spec(Row(0),Row(1e-300)); Pin(s,B(s,0),0); Pin(s,B(s,1),0); Connect(s,O(s,0),O(s,1));
        s.Connectors[0].Constrain = false;
        WaterfallConnectorSolution result = WaterfallSolver.Solve(s).Connectors[0];
        Check(result.Skew && result.Residual > 0,"Small soft residual is not suppressed");
        s.Connectors[0].Constrain=true; Failure(s,WaterfallSolveFailure.Inconsistent);
    }
    private static void NoMutation()
    {
        scenarios++;
        WaterfallSpec s = Original(); Forward(s,-1); WaterfallSolver.Solve(s);
        Check(!s.Columns[3].Segments[0].Value.HasValue && s.Columns[3].Segments[0].Calculated,"e remains symbolic in input");
        Check(s.Columns[2].Segments[0].Id=="s2_0" && s.Columns[2].Segments[0].Value==-12,"Input IDs/data retained");
    }
    private static WaterfallSolveException Failure(WaterfallSpec spec, WaterfallSolveFailure expected)
    {
        try { WaterfallSolver.Solve(spec); }
        catch (WaterfallSolveException ex) { Check(ex.Failure==expected,"Expected failure kind"); return ex; }
        throw new Exception("Expected " + expected);
    }
    private static void Invalid(WaterfallSpec spec)
    {
        try { WaterfallSolver.Solve(spec); }
        catch (ArgumentException) { assertions++; return; }
        throw new Exception("Invalid model was accepted");
    }
    private static void Near(double actual, double expected)
    {
        double scale = Math.Max(Math.Abs(expected),1);
        Check(!Double.IsNaN(actual) && !Double.IsInfinity(actual) && Math.Abs(actual/scale-expected/scale)<1e-12,"Expected " + expected + ", got " + actual);
    }
    private static void Check(bool value, string message)
    {
        assertions++; if (!value) throw new Exception(message);
    }
}
