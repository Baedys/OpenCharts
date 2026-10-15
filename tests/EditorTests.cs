using System;
using System.Collections.Generic;
using System.Reflection;
using OpenSlideCharts;

public static class EditorTests
{
    private static int checks;
    private static void Check(bool condition, string message) { checks++; if (!condition) throw new Exception(message); }
    [STAThread]
    public static int Main()
    {
        try
        {
            var original = new ChartSpec { Kind = "line", Title = "Évolution", LayoutMode = "Clustered" };
            original.Series.AddRange(new string[] { "North", "South" });
            original.Rows.Add(new ChartRow { Label = "A", Values = new List<double?> { 1.234, null } });
            original.Rows.Add(new ChartRow { Label = "B", Values = new List<double?> { 0, -2.5 } });
            string before = ChartStorage.Serialize(original);
            using (var editor = new SeriesEditor(original))
            {
                var method = typeof(SeriesEditor).GetMethod("ReadMatrix", BindingFlags.Instance | BindingFlags.NonPublic);
                var read = (ChartSpec)method.Invoke(editor, null);
                Check(read.Kind == "Line", "Lowercase line must retain its chart type.");
                Check(read.Rows[0].Values[0] == 1.234, "Three decimal places must roundtrip without ambiguity.");
                Check(!read.Rows[0].Values[1].HasValue && read.Rows[1].Values[0] == 0, "Missing and zero must stay distinct.");
                Check(read.Series[1] == "South" && read.Title == "Évolution", "Names must survive editing.");
                read.Rows[0].Values[0] = 99;
            }
            Check(ChartStorage.Serialize(original) == before, "Editor must not mutate caller's data before apply.");
            var legacy = ChartSamples.Create("Waterfall");
            using (var editor = new ChartEditor(legacy))
            {
                var method = typeof(ChartEditor).GetMethod("ReadSpec", BindingFlags.Instance | BindingFlags.NonPublic);
                var read = (ChartSpec)method.Invoke(editor, new object[] { true });
                Check(read.Rows[0].Total && read.Rows[2].Value == -80, "Single-series editor must retain absolute totals and signs.");
            }
            Console.WriteLine("PASS: " + checks + " editor data assertions (no interactive UI coverage)"); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
