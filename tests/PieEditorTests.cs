using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using OpenSlideCharts;

public static class PieEditorTests
{
    private static int checks;
    private static void Check(bool condition, string message) { checks++; if (!condition) throw new Exception(message); }
    private static object Field(object target, string name) { return target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target); }
    private static object Call(object target, string name, params object[] arguments) { return target.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, arguments); }
    private static bool Reject(Action action)
    {
        try { action(); return false; } catch (TargetInvocationException error) { return error.InnerException is ArgumentException; }
    }
    [STAThread]
    public static int Main()
    {
        try
        {
            ChartSpec source = ChartSamples.Create("Doughnut"); source.Title = "Répartition";
            source.Rows.Clear(); source.Rows.Add(new ChartRow { Label = "Europe", Value = 3.125 }); source.Rows.Add(new ChartRow { Label = "Asie", Value = 2 }); source.Rows.Add(new ChartRow { Label = "Autres", Value = 0 });
            source.Pie = new PieOptions { StartAngle = 35.5, HoleRatio = .625, Order = "SheetCounterclockwise", LabelContent = "Both", PercentDigits = 2, CenterTitle = "Total", ShowCategory = false };
            source.Pie.Slices.Add(new PieSliceStyle { Index = 1, Explode = .15, Color = "#123456" });
            source.WorkbookPath = "C:\\synthetic\\source.xlsx"; source.Worksheet = "Données"; source.RangeAddress = "A1:B3";
            string before = ChartStorage.Serialize(source);
            using (var editor = new PieEditor(source))
            {
                ChartSpec read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Kind == "Doughnut" && read.Title == "Répartition", "Chart type and title must roundtrip.");
                Check(read.Rows[0].Value == 3.125 && read.Rows[2].Value == 0, "Decimals and zero must roundtrip.");
                Check(read.Pie.StartAngle == 35.5 && read.Pie.HoleRatio == .625 && read.Pie.CenterTitle == "Total", "Rotation, hole and center title must survive.");
                Check(read.Pie.Order == "SheetCounterclockwise" && read.Pie.LabelContent == "Both" && read.Pie.PercentDigits == 2 && !read.Pie.ShowCategory, "Label and order choices must survive.");
                Check(read.Pie.Slices.Count == 1 && read.Pie.Slices[0].Index == 1 && read.Pie.Slices[0].Explode == .15 && read.Pie.Slices[0].Color == "#123456", "Slice settings must roundtrip.");
                Check(read.WorkbookPath == source.WorkbookPath && read.Worksheet == source.Worksheet && read.RangeAddress == source.RangeAddress, "Excel link must survive data edits.");
                read.Pie.Slices[0].Color = "#FFFFFF"; Check(ChartStorage.Serialize(source) == before, "Unapplied result must not mutate caller.");
                var grid = (DataGridView)Field(editor, "grid"); grid.Rows.RemoveAt(0);
                read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Rows.Count == 2 && read.Rows[0].Label == "Asie" && read.Pie.Slices[0].Index == 0, "Deleting a row must carry its surviving slice style to its new index.");
                Check(read.Pie.Slices[0].Color == "#123456" && read.Pie.Slices[0].Explode == .15, "Row deletion must preserve surviving color and explosion.");
                string unchanged = ChartStorage.Serialize(read);
                Check(Reject(delegate { Call(editor, "PasteData", "A\t-1\nB\t2"); }), "Negative paste must be refused.");
                Check(ChartStorage.Serialize((ChartSpec)Call(editor, "ReadSpec", true)) == unchanged, "Invalid paste must leave previous data intact.");
                Call(editor, "PasteData", "Libellé\tValeur\n\"Nord\tEst\"\t1,2500\nSud\t2");
                read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Rows.Count == 2 && read.Rows[0].Label == "Nord\tEst" && read.Rows[0].Value == 1.25, "Quoted TSV and decimal comma must import.");
                Check(read.Pie.Slices.Count == 0 && read.Pie.StartAngle == 35.5, "Replacing data must clear row styles and retain chart settings.");
                var kind = (ComboBox)Field(editor, "kind"); kind.SelectedIndex = 0;
                read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Kind == "Pie" && read.Pie.HoleRatio == 0 && read.Pie.CenterTitle == "", "Switching to pie must reset hole and center title.");
                kind.SelectedIndex = 1; read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Pie.HoleRatio == .5, "Switching to doughnut must set its original 50 percent default.");
                ((TextBox)Field(editor, "digits")).Text = "1.5";
                Check(Reject(delegate { Call(editor, "ReadSpec", true); }), "Fractional decimal count must be refused.");
                ((TextBox)Field(editor, "digits")).Text = "6";
                Check(((ChartSpec)Call(editor, "ReadSpec", true)).Pie.PercentDigits == 6, "Typed decimal count must remain editable.");
                ((TextBox)Field(editor, "hole")).Text = "100";
                read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(Reject(delegate { Call(editor, "ValidateCandidate", read); }), "Zero-area ring must fail before apply.");
            }
            Check(ChartStorage.Serialize(source) == before, "Closing editor must leave caller unchanged.");
            ChartSpec defaults = ChartSamples.Create("Pie"); defaults.Pie = null;
            using (var editor = new PieEditor(defaults))
            {
                var read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(!read.Pie.HoleRatio.HasValue && read.Kind == "Pie", "Implicit default hole must remain implicit while unchanged.");
                Check(Reject(delegate { Call(editor, "PasteData", "A\t1\textra"); }), "Three-column paste must not silently reinterpret colors.");
            }
            Console.WriteLine("PASS: " + checks + " pie editor assertions (no interactive UI coverage)"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
}
