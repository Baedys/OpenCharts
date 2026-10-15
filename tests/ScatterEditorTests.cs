using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using OpenSlideCharts;

public static class ScatterEditorTests
{
    private static int checks;
    private static void Check(bool condition, string message) { checks++; if (!condition) throw new Exception(message); }
    private static object Field(object target, string name) { return target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target); }
    private static object Call(object target, string name, params object[] args) { return target.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, args); }
    private static bool Reject(Action action) { try { action(); return false; } catch (TargetInvocationException e) { return e.InnerException is ArgumentException || e.InnerException is NotSupportedException; } }
    private static ChartSpec Sample()
    {
        var spec = new ChartSpec { Kind = "Scatter", Title = "Test synthétique", Scatter = new ScatterSpec() };
        spec.Scatter.Points.Add(new ScatterPoint { Label = "Nord", X = 1.234, Y = -2, Group = "", Color = "#123456", Marker = "Square" });
        spec.Scatter.Points.Add(new ScatterPoint { Label = "Sud", X = 3, Y = 4, Group = "" });
        spec.Scatter.Points.Add(new ScatterPoint { Label = "Absent", X = null, Y = 5, Group = "A" });
        spec.Scatter.XAxis.Minimum = -1; spec.Scatter.XAxis.Maximum = 8; spec.Scatter.YAxis.Reversed = true;
        spec.Scatter.LabelContent = "Label"; spec.Scatter.SecondLabelContent = "XY";
        spec.Scatter.Trendlines.Add(new ScatterTrendline { Group = "", ShowRSquared = false, ShowEquation = false });
        spec.Scatter.Partitions.Add(new ScatterPartition { X1 = 0, Y1 = 1, X2 = 4, Y2 = 1, Label = "Seuil" });
        spec.WorkbookPath = "C:\\synthetic\\source.xlsx"; spec.Worksheet = "Données"; spec.RangeAddress = "A1:G3";
        return spec;
    }
    [STAThread]
    public static int Main()
    {
        try
        {
            ChartSpec original = Sample(); string before = ChartStorage.Serialize(original);
            using (var editor = new ScatterEditor(original))
            {
                ChartSpec read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Kind == "Scatter" && read.Scatter.Points.Count == 3 && read.Scatter.Points[0].X == 1.234 && !read.Scatter.Points[2].X.HasValue, "Decimals and missing coordinates must roundtrip.");
                Check(read.Scatter.Points[0].Color == "#123456" && read.Scatter.Points[0].Marker == "Square", "Point styles must survive unchanged editing.");
                Check(read.Scatter.XAxis.Minimum == -1 && read.Scatter.XAxis.Maximum == 8 && read.Scatter.YAxis.Reversed, "Axis settings must survive.");
                Check(read.Scatter.SecondLabelContent == "XY" && !read.Scatter.ShowLabels.HasValue, "Two-label content and automatic threshold must survive.");
                Check(read.Scatter.Trendlines.Count == 1 && read.Scatter.Trendlines[0].Group == "" && !read.Scatter.Trendlines[0].ShowRSquared && !read.Scatter.Trendlines[0].ShowEquation, "A trend without group text or annotations must not disappear.");
                Check(read.Scatter.Partitions[0].Y1 == 1 && read.WorkbookPath == original.WorkbookPath, "Partition and workbook link must survive.");
                read.Scatter.Points[0].Color = "#FFFFFF"; Check(ChartStorage.Serialize(original) == before, "Read result must not alias caller.");
                var data = (DataGridView)Field(editor, "data"); data.Rows.RemoveAt(1);
                read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Scatter.Points.Count == 2 && read.Scatter.Points[0].Color == "#123456" && read.Scatter.Points[1].Label == "Absent", "Deleting a row must preserve surviving point styles.");
                data.Rows[0].Cells[1].Value = "invalid numeric text";
                ChartSpec excelSource = (ChartSpec)Call(editor, "ReadExcelSource");
                ChartSpec imported = ScatterData.ReadCells(excelSource, new object[,] { { "N", 2, 3 }, { "A", 4, 5 } });
                Check(imported.Scatter.Points[0].Color == "#123456" && imported.Scatter.Points[0].Marker == "Square", "Excel must preserve styles without parsing numeric text it replaces.");
                data.Rows[0].Cells[1].Value = "1.2340";
                string current = ChartStorage.Serialize((ChartSpec)Call(editor, "ReadSpec", true));
                Check(Reject(delegate { Call(editor, "PasteData", "Bad\t1\t2\t-3"); }), "Negative bubble sizes must be rejected.");
                Check(current == ChartStorage.Serialize((ChartSpec)Call(editor, "ReadSpec", true)), "Invalid paste must leave grid data intact.");
                Call(editor, "PasteData", "Libellé\tX\tY\tTaille\tGroupe\n\"Nord\tEst\"\t1,2500\t2\t4\tB\nSud\t3\t4\t1\tB");
                read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Kind == "Bubble" && read.Scatter.Points[0].X == 1.25 && read.Scatter.Points[0].Label == "Nord\tEst", "TSV Size column must convert to Bubble without losing quoted labels.");
                Check(read.Scatter.Points[0].Color == null && read.Scatter.Points[0].Marker == null, "Complete pasted data replacement must not retain stale point styles.");
                Call(editor, "PasteData", "P\t1\t2\nQ\t3\t4");
                read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Kind == "Scatter" && !read.Scatter.Points[0].Size.HasValue, "Removing the Size column must return to Scatter.");
                ((CheckBox)Field(editor, "xLog")).Checked = true;
                read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(Reject(delegate { Call(editor, "ValidateCandidate", read); }), "Logarithmic axes with partitions/trends must fail before closing.");
            }
            Check(ChartStorage.Serialize(original) == before, "Closing editor must not mutate caller.");
            var defaults = Sample(); defaults.Scatter.XAxis = null; defaults.Scatter.YAxis = null;
            using (var editor = new ScatterEditor(defaults))
            {
                var read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Scatter.XAxis != null && !read.Scatter.XAxis.Minimum.HasValue && !read.Scatter.YAxis.Reversed, "Null axis options accepted by the core must also be editable as defaults.");
                Check(defaults.Scatter.XAxis == null && defaults.Scatter.YAxis == null, "Opening null axis defaults must not mutate source.");
            }
            var groups = Sample(); groups.Scatter.Points[0].Group = " A"; groups.Scatter.Points[1].Group = "A"; groups.Scatter.Trendlines[0].Group = " A";
            using (var editor = new ScatterEditor(groups))
            {
                var read = (ChartSpec)Call(editor, "ReadSpec", true);
                Check(read.Scatter.Points[0].Group == " A" && read.Scatter.Points[1].Group == "A" && read.Scatter.Trendlines[0].Group == " A", "Group identity, including whitespace, must survive editing.");
                Check(((ChartSpec)Call(editor, "ReadExcelSource")).Scatter.Points[0].Group == " A", "Excel source styles must preserve exact group identity.");
            }
            DataReplacement(); StorageBounds();
            Console.WriteLine("PASS: " + checks + " scatter editor/data/storage assertions (no interactive UI coverage)"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
    private static void DataReplacement()
    {
        ChartSpec source = Sample(); string before = ChartStorage.Serialize(source);
        ChartSpec read = ScatterData.ReadCells(source, new object[,] { { "N", 10, 20 }, { "S", 30, 40 }, { "M", null, 5 } });
        Check(read.Scatter.Points[0].Color == "#123456" && read.Scatter.Points[0].Marker == "Square", "Same-count source must preserve styles for omitted columns.");
        Check(ChartStorage.Serialize(source) == before, "Excel data parsing must not mutate source.");
        read = ScatterData.ReadCells(source, new object[,] { { "N", 10, 20 }, { "S", 30, 40 }, { null, null, null } });
        Check(read.Scatter.Points.Count == 2 && read.Scatter.Points[0].Color == null && read.ImportWarnings.Count == 1, "Changed-count source must reset positional styles with warning.");
        read = ScatterData.ReadCells(source, new object[,] { { "N", 1, 2, 4, "A", null, null }, { "S", 2, 3, 1, "A", "#ABCDEF", "Circle" } });
        Check(read.Kind == "Bubble" && read.Scatter.Points[0].Color == null && read.Scatter.Points[1].Color == "#ABCDEF", "Explicit style columns must replace prior values, including blanks.");
        var sizesOnly = new List<ScatterPoint> { new ScatterPoint { X = 1, Y = 2, Size = 1 }, new ScatterPoint { Size = 100 } };
        source.Scatter.MarkerScheme = "Shapes";
        read = ScatterData.ReplacePoints(source, sizesOnly);
        Check(read.Kind == "Bubble" && read.Scatter.MarkerScheme == "Colors" && read.Scatter.Points[1].Size == 100, "Size-only scale reference and marker scheme conversion must survive.");
        sizesOnly[0].X = 999; Check(read.Scatter.Points[0].X == 1, "Replacement point list must be deeply copied.");
        Array oneBased = Array.CreateInstance(typeof(object), new int[] { 2, 3 }, new int[] { 1, 1 });
        oneBased.SetValue("A", 1, 1); oneBased.SetValue(1, 1, 2); oneBased.SetValue(2, 1, 3);
        oneBased.SetValue("B", 2, 1); oneBased.SetValue(3, 2, 2); oneBased.SetValue(4, 2, 3);
        read = ScatterData.ReadCells(source, (object[,])oneBased);
        Check(read.Scatter.Points.Count == 2 && read.Scatter.Points[1].Y == 4, "Excel's one-based Value2 matrix must parse.");
        read = ScatterData.ReadCells(source, new object[,] { { "N", 1, 2, null, " A" }, { "S", 3, 4, null, "A" } });
        Check(read.Scatter.Points[0].Group == " A" && read.Scatter.Points[1].Group == "A", "Excel groups must retain significant whitespace.");
    }
    private static void StorageBounds()
    {
        var spec = new ChartSpec { Kind = "Scatter", Scatter = new ScatterSpec() };
        for (int i = 0; i < 1000; i++) spec.Scatter.Points.Add(new ScatterPoint { Label = "Point " + i, Group = "Synthetic", X = i, Y = i * 2 });
        var shape = new FakeShape(); ChartStorage.Write(shape, spec);
        Check(Int32.Parse(shape.Tags.Item("OSC_COUNT")) > 889, "Large scatter fixture must cross the previous metadata bound.");
        Check(ChartStorage.Serialize(ChartStorage.Read(shape)) == ChartStorage.Serialize(spec), "1000-point model must roundtrip within the new metadata budget.");
        shape.Tags.Add("OSC_COUNT", "11652"); bool badCount = false;
        try { ChartStorage.Read(shape); } catch (System.IO.InvalidDataException) { badCount = true; }
        Check(badCount, "Corrupt metadata count must stay bounded.");
        var oversized = new ChartSpec { Kind = "Scatter", Scatter = new ScatterSpec() };
        oversized.Scatter.Points.Add(new ScatterPoint { X = 1, Y = 2, Label = new string('A', 1060000) });
        var untouched = new FakeShape(); bool rejected = false;
        try { ChartStorage.Write(untouched, oversized); } catch (ArgumentException) { rejected = true; }
        Check(rejected && untouched.Tags.Count == 0, "Oversized metadata must be refused before writing tags.");
    }
}
public sealed class FakeShape
{
    public FakeTags Tags = new FakeTags();
    public string AlternativeText;
}
public sealed class FakeTags
{
    private readonly Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    public int Count { get { return values.Count; } }
    public void Add(string name, string value) { values[name] = value.ToUpperInvariant(); }
    public string Item(string name) { string value; return values.TryGetValue(name, out value) ? value : ""; }
}
