using System;
using System.Collections.Generic;
using System.Reflection;
using OpenSlideCharts;
using System.Windows.Forms;

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
            var styled = ChartStorage.Clone(original); styled.Kind = "Column"; styled.LayoutMode = "Stacked";
            styled.Appearance = new ChartAppearance { ShowLegend = false, ShowTotals = true, SeriesLabelsAtFirstStack = true,
                Minimum = -3, Maximum = 10, SeriesColors = new List<string> { "#123456", "#ABCDEF" }, DisplayOrder = new List<int> { 1, 0 },
                PlotRect = new ChartRectangle { X = .2, Y = .1, Width = .7, Height = .8 } };
            styled.Appearance.ValueLabelOffsets.Add(new ChartLabelOffset { SeriesIndex = 1, CategoryIndex = 0, X = 2, Y = -1 });
            using (var editor = new SeriesEditor(styled))
            {
                var readMethod = typeof(SeriesEditor).GetMethod("ReadMatrix", BindingFlags.Instance | BindingFlags.NonPublic);
                var grid = (DataGridView)typeof(SeriesEditor).GetField("grid", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
                var read = (ChartSpec)readMethod.Invoke(editor, null);
                Check(read.Appearance.Maximum == 10 && read.Appearance.SeriesColors[1] == "#ABCDEF", "Unchanged appearance must survive editing.");
                grid.Rows[0].Cells[1].Value = "20";
                read = (ChartSpec)readMethod.Invoke(editor, null);
                Check(!read.Appearance.Maximum.HasValue && !read.Appearance.Minimum.HasValue, "Numeric edits must release snapshot axis limits.");
                grid.Columns.RemoveAt(1);
                read = (ChartSpec)readMethod.Invoke(editor, null);
                Check(read.Appearance.DisplayOrder.Count == 1 && read.Appearance.DisplayOrder[0] == 0 && read.Appearance.SeriesColors[0] == "#ABCDEF", "Removing a series must remap colors and display order.");
                Check(read.Appearance.ValueLabelOffsets.Count == 1 && read.Appearance.ValueLabelOffsets[0].SeriesIndex == 0 && read.Appearance.ValueLabelOffsets[0].X == 2, "Removing a series must remap label offsets.");
                var kind = (ComboBox)typeof(SeriesEditor).GetField("kind", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
                kind.SelectedItem = "Bar";
                read = (ChartSpec)readMethod.Invoke(editor, null);
                Check(read.Appearance.PlotRect == null && read.Appearance.ShowLegend && !read.Appearance.SeriesLabelsAtFirstStack, "Changing topology must rebuild plot and labels.");
                Check(read.Appearance.ValueLabelOffsets.Count == 0, "Changing topology must discard observed label offsets.");
            }
            styled.Appearance.SourceFrame = new ChartRectangle { Width = 220, Height = 104 };
            styled.Appearance.ShowTotals = false; styled.Appearance.SeriesLabelsAtFirstStack = false;
            styled.Appearance.ShowValueAxis = false; styled.Appearance.ShowGridlines = false; styled.Appearance.ShowTitle = false;
            styled.Appearance.CategoryFontSize = 8; styled.Appearance.CategoryLabelHeight = 8;
            styled.Appearance.PlotRect = new ChartRectangle { X = .2, Y = .15, Width = .7, Height = .65 };
            using (var editor = new SeriesEditor(styled))
            {
                var validate = typeof(SeriesEditor).GetMethod("ValidateCandidate", BindingFlags.Instance | BindingFlags.NonPublic);
                validate.Invoke(editor, new object[] { styled });
                typeof(SeriesEditor).GetMethod("PasteMatrix", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(editor, new object[] { "Category\tNorth\tSouth\nA\t1.2340\t\nB\t0\t-2.5" });
                var pasted = (ChartSpec)typeof(SeriesEditor).GetMethod("ReadMatrix", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(editor, null);
                validate.Invoke(editor, new object[] { pasted });
                Check(pasted.Appearance.PlotRect != null && pasted.Rows[0].Values[0] == 1.234, "Pasting into a small imported chart must retain its supported appearance.");
                var changed = ChartStorage.Clone(styled); changed.Kind = "Bar"; changed.Appearance.PlotRect = null;
                bool rejected = false;
                try { validate.Invoke(editor, new object[] { changed }); }
                catch (TargetInvocationException e) { rejected = e.InnerException is ArgumentException; }
                Check(rejected, "Small-frame type change must fail before closing the editor.");
                using (var enlarged = new SeriesEditor(styled, 720, 400))
                {
                    validate.Invoke(enlarged, new object[] { changed });
                    Check(true, "The actual enlarged group dimensions must override its original source frame.");
                }
            }
            Cascade();
            Console.WriteLine("PASS: " + checks + " editor data assertions (no interactive UI coverage)"); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
    private static void Cascade()
    {
        var spec = WaterfallProfiles.Sample(); string original = ChartStorage.Serialize(spec);
        using (var editor = new WaterfallEditor(spec))
        {
            var readMethod = typeof(WaterfallEditor).GetMethod("ReadSpec", BindingFlags.Instance | BindingFlags.NonPublic);
            var data = (DataGridView)typeof(WaterfallEditor).GetField("data", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
            var read = (ChartSpec)readMethod.Invoke(editor, null);
            Check(ChartStorage.Serialize(read) == original, "Waterfall editor must preserve unchanged data, IDs and constraints.");
            data.Rows[0].Cells[2].Value = "40";
            read = (ChartSpec)readMethod.Invoke(editor, null);
            var solved = WaterfallSolver.Solve(read.Waterfall);
            Check(solved.Columns[3].Output == 111 && solved.Columns[5].Output == 106, "Numeric editing must recalculate all dependent e values.");
            Check(read.Waterfall.Connectors[0].Id == spec.Waterfall.Connectors[0].Id, "Numeric editing must retain connector identities.");
            Check(ChartStorage.Serialize(spec) == original, "Waterfall editing before Apply must not mutate the source.");
            string matrix = "Segment\tStart\tGrowth\tCosts\tSubtotal\tOther\tTotal\nA\t100\t35\t-12\te\t-5\te\nB\t\t\t-17\t\t\t";
            var paste = typeof(WaterfallEditor).GetMethod("PasteMatrix", BindingFlags.Instance | BindingFlags.NonPublic);
            paste.Invoke(editor, new object[] { matrix });
            read = (ChartSpec)readMethod.Invoke(editor, null); solved = WaterfallSolver.Solve(read.Waterfall);
            Check(solved.Columns[3].Output == 106 && solved.Columns[5].Output == 101, "Matrix paste must retain blanks and solve e cells.");
            Check(read.Waterfall.Columns[2].Segments[1].Value == -17 && read.Waterfall.Columns[2].Segments[1].Id == spec.Waterfall.Columns[2].Segments[1].Id, "Paste must preserve stable segment identities.");
            string before = ChartStorage.Serialize(read); bool failed = false;
            try { paste.Invoke(editor, new object[] { "Segment\tA\nOne\tbad-number" }); } catch (TargetInvocationException) { failed = true; }
            Check(failed && ChartStorage.Serialize((ChartSpec)readMethod.Invoke(editor, null)) == before, "Invalid pasted numbers must leave the grid intact.");
            var signs = (CheckBox)typeof(WaterfallEditor).GetField("signs", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
            var order = (ComboBox)typeof(WaterfallEditor).GetField("order", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
            signs.Checked = true; order.SelectedIndex = 1;
            read = (ChartSpec)readMethod.Invoke(editor, null);
            Check(read.Waterfall.ShowSigns && read.Waterfall.SegmentOrder == "SheetBottomToTop", "Waterfall number signs and sheet order must persist.");
            data.Rows[0].Cells[1].Value = "e"; data.Rows[0].Cells[4].Value = "15"; data.Rows[0].Cells[6].Value = "4";
            var profile = (ComboBox)typeof(WaterfallEditor).GetField("profile", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
            profile.SelectedIndex = 3;
            typeof(WaterfallEditor).GetMethod("Arrange", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(editor, null);
            read = (ChartSpec)readMethod.Invoke(editor, null);
            Check(read.Waterfall.Pins.TrueForAll(delegate(WaterfallPin p) { return p.AutoGenerated; }), "Recreating a profile must retain automatic pin provenance.");
            var pins = (DataGridView)typeof(WaterfallEditor).GetField("pins", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(editor);
            pins.Rows[0].Cells[2].Value = "1";
            read = (ChartSpec)readMethod.Invoke(editor, null);
            Check(!read.Waterfall.Pins[0].AutoGenerated, "Editing a pin value must protect it as a manual constraint.");
        }
    }
}
