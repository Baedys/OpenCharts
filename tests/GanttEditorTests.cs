using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using OpenSlideCharts;

class GanttEditorTests
{
    private static int assertions;
    private static void Check(bool value, string message) { assertions++; if (!value) throw new Exception(message); }
    private static void Reject(Action action, string message) { bool rejected = false; try { action(); } catch (ArgumentException) { rejected = true; } catch (FormatException) { rejected = true; } catch (NotSupportedException) { rejected = true; } Check(rejected, message); }
    private static T Field<T>(object value, string name) { return (T)value.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(value); }
    private static ChartSpec Read(GanttEditor editor) { return (ChartSpec)typeof(GanttEditor).GetMethod("ReadSpec", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(editor, null); }
    internal static ChartSpec Sample()
    {
        var g = new GanttSpec { StartDate = "2024-12-30", EndDate = "2025-01-12", ScaleMode = "Manual", ShowResponsible = true, ActivityHeader = "Projet", ResponsibleHeader = "Équipe" };
        g.Scales.Add(new GanttScale { Unit = "Weeks", Format = "ww" }); g.Scales.Add(new GanttScale { Unit = "Days", Format = "dd" });
        g.Activities.Add(new GanttActivity { Id = "a", Label = "Étude", Responsible = "Produit", Remark = "Texte conservé", Indent = 1, Color = "#F0F2F5", Separator = true });
        g.Items.Add(new GanttItem { Id = "bar", RowId = "a", StartDate = "2024-12-30", EndDate = "2025-01-03", Label = "Travail", Color = "#123456", ShowDates = true });
        g.Items.Add(new GanttItem { Id = "milestone", RowId = g.MilestoneRow.Id, Kind = "Milestone", StartDate = "2025-01-03", Label = "Version" });
        g.Links.Add(new GanttLink { FromId = "bar", ToId = "milestone" });
        g.Shades.Add(new GanttShade { StartDate = "2025-01-01", EndDate = "2025-01-01", Label = "Férié", Color = "#F1E7CE" });
        return new ChartSpec { Kind = "Gantt", Title = "  Calendrier  ", Gantt = g };
    }
    [STAThread] static int Main()
    {
        try { Dates(); Import(); Editing(); Fit(); Console.WriteLine("PASS: " + assertions + " Gantt editor/data assertions."); return 0; }
        catch (Exception error) { Console.WriteLine(error); return 1; }
    }
    private static void Dates()
    {
        Check(GanttData.ReadDate(1d, false) == new DateTime(1900, 1, 1), "1900 serial 1");
        Check(GanttData.ReadDate(59d, false) == new DateTime(1900, 2, 28), "1900 serial 59");
        Check(GanttData.ReadDate(61.5, false) == new DateTime(1900, 3, 1, 12, 0, 0), "1900 after fictitious leap day + noon");
        Check(GanttData.ReadDate(0d, true) == new DateTime(1904, 1, 1), "1904 serial zero");
        Check(GanttData.ReadDate(59.5, true) == new DateTime(1904, 2, 29, 12, 0, 0), "1904 actual leap day");
        Check(GanttData.ReadDate("2024-02-29T23:59:00", false) == new DateTime(2024, 2, 29, 23, 59, 0), "ISO exact minute retained");
        Check(GanttDates.ParseInstant(GanttDates.Iso(new DateTime(2024, 2, 29, 12, 34, 56).AddTicks(1234))).Ticks == new DateTime(2024, 2, 29, 12, 34, 56).AddTicks(1234).Ticks, "ISO exact tick roundtrip");
        Reject(delegate { GanttData.ReadDate(60.5, false); }, "Excel fictional day refused");
        Reject(delegate { GanttData.ReadDate(0d, false); }, "Excel 1900 zero refused");
        Reject(delegate { GanttData.ReadDate(Double.NaN, false); }, "NaN refused");
        Reject(delegate { GanttData.ReadDate(true, false); }, "Boolean refused");
        Reject(delegate { GanttData.ReadDate("29/02/2024", false); }, "Ambiguous locale date refused");
    }
    private static void Import()
    {
        ChartSpec source = Sample(); string before = ChartStorage.Serialize(source);
        object[,] table = new object[,] { { "Jalons importés", "2025-01-03T12:00:00", null }, { "Étude importée", "2024-12-30T12:00:00", "2025-01-04" } };
        ChartSpec first = GanttData.ReadCells(source, table, false);
        Check(ChartStorage.Serialize(source) == before, "Import does not mutate source");
        Check(first.Gantt.Activities[0].Id == "a" && first.Gantt.Activities[0].Label == "Étude importée", "Position updates label, preserves row ID");
        Check(first.Gantt.MilestoneRow.Label == "Jalons importés" && first.Gantt.DateAnchors.Count == 3, "First row is milestone anchors, not activity");
        Check(first.Gantt.Items.Count == 2 && first.Gantt.Items[0].EndDate == "2025-01-03", "No implicit item changes");
        Check(first.Gantt.Links.Count == 1 && first.Gantt.Shades.Count == 1 && first.Gantt.Activities[0].Responsible == "Produit", "Options and decorations retained");
        GanttDateAnchor startAnchor = first.Gantt.DateAnchors.Find(delegate(GanttDateAnchor a) { return a.RowId == "a" && a.Index == 0; });
        GanttDateAnchor endAnchor = first.Gantt.DateAnchors.Find(delegate(GanttDateAnchor a) { return a.RowId == "a" && a.Index == 1; });
        first.Gantt.Items[0].StartAnchor = startAnchor.Id; first.Gantt.Items[0].EndAnchor = endAnchor.Id; first.Gantt.Items[0].DateMode = "Exact";
        table[1, 1] = "2024-12-31T13:00:00"; string linkedBefore = ChartStorage.Serialize(first);
        ChartSpec refreshed = GanttData.ReadCells(first, table, false);
        Check(refreshed.Gantt.DateAnchors.Find(delegate(GanttDateAnchor a) { return a.Id == startAnchor.Id; }).Date == "2024-12-31T13:00:00", "Refresh retains anchor identity and exact hour");
        Check(refreshed.Gantt.Items[0].StartAnchor == startAnchor.Id && refreshed.Gantt.Items[0].DateMode == "Exact", "Item remains linked");
        Check(ChartStorage.Serialize(first) == linkedBefore, "Refresh source immutable");
        table[1, 1] = null; Reject(delegate { GanttData.ReadCells(first, table, false); }, "Missing referenced date rejects transaction");
        Reject(delegate { GanttData.ReadCells(first, new object[3, 2], false); }, "Changed activity structure with objects refused");
        Check(ChartStorage.Serialize(first) == linkedBefore, "Rejected imports retain source");
        var empty = new ChartSpec { Kind = "Gantt", Gantt = new GanttSpec { StartDate = "2024-01-01", EndDate = "2025-12-31" } };
        ChartSpec created = GanttData.ReadTsv(empty, "Jalons\t2025-01-03T12:00:00\nÉtude\t2024-12-30\t2025-01-04\nProduction\t2025-01-06");
        Check(created.Gantt.Activities.Count == 2 && created.Gantt.Items.Count == 0 && created.Gantt.DateAnchors.Count == 4, "Excel table creates anchors/labels only");
        Check(empty.Gantt.Activities.Count == 0 && empty.Gantt.DateAnchors.Count == 0, "First import isolates source");
        var colliding = ChartStorage.Clone(empty); colliding.Gantt.Activities.Add(new GanttActivity { Id = "activity_2" });
        ChartSpec expanded = GanttData.ReadCells(colliding, new object[3, 2], false);
        Check(expanded.Gantt.Activities[0].Id == "activity_2" && expanded.Gantt.Activities[1].Id != "activity_2", "New row IDs cannot collide with a retained row");
        Array oneBased = Array.CreateInstance(typeof(object), new[] { 2, 2 }, new[] { 1, 1 });
        oneBased.SetValue("Jalons", 1, 1); oneBased.SetValue("Étude", 2, 1); oneBased.SetValue(0d, 2, 2);
        ChartSpec base1904 = GanttData.ReadCells(empty, (object[,])oneBased, true);
        Check(GanttDates.ParseInstant(base1904.Gantt.DateAnchors[0].Date) == new DateTime(1904, 1, 1), "COM one-based matrix + 1904 workbook");
    }
    private static void Editing()
    {
        ChartSpec source = Sample(); string before = ChartStorage.Serialize(source);
        using (var editor = new GanttEditor(source, 900, 470))
        {
            ChartSpec candidate = Read(editor); ChartEngine.Build(candidate, 900, 470);
            Check(candidate.Title == source.Title, "Custom title whitespace retained");
            Check(candidate.Gantt.ActivityHeader == "Projet" && candidate.Gantt.ResponsibleHeader == "Équipe", "Headings retained");
            Check(candidate.Gantt.Activities[0].Color == "#F0F2F5" && candidate.Gantt.Activities[0].Indent == 1 && candidate.Gantt.Activities[0].Separator, "Activity style retained");
            Check(candidate.Gantt.Items[0].Id == "bar" && candidate.Gantt.Items[0].Color == "#123456" && candidate.Gantt.Items[0].ShowDates, "Item identity/style retained");
            Check(candidate.Gantt.Links.Count == 1 && candidate.Gantt.Shades[0].Label == "Férié", "Annotations retained");
            Check(!Field<DataGridView>(editor, "items").Columns[0].Visible && Field<DataGridView>(editor, "items").Columns[1] is DataGridViewComboBoxColumn, "User chooses activities instead of typing IDs");
            Field<ComboBox>(editor, "weekend").SelectedIndex = 2;
            candidate = Read(editor); Check(candidate.Gantt.WeekendDisplay == "Hidden", "Weekend option read");
            CheckBox[] days = Field<CheckBox[]>(editor, "weekendChecks"); days[5].Checked = true; days[0].Checked = false;
            candidate = Read(editor); Check(candidate.Gantt.WeekendDays.Count == 2 && candidate.Gantt.WeekendDays.Contains(5) && candidate.Gantt.WeekendDays.Contains(6), "Explicit weekend checkboxes retain weekday choices");
            Check(ChartStorage.Serialize(source) == before && ChartStorage.Serialize(editor.Spec) == before, "Candidate reading leaves source and committed editor model unchanged");
            Field<TextBox>(editor, "start").Text = "bad";
            Reject(delegate { ChartEngine.Build(Read(editor), 900, 470); }, "Invalid range rejected before committing model");
            Check(ChartStorage.Serialize(editor.Spec) == before, "Invalid candidate preserves committed model");
        }
        ChartSpec withAnchor = GanttData.ReadCells(source, new object[,] { { "Jalons", null, null }, { "Étude", "2024-12-30T12:00:00", "2025-01-04" } }, false);
        withAnchor.Gantt.Items[0].DateMode = "Exact"; withAnchor.Gantt.Items[0].StartAnchor = withAnchor.Gantt.DateAnchors[0].Id; withAnchor.Gantt.Items[0].EndAnchor = withAnchor.Gantt.DateAnchors[1].Id;
        using (var editor = new GanttEditor(withAnchor, 900, 470))
        {
            ChartSpec candidate = Read(editor); ChartEngine.Build(candidate, 900, 470);
            Check(candidate.Gantt.DateAnchors[0].Index == 0 && candidate.Gantt.DateAnchors[1].Index == 1, "User date column numbers roundtrip to zero-based anchor indices");
            Check(candidate.Gantt.Items[0].StartAnchor == withAnchor.Gantt.Items[0].StartAnchor && candidate.Gantt.Items[0].EndAnchor == withAnchor.Gantt.Items[0].EndAnchor, "Anchor choices retained");
        }
        ChartSpec nullable = Sample(); nullable.Gantt.ScaleMode = "Auto"; nullable.Gantt.Scales = null; nullable.Gantt.DateAnchors = null; nullable.Gantt.Shades = null; nullable.Gantt.Links = null;
        string nullableBefore = ChartStorage.Serialize(nullable);
        using (var editor = new GanttEditor(nullable, 900, 470)) { ChartSpec candidate = Read(editor); ChartEngine.Build(candidate, 900, 470); Check(candidate.Gantt.Items.Count == 2, "Valid null optional lists open and edit"); }
        Check(GanttData.FitScale(nullable).Gantt.EndDate == "2025-01-03", "Fit accepts null date anchors");
        Check(GanttData.ReadCells(nullable, new object[2, 2], false).Gantt.Items.Count == 2, "Import accepts null date anchors");
        Check(ChartStorage.Serialize(nullable) == nullableBefore, "Null list normalization leaves source unchanged");
        ChartSpec reversed = Sample(); reversed.Gantt.Scales = new List<GanttScale> { new GanttScale { Unit = "Weeks" }, new GanttScale { Unit = "Months" } };
        using (var editor = new GanttEditor(reversed, 900, 470)) { ChartSpec candidate = Read(editor); Check(candidate.Gantt.Scales[0].Unit == "Weeks" && candidate.Gantt.Scales[1].Unit == "Months", "Editing preserves manual calendar band order"); }
        var empty = new ChartSpec { Kind = "Gantt", Gantt = new GanttSpec { StartDate = "2025-01-01", EndDate = "2025-02-01" } };
        using (var editor = new GanttEditor(empty, 900, 470))
        {
            DataGridView grid = Field<DataGridView>(editor, "anchors"); int errors = 0;
            grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { errors++; e.ThrowException = false; };
            ChartSpec imported = GanttData.ReadTsv(empty, "Jalons\t2025-01-03\nActivité\t2025-01-04");
            typeof(GanttEditor).GetMethod("LoadSpec", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(editor, new object[] { imported });
            Check(Convert.ToString(grid.Rows[1].Cells[1].FormattedValue) == "Activité" && errors == 0, "Imported activity choices refreshed before displaying anchors");
            Check(Read(editor).Gantt.DateAnchors[1].RowId == imported.Gantt.Activities[0].Id, "Imported new identities survive readback");
            imported.Gantt.DateAnchors.RemoveAt(0);
            typeof(GanttEditor).GetMethod("LoadSpec", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(editor, new object[] { imported });
            Check(Convert.ToString(grid.Rows[0].Cells[1].FormattedValue) == "Activité" && errors == 0, "Reload after deleting anchors remains valid");
        }
    }
    private static void Fit()
    {
        ChartSpec source = Sample(); string before = ChartStorage.Serialize(source);
        ChartSpec fitted = GanttData.FitScale(source);
        Check(fitted.Gantt.StartDate == "2024-12-30" && fitted.Gantt.EndDate == "2025-01-03", "Fit all items inclusive dates");
        Check(ChartStorage.Serialize(source) == before, "Fit source isolated");
        source.Gantt.Items.RemoveAt(1); source.Gantt.Links.Clear(); source.Gantt.Items[0].DateMode = "Exact"; source.Gantt.Items[0].StartDate = "2025-01-01T12:00:00"; source.Gantt.Items[0].EndDate = "2025-01-02T00:00:00";
        fitted = GanttData.FitScale(source); Check(fitted.Gantt.StartDate == "2025-01-01" && fitted.Gantt.EndDate == "2025-01-01", "Exact midnight end excludes ending day");
        source.Gantt.Items.Clear(); Reject(delegate { GanttData.FitScale(source); }, "No temporal item fit refused");
    }
}
