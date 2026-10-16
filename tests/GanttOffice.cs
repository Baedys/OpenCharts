using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using OpenSlideCharts;

// Synthetic renderer integration, not a claim of reference-product parity.
// No Office PIA, Excel process, add-in registration or COMAddIns mutation.
public static class GanttOffice
{
    private static int assertions;
    private sealed class Fixture
    {
        internal string Name;
        internal ChartSpec Spec;
        internal ChartRectangle Frame = new ChartRectangle { X = 30, Y = 45, Width = 900, Height = 450 };
        internal List<string> Names = new List<string>();
    }
    private static void Check(bool value, string message) { assertions++; if (!value) throw new Exception(message); }
    private static void Near(double a, double b, string message) { Check(Math.Abs(a - b) < 1e-7, message + ": " + a + " vs " + b); }
    private static string N(double value) { return value.ToString("R", CultureInfo.InvariantCulture); }
    private static ChartSpec Spec(string first = "2024-02-26", string last = "2024-03-04")
    {
        ChartSpec spec = new ChartSpec { Kind = "Gantt", Title = "Synthetic dated Gantt", Gantt = new GanttSpec { StartDate = first, EndDate = last, WeekendDisplay = "Plain", ScaleMode = "Manual" } };
        spec.Gantt.Activities.Add(new GanttActivity { Id = "A", Label = "Activity A" });
        spec.Gantt.Scales.Add(new GanttScale { Unit = "Days", Format = "dd" }); return spec;
    }
    private static GanttItem Add(ChartSpec spec, string id, string first, string last, string kind = "Bar", string row = "A", string label = null)
    {
        GanttItem item = new GanttItem { Id = id, RowId = row, StartDate = first, EndDate = last, Kind = kind, Label = label ?? id, Color = "#367F9C" };
        spec.Gantt.Items.Add(item); return item;
    }
    private static List<Fixture> Fixtures()
    {
        List<Fixture> list = new List<Fixture>();
        ChartSpec scales = Spec("2023-12-15", "2024-04-15"); scales.Gantt.Scales.Clear();
        scales.Gantt.Scales.Add(new GanttScale { Unit = "Years", Format = "yyyy" });
        scales.Gantt.Scales.Add(new GanttScale { Unit = "Quarters", Format = "'Q'q" });
        scales.Gantt.Scales.Add(new GanttScale { Unit = "Months", Format = "mmm", Separators = false });
        Add(scales, "phase", "2024-01-10", "2024-03-20", "Chevron", "A", "Quarter rollover");
        list.Add(new Fixture { Name = "01-years-quarters-months", Spec = scales });
        ChartSpec leap = Spec(); Add(leap, "leap", "2024-02-28", "2024-03-01").ShowDates = true;
        list.Add(new Fixture { Name = "02-leap-day", Spec = leap });
        ChartSpec iso = Spec("2020-12-28", "2021-01-10"); iso.Gantt.Scales.Insert(0, new GanttScale { Unit = "Weeks", Format = "www" });
        Add(iso, "rollover", "2020-12-31", "2021-01-04"); list.Add(new Fixture { Name = "03-iso-week-year-boundary", Spec = iso });
        ChartSpec shaded = Spec(); shaded.Gantt.WeekendDisplay = "Shaded";
        shaded.Gantt.Shades.Add(new GanttShade { StartDate = "2024-02-29", EndDate = "2024-03-01", Label = "Synthetic holiday", Color = "#EFE2B9" });
        Add(shaded, "work", "2024-02-27", "2024-03-04"); list.Add(new Fixture { Name = "04-weekends-and-holiday-shades", Spec = shaded });
        ChartSpec hidden = Spec("2024-03-01", "2024-03-04"); hidden.Gantt.WeekendDisplay = "Hidden";
        Add(hidden, "friday", "2024-03-01", "2024-03-01");
        Add(hidden, "weekend", "2024-03-02", "2024-03-03", "Bar", "A", "Hidden weekend label retained");
        Add(hidden, "hiddenMilestone", "2024-03-02", null, "Milestone", "__milestones", "Weekend milestone label");
        list.Add(new Fixture { Name = "05-hidden-weekend-labels", Spec = hidden });
        ChartSpec custom = Spec("2024-01-01", "2024-01-12"); custom.Gantt.WeekendDisplay = "Hidden"; custom.Gantt.WeekendDays = new List<int> { 1, 4 }; custom.Gantt.WeekStartsOn = 0;
        custom.Gantt.Scales.Insert(0, new GanttScale { Unit = "Weeks", Format = "ww" }); Add(custom, "custom", "2024-01-02", "2024-01-09");
        list.Add(new Fixture { Name = "06-custom-hidden-weekdays", Spec = custom });
        ChartSpec types = Spec();
        Add(types, "bar", "2024-02-27", "2024-03-01", "Bar"); Add(types, "pentagon", "2024-02-28", "2024-03-02", "Pentagon");
        Add(types, "chevron", "2024-02-29", "2024-03-03", "Chevron");
        Add(types, "milestone", "2024-03-01", null, "Milestone", "__milestones");
        Add(types, "bottomBracket", "2024-02-27", "2024-03-03", "Bracket", "__milestones");
        list.Add(new Fixture { Name = "07-five-types-overlapping-lanes", Spec = types });
        ChartSpec clipping = Spec(); Add(clipping, "clipped", "2024-02-01", "2024-03-31", "Bar", "A", "Clipped interval");
        Add(clipping, "outside", "2025-01-01", "2025-01-10", "Bar", "A", "Restore original outside label");
        list.Add(new Fixture { Name = "08-clipped-and-outside-restoration", Spec = clipping });
        ChartSpec anchorsOnly = Spec("2024-01-01", "2024-01-07");
        list.Add(new Fixture { Name = "09-empty-then-date-anchors", Spec = anchorsOnly });
        ChartSpec modes = Spec("2024-01-01", "2024-01-02"); modes.Gantt.Activities.Add(new GanttActivity { Id = "B", Label = "Inclusive days" });
        GanttItem exact = Add(modes, "exact", "2024-01-01T00:00:00", "2024-01-02T00:00:00", "Bar", "A", "Exact end at midnight"); exact.DateMode = "Exact"; exact.ShowDates = true;
        Add(modes, "inclusive", "2024-01-01", "2024-01-02", "Bar", "B", "Inclusive final day").ShowDates = true;
        GanttItem noon = Add(modes, "noon", "2024-01-01T12:00:00", null, "Milestone", "__milestones", "Noon"); noon.DateMode = "Exact";
        list.Add(new Fixture { Name = "10-exact-midnight-versus-inclusive", Spec = modes });
        ChartSpec anchored = Spec("2024-01-01", "2024-01-03");
        anchored.Gantt.DateAnchors.Add(new GanttDateAnchor { Id = "start", RowId = "A", Index = 0, Date = "2024-01-01T06:00:00" });
        anchored.Gantt.DateAnchors.Add(new GanttDateAnchor { Id = "end", RowId = "A", Index = 1, Date = "2024-01-02T18:00:00" });
        GanttItem anchoredBar = Add(anchored, "anchored", "2024-01-01", "2024-01-03", "Bar", "A", "Anchored refresh");
        anchoredBar.DateMode = "Exact"; anchoredBar.StartAnchor = "start"; anchoredBar.EndAnchor = "end"; anchoredBar.ShowDates = true;
        list.Add(new Fixture { Name = "11-anchored-refresh-midday", Spec = anchored });
        ChartSpec links = Spec("2024-01-01", "2024-01-21"); links.Gantt.Activities.Add(new GanttActivity { Id = "B", Label = "Activity B" });
        Add(links, "from", "2024-01-02", "2024-01-04", "Pentagon"); Add(links, "to", "2024-01-08", "2024-01-10", "Chevron", "B");
        Add(links, "sameRow", "2024-01-15", "2024-01-17");
        links.Gantt.Links.Add(new GanttLink { FromId = "from", ToId = "to", Color = "#A84B48" });
        links.Gantt.Links.Add(new GanttLink { FromId = "from", ToId = "sameRow", Color = "#568B50" });
        list.Add(new Fixture { Name = "12-visual-links", Spec = links });
        ChartSpec text = Spec(); text.Gantt.ShowResponsible = true; text.Gantt.ShowRemark = true; text.Gantt.LabelFraction = .42;
        text.Gantt.Activities[0].Label = new String('W', 70) + " 東京市場分析"; text.Gantt.Activities[0].Responsible = "Alice\nBob";
        text.Gantt.Activities[0].Remark = "販売計画比較 " + new String('W', 40); text.Gantt.Activities[0].Indent = 2;
        text.Gantt.Activities[0].Color = "#EBEFF8"; text.Gantt.Activities[0].Separator = true;
        Add(text, "long", "2024-02-27", "2024-03-01", "Chevron", "A", new String('W', 80) + " 東京計画").ShowDates = true;
        list.Add(new Fixture { Name = "13-three-columns-long-cjk", Spec = text });
        ChartSpec minimum = Spec(); minimum.Title = "Small frame"; minimum.Gantt.FontSize = 8;
        Add(minimum, "small", "2024-02-27", "2024-02-28", "Bar", "A", "Short");
        list.Add(new Fixture { Name = "14-minimum-frame", Spec = minimum, Frame = new ChartRectangle { X = 320, Y = 150, Width = 320, Height = 240 } });
        ChartSpec first = Spec("0001-01-01", "0001-01-01"); Add(first, "first", "0001-01-01", "0001-01-01").ShowDates = true;
        list.Add(new Fixture { Name = "15-first-representable-day", Spec = first });
        ChartSpec last = Spec("9999-12-31", "9999-12-31"); Add(last, "last", "9999-12-31", "9999-12-31").ShowDates = true;
        list.Add(new Fixture { Name = "16-last-representable-day", Spec = last });
        return list;
    }
    private static List<Primitive> Role(Layout layout, string role, string id = null)
    { List<Primitive> found = new List<Primitive>(); foreach (Primitive p in layout.Items) if (p.Role == role && (id == null || p.EntityId == id)) found.Add(p); return found; }
    private static Primitive One(Layout layout, string role, string id)
    { List<Primitive> found = Role(layout, role, id); Check(found.Count == 1, "one " + role + " for " + id); return found[0]; }
    private static bool IsProcessLabel(ChartSpec spec, Primitive p)
    {
        if (p.Role != "GANTT_ITEM_LABEL") return false;
        foreach (GanttItem item in spec.Gantt.Items)
            if (item.Id == p.EntityId) return item.Kind == "Pentagon" || item.Kind == "Chevron";
        return false;
    }
    private static Layout CheckLayout(Fixture fixture)
    {
        Layout layout = ChartEngine.Build(fixture.Spec, fixture.Frame.Width, fixture.Frame.Height);
        GanttSpec g = fixture.Spec.Gantt; Primitive edge = Role(layout, "GANTT_SCALE_EDGE")[0];
        foreach (Primitive p in layout.Items)
        {
            Check(p.X >= -1e-7 && p.Y >= -1e-7 && p.X + p.W >= -1e-7 && p.Y + p.H >= -1e-7 && p.X <= fixture.Frame.Width + 1e-7 && p.X + p.W <= fixture.Frame.Width + 1e-7 && p.Y <= fixture.Frame.Height + 1e-7 && p.Y + p.H <= fixture.Frame.Height + 1e-7, "primitive remains inside chart frame");
            if (p.Role == "GANTT_ITEM_LABEL" && Role(layout, "GANTT_ITEM", p.EntityId).Count != 0)
                Check(layout.Items.IndexOf(p) > layout.Items.IndexOf(Role(layout, "GANTT_ITEM", p.EntityId)[0]), "item label follows its native geometry in z-order");
            if (IsProcessLabel(fixture.Spec, p) && Role(layout, "GANTT_ITEM", p.EntityId).Count != 0)
            {
                Primitive process = One(layout, "GANTT_ITEM", p.EntityId);
                Check(p.X >= process.X && p.Y >= process.Y && p.X + p.W <= process.X + process.W && p.Y + p.H <= process.Y + process.H, "process label layout is inside its arrow");
            }
        }
        switch (fixture.Name)
        {
            case "01-years-quarters-months": Check(Role(layout, "GANTT_SCALE_EDGE").Count == 3, "year/quarter/month scales present"); break;
            case "02-leap-day": Near(One(layout, "GANTT_ITEM", "leap").W, edge.W * 3 / 8, "inclusive leap interval covers three days"); break;
            case "03-iso-week-year-boundary":
                bool week53 = false, week1 = false;
                foreach (Primitive p in Role(layout, "GANTT_SCALE_LABEL")) if (p.EntityId.StartsWith("Weeks:", StringComparison.Ordinal)) { week53 |= p.Text.Contains("53"); week1 |= p.Text.EndsWith("1", StringComparison.Ordinal); }
                Check(week53 && week1, "ISO week 53 to week 1 rollover labels"); break;
            case "04-weekends-and-holiday-shades": Check(Role(layout, "GANTT_WEEKEND").Count == 2 && Role(layout, "GANTT_SHADE").Count == 1, "weekend and custom shade bands independent"); break;
            case "05-hidden-weekend-labels":
                Near(One(layout, "GANTT_ITEM", "friday").W, edge.W / 2, "hidden weekends compress four days to two");
                Check(Role(layout, "GANTT_ITEM", "weekend").Count == 0 && Role(layout, "GANTT_ITEM", "hiddenMilestone").Count == 0, "hidden-only geometry omitted");
                Check(Role(layout, "GANTT_ITEM_LABEL", "weekend").Count == 1 && Role(layout, "GANTT_ITEM_LABEL", "hiddenMilestone").Count == 1, "hidden-only labels retained"); break;
            case "06-custom-hidden-weekdays": Check(new GanttCalendar(g).TotalUnits == 8 && Role(layout, "GANTT_WEEKEND").Count == 0, "custom Monday/Thursday weekdays compressed"); break;
            case "07-five-types-overlapping-lanes":
                Check(Role(layout, "GANTT_ITEM", "bottomBracket").Count == 3, "bottom bracket uses three native lines");
                Check(One(layout, "GANTT_ITEM", "bar").Y != One(layout, "GANTT_ITEM", "pentagon").Y && One(layout, "GANTT_ITEM", "pentagon").Y != One(layout, "GANTT_ITEM", "chevron").Y, "overlapping items use separate lanes");
                Check(One(layout, "GANTT_ACTIVITY", "__milestones").Y > One(layout, "GANTT_ACTIVITY", "A").Y, "fixed milestone row remains below activities"); break;
            case "08-clipped-and-outside-restoration":
                if (g.StartDate == "2024-02-26") { Near(One(layout, "GANTT_ITEM", "clipped").W, edge.W, "clipped bar spans chart"); Check(Role(layout, "GANTT_OVERFLOW_START").Count == 1 && Role(layout, "GANTT_OVERFLOW_END").Count == 1, "both overflow caps present"); Check(Role(layout, "GANTT_ITEM", "outside").Count == 0, "outside geometry hidden"); }
                else Check(One(layout, "GANTT_ITEM_LABEL", "outside").Text == "Restore original outside label", "range restoration recovers original label");
                Check(g.Items[1].Label == "Restore original outside label", "outside model label retained"); break;
            case "09-empty-then-date-anchors": Check(g.Items.Count == 0 && Role(layout, "GANTT_ITEM").Count == 0, "date-anchor imports never invent timeline objects"); break;
            case "10-exact-midnight-versus-inclusive":
                Near(One(layout, "GANTT_ITEM", "exact").W * 2, One(layout, "GANTT_ITEM", "inclusive").W, "Exact midnight end excludes next day");
                Primitive noon = One(layout, "GANTT_ITEM", "noon"); Near(noon.X + noon.W / 2, edge.X + edge.W / 4, "exact noon milestone retains intraday position"); break;
            case "11-anchored-refresh-midday":
                GanttCalendar calendar = new GanttCalendar(g); Primitive bar = One(layout, "GANTT_ITEM", "anchored");
                Near(bar.X, edge.X + edge.W * calendar.Fraction(GanttDates.ParseInstant(g.DateAnchors[0].Date)), "bound anchor drives exact start");
                Near(bar.W, edge.W * (calendar.Fraction(GanttDates.ParseInstant(g.DateAnchors[1].Date)) - calendar.Fraction(GanttDates.ParseInstant(g.DateAnchors[0].Date))), "bound anchors drive exact duration"); break;
            case "12-visual-links": Check(Role(layout, "GANTT_LINK_ARROW").Count == 2 && Role(layout, "GANTT_LINK").Count >= 6, "native visual links and arrows present"); break;
            case "13-three-columns-long-cjk": Check(Role(layout, "GANTT_HEADER").Count == 3 && Role(layout, "GANTT_RESPONSIBLE").Count > 0 && Role(layout, "GANTT_REMARK").Count > 0, "three label columns rendered"); break;
            case "15-first-representable-day": Near(One(layout, "GANTT_ITEM", "first").W, edge.W, "inclusive minimum date is full day"); break;
            case "16-last-representable-day": Near(One(layout, "GANTT_ITEM", "last").W, edge.W, "inclusive maximum date is full day"); break;
        }
        return layout;
    }

    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length != 1) { Console.Error.WriteLine("Usage: GanttOffice.exe OUTPUT_DIRECTORY"); return 2; }
        Process[] processes = Process.GetProcessesByName("POWERPNT"); bool existing = processes.Length > 0; foreach (Process process in processes) process.Dispose();
        if (existing) { Console.Error.WriteLine("Close PowerPoint before isolated Gantt integration tests."); return 2; }
        string output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output); string deckPath = Path.Combine(output, "OpenSlideCharts-gantt.pptx");
        List<Fixture> fixtures = Fixtures(); dynamic ppt = null, deck = null;
        StringBuilder measurements = new StringBuilder("fixture\tstage\titem\trole\tentity\twidth_pt\theight_pt\ttext_bound_width_pt\ttext_bound_height_pt\n");
        StringBuilder metadata = new StringBuilder("fixture\tstage\tjson_utf8_bytes\tmetadata_tags\tprimitives\n");
        try
        {
            foreach (Fixture fixture in fixtures) CheckLayout(fixture);
            ppt = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
            if ((int)ppt.Presentations.Count != 0) throw new InvalidOperationException("Another presentation exists; test aborted.");
            ppt.Visible = -1; deck = ppt.Presentations.Add(-1); deck.PageSetup.SlideWidth = 960f; deck.PageSetup.SlideHeight = 540f;
            for (int i = 0; i < fixtures.Count; i++)
            {
                Fixture fixture = fixtures[i]; dynamic slide = null, group = null; Stopwatch elapsed = Stopwatch.StartNew();
                try
                {
                    slide = deck.Slides.Add(i + 1, 12); slide.FollowMasterBackground = 0; slide.Background.Fill.Solid(); slide.Background.Fill.ForeColor.RGB = 0xF4F1EB;
                    ChartRectangle frame = fixture.Frame; group = OfficeRenderer.Render((object)slide, fixture.Spec, frame.X, frame.Y, frame.Width, frame.Height);
                    CheckGroup((object)group, fixture, "render", measurements, metadata);
                    if (fixture.Name == "08-clipped-and-outside-restoration")
                    {
                        ChartSpec original = fixture.Spec, expanded = ChartStorage.Clone(original); expanded.Gantt.StartDate = "2025-01-01"; expanded.Gantt.EndDate = "2025-01-14";
                        object fresh = Replace((object)slide, (object)group, fixture, expanded); Release((object)group); group = fresh;
                        CheckGroup((object)group, fixture, "restored-range", measurements, metadata);
                        slide.Export(Path.Combine(output, fixture.Name + "-restored-range.png"), "PNG", 1440, 810);
                        fresh = Replace((object)slide, (object)group, fixture, original); Release((object)group); group = fresh;
                    }
                    if (fixture.Name == "09-empty-then-date-anchors" || fixture.Name == "11-anchored-refresh-midday")
                    {
                        slide.Export(Path.Combine(output, fixture.Name + "-before-refresh.png"), "PNG", 1440, 810);
                        ChartSpec source = fixture.Spec; string before = ChartStorage.Serialize(source);
                        object[,] cells = { { "Jalons", null, null }, { "Activity A refreshed", 45292.5, new DateTime(2024, 1, 3).AddTicks(1234567) } };
                        ChartSpec refreshed = GanttData.ReadCells(source, cells, false);
                        Check(ChartStorage.Serialize(source) == before, "anchor refresh leaves source model unchanged");
                        Check(refreshed.Gantt.DateAnchors.Count == 2 && refreshed.Gantt.DateAnchors[0].Date == "2024-01-01T12:00:00" && refreshed.Gantt.DateAnchors[1].Date == "2024-01-03T00:00:00.1234567", "numeric Excel noon and exact DateTime ticks retained without Excel process");
                        if (fixture.Name == "11-anchored-refresh-midday") Check(refreshed.Gantt.DateAnchors[0].Id == "start" && refreshed.Gantt.DateAnchors[1].Id == "end" && refreshed.Gantt.Items[0].StartDate == source.Gantt.Items[0].StartDate, "anchor IDs and item fallback dates survive refresh");
                        object fresh = Replace((object)slide, (object)group, fixture, refreshed); Release((object)group); group = fresh;
                    }
                    CheckGroup((object)group, fixture, "final", measurements, metadata);
                    if (fixture.Name == "11-anchored-refresh-midday") InvalidReplacements((object)slide, (object)group, fixture, measurements, metadata);
                    slide.Export(Path.Combine(output, fixture.Name + ".png"), "PNG", 1440, 810);
                    Console.WriteLine("Rendered/verified/exported " + fixture.Name + " in " + elapsed.ElapsedMilliseconds + " ms");
                }
                finally { Release((object)group); Release((object)slide); }
            }
            deck.SaveAs(deckPath, 24); deck.Close(); Release((object)deck); deck = null;
            CheckPackage(deckPath, fixtures);
            deck = ppt.Presentations.Open(deckPath, 0, 0, 0); Check((int)deck.Slides.Count == fixtures.Count, "slide count after reopen");
            for (int i = 0; i < fixtures.Count; i++)
            {
                dynamic slide = null, group = null; Stopwatch elapsed = Stopwatch.StartNew();
                try
                {
                    slide = deck.Slides.Item(i + 1); Check((int)slide.Shapes.Count == 1, "one native group after reopen"); group = slide.Shapes.Item(1);
                    CheckGroup((object)group, fixtures[i], "reopen", measurements, metadata);
                    slide.Export(Path.Combine(output, fixtures[i].Name + "-reopened.png"), "PNG", 1440, 810);
                    Console.WriteLine("Reopened/verified/exported " + fixtures[i].Name + " in " + elapsed.ElapsedMilliseconds + " ms");
                }
                finally { Release((object)group); Release((object)slide); }
            }
            Console.WriteLine("PASS: " + assertions + " Gantt Office assertions; " + fixtures.Count + " synthetic fixtures, native geometry/text, exact metadata, save/reopen and rollback. No reference-product parity claim."); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            try { File.WriteAllText(Path.Combine(output, "gantt-native-measurements.tsv"), measurements.ToString(), new UTF8Encoding(false)); File.WriteAllText(Path.Combine(output, "gantt-metadata.tsv"), metadata.ToString(), new UTF8Encoding(false)); }
            catch (Exception error) { Console.Error.WriteLine("Could not save diagnostics: " + error.Message); }
            if ((object)deck != null) { try { deck.Close(); } catch { } Release((object)deck); }
            if ((object)ppt != null) { try { if ((int)ppt.Presentations.Count == 0) ppt.Quit(); else Console.Error.WriteLine("Other presentations exist: PowerPoint left open."); } catch { } Release((object)ppt); }
        }
    }
    private static object Replace(object slideObject, object oldObject, Fixture fixture, ChartSpec replacement)
    {
        dynamic slide = slideObject, old = oldObject; int previous = (int)old.Id;
        object result = OfficeRenderer.Replace(slideObject, oldObject, replacement); dynamic fresh = result;
        Check((int)slide.Shapes.Count == 1 && (int)fresh.Id != previous, "valid replacement is transactional and preserves single chart"); fixture.Spec = replacement; return result;
    }
    private static void CheckGroup(object value, Fixture fixture, string stage, StringBuilder measurements, StringBuilder metadata)
    {
        dynamic group = value; Layout layout = CheckLayout(fixture); ChartRectangle frame = fixture.Frame;
        Check((int)group.Type == 6 && (string)group.Tags.Item("OSC_VERSION") == "1", stage + ": tagged native group");
        string expected = ChartStorage.Serialize(fixture.Spec), actual = ChartStorage.Serialize(ChartStorage.Read(value));
        Check(expected == actual, stage + ": complete canonical model retained");
        byte[] bytes = Encoding.UTF8.GetBytes(expected); int tags = Int32.Parse((string)group.Tags.Item("OSC_COUNT"), CultureInfo.InvariantCulture);
        Check(tags == (bytes.Length * 2 + 179) / 180, stage + ": metadata tag count matches all UTF8 bytes");
        StringBuilder hex = new StringBuilder();
        for (int i = 0; i < tags; i++) { string part = (string)group.Tags.Item("OSC_DATA_" + i.ToString("D4", CultureInfo.InvariantCulture)); Check(part.Length > 0 && part.Length <= 180, stage + ": metadata chunk size"); hex.Append(part); }
        Check(hex.ToString() == BitConverter.ToString(bytes).Replace("-", ""), stage + ": every metadata byte preserved");
        metadata.Append(fixture.Name).Append('\t').Append(stage).Append('\t').Append(bytes.Length).Append('\t').Append(tags).Append('\t').Append(layout.Items.Count).Append('\n');
        Check((int)group.GroupItems.Count == layout.Items.Count + 1, stage + ": native primitives plus transparent frame only");
        Check(Math.Abs((double)group.Left - frame.X) < .05 && Math.Abs((double)group.Top - frame.Y) < .05 && Math.Abs((double)group.Width - frame.Width) < .05 && Math.Abs((double)group.Height - frame.Height) < .05, stage + ": chart frame preserved");
        fixture.Names.Clear(); dynamic bounds = group.GroupItems.Item(1); fixture.Names.Add((string)bounds.Name); Release((object)bounds);
        for (int i = 0; i < layout.Items.Count; i++)
        {
            Primitive p = layout.Items[i]; dynamic child = null;
            try
            {
                child = group.GroupItems.Item(i + 2); fixture.Names.Add((string)child.Name);
                if (!String.IsNullOrEmpty(p.Role)) Check((string)child.Tags.Item("OSC_ROLE") == p.Role && (string)child.Tags.Item("OSC_ITEM") == i.ToString(CultureInfo.InvariantCulture), stage + ": role/index tags preserve layout mapping");
                double width = (double)child.Width, height = (double)child.Height;
                double expectedWidth = p.Kind == "line" ? Math.Abs(p.W) : p.Kind == "text" ? Math.Max(1, p.W) : Math.Max(.05, p.W);
                double expectedHeight = p.Kind == "line" ? Math.Abs(p.H) : p.Kind == "text" ? Math.Max(1, p.H) : Math.Max(.05, p.H);
                if (p.Kind == "path") { expectedWidth = p.W; expectedHeight = p.H; }
                double left = p.Kind == "line" ? Math.Min(p.X, p.X + p.W) : p.X, top = p.Kind == "line" ? Math.Min(p.Y, p.Y + p.H) : p.Y;
                Check(Math.Abs(width - expectedWidth) < .05 && Math.Abs(height - expectedHeight) < .05 && Math.Abs((double)child.Left - frame.X - left) < .05 && Math.Abs((double)child.Top - frame.Y - top) < .05, stage + ": native bbox matches " + p.Role + "/" + p.EntityId);
                double? bw = null, bh = null;
                if (p.Kind == "text")
                {
                    string text = Normalize((string)child.TextFrame.TextRange.Text); Check(text == Normalize(p.Text ?? ""), stage + ": editable native text");
                    Check((int)child.TextFrame.TextRange.Font.Color.RGB == p.Color, stage + ": text color");
                    bw = (double)child.TextFrame.TextRange.BoundWidth; bh = (double)child.TextFrame.TextRange.BoundHeight;
                    Check(bw <= width + .5 && bh <= height + .5, stage + ": glyph bounds " + N(bw.Value) + "x" + N(bh.Value) + " exceed " + N(width) + "x" + N(height) + " for " + p.Role + ": " + text);
                    if (IsProcessLabel(fixture.Spec, p) && Role(layout, "GANTT_ITEM", p.EntityId).Count != 0)
                    {
                        Primitive process = One(layout, "GANTT_ITEM", p.EntityId); dynamic nativeProcess = null;
                        try
                        {
                            nativeProcess = group.GroupItems.Item(layout.Items.IndexOf(process) + 2);
                            double nativeLeft = (double)nativeProcess.Left, nativeTop = (double)nativeProcess.Top;
                            double textLeft = (double)child.Left, textTop = (double)child.Top;
                            Check(textLeft >= nativeLeft - .05 && textTop >= nativeTop - .05 && textLeft + width <= nativeLeft + (double)nativeProcess.Width + .05 && textTop + height <= nativeTop + (double)nativeProcess.Height + .05, stage + ": native process label bbox remains inside its native arrow: " + p.EntityId);
                        }
                        finally { Release((object)nativeProcess); }
                    }
                }
                else if (p.Kind == "line")
                {
                    Check(((int)child.Type == 9 || (int)child.Type == 17) && (int)child.Line.ForeColor.RGB == p.Color, stage + ": native line/color");
                    if (p.DashStyle > 0) Check((int)child.Line.DashStyle == p.DashStyle, stage + ": native separator dash style preserved");
                }
                else
                {
                    Check(p.Kind == "path" ? (int)child.Type == 5 : p.Kind == "rect" && (int)child.Type == 1, stage + ": native freeform/rectangle");
                    Check((int)child.Fill.ForeColor.RGB == p.Color && (int)child.Line.Visible == 0, stage + ": shape fill/color");
                }
                measurements.Append(fixture.Name).Append('\t').Append(stage).Append('\t').Append(i).Append('\t').Append(p.Role).Append('\t').Append(p.EntityId).Append('\t').Append(N(width)).Append('\t').Append(N(height)).Append('\t').Append(bw.HasValue ? N(bw.Value) : "").Append('\t').Append(bh.HasValue ? N(bh.Value) : "").Append('\n');
            }
            finally { Release((object)child); }
        }
    }
    private static string Normalize(string text) { return text.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n'); }
    private static void InvalidReplacements(object slideObject, object groupObject, Fixture fixture, StringBuilder measurements, StringBuilder metadata)
    {
        dynamic slide = slideObject, group = groupObject; int id = (int)group.Id; string name = (string)group.Name; int count = (int)slide.Shapes.Count;
        List<ChartSpec> invalid = new List<ChartSpec>();
        ChartSpec date = ChartStorage.Clone(fixture.Spec); date.Gantt.DateAnchors[0].Date = "2024-02-30"; invalid.Add(date);
        ChartSpec duration = ChartStorage.Clone(fixture.Spec); duration.Gantt.DateAnchors[1].Date = duration.Gantt.DateAnchors[0].Date; invalid.Add(duration);
        ChartSpec row = ChartStorage.Clone(fixture.Spec); row.Gantt.Items[0].RowId = "missing"; invalid.Add(row);
        ChartSpec anchor = ChartStorage.Clone(fixture.Spec); anchor.Gantt.Items[0].StartAnchor = "missing"; invalid.Add(anchor);
        ChartSpec format = ChartStorage.Clone(fixture.Spec); format.Gantt.Scales[0].Format = "unsupported"; invalid.Add(format);
        ChartSpec geometry = ChartStorage.Clone(fixture.Spec); geometry.Gantt.LabelFraction = Double.NaN; invalid.Add(geometry);
        foreach (ChartSpec bad in invalid)
        {
            bool rejected = false; try { OfficeRenderer.Replace(slideObject, groupObject, bad); } catch (ArgumentException) { rejected = true; } catch (FormatException) { rejected = true; } catch (NotSupportedException) { rejected = true; }
            Check(rejected && (int)slide.Shapes.Count == count && (int)group.Id == id && (string)group.Name == name, "invalid replacement preserves original native group");
            CheckGroup(groupObject, fixture, "rollback", measurements, metadata);
        }
    }
    private static void CheckPackage(string filename, List<Fixture> fixtures)
    {
        using (Package package = Package.Open(filename, FileMode.Open, FileAccess.Read))
        {
            foreach (PackagePart part in package.GetParts())
            {
                string path = part.Uri.OriginalString;
                Check(!path.StartsWith("/ppt/charts/", StringComparison.OrdinalIgnoreCase) && !path.StartsWith("/ppt/embeddings/", StringComparison.OrdinalIgnoreCase) && !path.StartsWith("/ppt/media/", StringComparison.OrdinalIgnoreCase), "package contains no Office chart, OLE or image assets");
            }
            for (int i = 0; i < fixtures.Count; i++)
            {
                XmlDocument xml = new XmlDocument(); using (Stream stream = package.GetPart(new Uri("/ppt/slides/slide" + (i + 1) + ".xml", UriKind.Relative)).GetStream(FileMode.Open, FileAccess.Read)) xml.Load(stream);
                XmlNamespaceManager ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
                Check(xml.SelectNodes("//p:pic|//p:graphicFrame|//p:oleObj", ns).Count == 0, "saved slide uses editable shapes only");
                XmlNode group = xml.SelectSingleNode("/p:sld/p:cSld/p:spTree/p:grpSp", ns); Check(group != null, "saved native group exists");
                XmlNodeList shapes = group.SelectNodes("p:sp|p:cxnSp", ns); Check(shapes.Count == fixtures[i].Names.Count, "saved primitive count");
                for (int j = 0; j < shapes.Count; j++)
                {
                    XmlNode identity = shapes[j].SelectSingleNode("p:nvSpPr/p:cNvPr|p:nvCxnSpPr/p:cNvPr", ns);
                    Check(identity != null && identity.Attributes["name"].Value == fixtures[i].Names[j], "saved native z-order preserves geometry/labels");
                }
            }
        }
    }
    private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) { try { Marshal.FinalReleaseComObject(value); } catch { } } }
}
