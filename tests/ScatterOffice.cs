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

// Native renderer integration only: synthetic data, no PIA, Excel or add-in changes.
public static class ScatterOffice
{
    private static int assertions;
    private const double NativeQuantum = 1.0 / 12700;
    private sealed class Fixture
    {
        internal string Name;
        internal ChartSpec Spec;
        internal bool Warning, BubbleScale, MicroBubble;
        internal ChartRectangle Frame = new ChartRectangle { X = 30, Y = 45, Width = 900, Height = 450 };
        internal List<string> ShapeNames = new List<string>();
    }
    private static void Check(bool condition, string message)
    { assertions++; if (!condition) throw new Exception(message); }
    private static string N(double number) { return number.ToString("R", CultureInfo.InvariantCulture); }
    private static ScatterPoint P(double? x, double? y, double? size = null, string group = null, string marker = null, string label = null)
    { return new ScatterPoint { X = x, Y = y, Size = size, Group = group, Marker = marker, Label = label ?? "Synthetic point" }; }
    private static ChartSpec Spec(string kind, params ScatterPoint[] points)
    {
        ChartSpec spec = new ChartSpec { Kind = kind, Title = "Synthetic " + kind, Scatter = new ScatterSpec { ShowLabels = false } };
        spec.Scatter.Points.AddRange(points); return spec;
    }
    private static void Fixed(ChartSpec spec, double low, double high)
    { spec.Scatter.XAxis.Minimum = low; spec.Scatter.XAxis.Maximum = high; spec.Scatter.YAxis.Minimum = low; spec.Scatter.YAxis.Maximum = high; }
    private static List<Fixture> Fixtures()
    {
        List<Fixture> fixtures = new List<Fixture>();
        ChartSpec markers = Spec("Scatter"); markers.Scatter.ShowLegend = false;
        string[] names = { "Circle", "Square", "Triangle", "Diamond", "Plus", "Cross", "None" };
        for (int i = 0; i < names.Length; i++) markers.Scatter.Points.Add(P(i, i % 3, null, "Markers", names[i], names[i]));
        fixtures.Add(new Fixture { Name = "01-explicit-markers", Spec = markers });
        ChartSpec groups = Spec("Scatter", P(1, 2, null, "North"), P(2, 4, null, "South"), P(3, 1, null, "West"), P(4, 3, null, "North"));
        groups.Scatter.MarkerScheme = "Shapes"; groups.Scatter.XTitle = "Synthetic X"; groups.Scatter.YTitle = "Synthetic Y";
        fixtures.Add(new Fixture { Name = "02-groups-and-shapes", Spec = groups });
        ChartSpec fixedAxes = Spec("Scatter", P(-8, 8), P(0, 0), P(8, -8), P(30, 30)); Fixed(fixedAxes, -10, 10);
        fixtures.Add(new Fixture { Name = "03-fixed-linear-axes", Spec = fixedAxes });
        ChartSpec log = Spec("Scatter", P(.1, .01), P(1, 1), P(10, 100), P(100, 10000));
        log.Scatter.XAxis = new ScatterAxisSpec { Logarithmic = true, Minimum = .1, Maximum = 100 };
        log.Scatter.YAxis = new ScatterAxisSpec { Logarithmic = true, Minimum = .01, Maximum = 10000 };
        fixtures.Add(new Fixture { Name = "04-logarithmic-axes", Spec = log });
        ChartSpec invalidLog = Spec("Scatter", P(-1, 2), P(0, 3), P(1, 1), P(10, 10), P(4, -4));
        invalidLog.Scatter.XAxis.Logarithmic = true; invalidLog.Scatter.YAxis.Logarithmic = true;
        fixtures.Add(new Fixture { Name = "05-log-baseline-warning", Spec = invalidLog, Warning = true });
        ChartSpec reversed = Spec("Scatter", P(2, 8), P(5, 3), P(8, 6)); Fixed(reversed, 0, 10);
        reversed.Scatter.XAxis.Reversed = true; reversed.Scatter.YAxis.Reversed = true;
        fixtures.Add(new Fixture { Name = "06-reversed-axes", Spec = reversed });
        ChartSpec area = Spec("Bubble", P(2, 2, 1, "Sizes"), P(4, 5, 4, "Sizes"), P(7, 3, 9, "Sizes"), P(3, 4, 0, "Sizes"), P(5, 6, null, "Sizes"));
        area.Scatter.ShowLegend = false; area.Scatter.BubbleSizing = "Area";
        fixtures.Add(new Fixture { Name = "07-bubble-area", Spec = area, BubbleScale = true });
        ChartSpec diameter = Spec("Bubble", P(2, 2, 1), P(4, 5, 4), P(7, 3, 9)); diameter.Scatter.ShowLegend = false; diameter.Scatter.BubbleSizing = "Diameter";
        fixtures.Add(new Fixture { Name = "08-bubble-diameter", Spec = diameter, BubbleScale = true });
        ChartSpec tiny = Spec("Bubble", P(2, 2, 1e-12), P(8, 8, 1)); tiny.Scatter.ShowLegend = false; Fixed(tiny, 0, 10);
        fixtures.Add(new Fixture { Name = "09-microbubble-no-floor", Spec = tiny, BubbleScale = true, MicroBubble = true });
        ChartSpec labels = Spec("Scatter", P(2, 2, null, "Group W", null, "WWWWWWWWWWWWWWWWWW"), P(8, 8, null, "Group CJK", null, "東京市場分析販売計画比較"));
        Fixed(labels, 0, 10); labels.Scatter.ShowLabels = true; labels.Scatter.LabelContent = "Label"; labels.Scatter.SecondLabelContent = "XY"; labels.Scatter.ShowLegend = false;
        fixtures.Add(new Fixture { Name = "10-two-long-labels", Spec = labels });
        ChartSpec partitions = Spec("Scatter", P(2, 3), P(7, 8)); Fixed(partitions, 0, 10);
        partitions.Scatter.Partitions.Add(new ScatterPartition { X1 = 5, Y1 = 0, X2 = 5, Y2 = 10, Label = "Vertical", Color = "#006699" });
        partitions.Scatter.Partitions.Add(new ScatterPartition { X1 = 0, Y1 = 5, X2 = 10, Y2 = 5, Label = "Horizontal", Color = "#996600" });
        partitions.Scatter.Partitions.Add(new ScatterPartition { X1 = 0, Y1 = 0, X2 = 10, Y2 = 10, Label = "Oblique", Color = "#990066" });
        fixtures.Add(new Fixture { Name = "11-three-partitions", Spec = partitions });
        ChartSpec trend = Spec("Scatter", P(0, 0, null, "A"), P(1, 2, null, "A"), P(2, 1, null, "A"), P(3, 3, null, "B"), P(4, 4, null, "B"));
        trend.Scatter.Trendlines.Add(new ScatterTrendline { ShowRSquared = true, ShowEquation = true, Color = "#663399" });
        trend.Scatter.Trendlines.Add(new ScatterTrendline { Group = "A", ShowRSquared = true, Color = "#007799" });
        fixtures.Add(new Fixture { Name = "12-scatter-regression", Spec = trend });
        ChartSpec weighted = Spec("Bubble", P(0, 0, 1, "Weighted"), P(1, 2, 2, "Weighted"), P(2, 1, 3, "Weighted"));
        weighted.Scatter.Trendlines.Add(new ScatterTrendline { ShowRSquared = true, ShowEquation = true });
        fixtures.Add(new Fixture { Name = "13-weighted-bubble-regression", Spec = weighted });
        ChartSpec missing = Spec("Scatter", P(0, 0), P(1, 2), P(null, 4), P(2, null));
        missing.Scatter.Points[0].Color = "#D1495B"; missing.Scatter.Points[1].Color = "#00798C";
        fixtures.Add(new Fixture { Name = "14-missing-coordinates", Spec = missing });
        ChartSpec thousand = Spec("Scatter"); thousand.Scatter.ShowLegend = false; thousand.Scatter.ShowLabels = null;
        for (int i = 0; i < 1000; i++) thousand.Scatter.Points.Add(P(i, (i * 37) % 101, null, "Synthetic", "Circle", "Point" + i.ToString(CultureInfo.InvariantCulture)));
        fixtures.Add(new Fixture { Name = "15-thousand-points-metadata", Spec = thousand });
        ChartSpec adjacentX = Spec("Scatter", P(1e16, 0), P(1e16 + 2, 1)); adjacentX.Scatter.ShowLegend = false;
        adjacentX.Scatter.Trendlines.Add(new ScatterTrendline { ShowRSquared = true, ShowEquation = true });
        fixtures.Add(new Fixture { Name = "16-adjacent-large-x-trend", Spec = adjacentX });
        ChartSpec adjacentY = Spec("Scatter", P(0, 1e16), P(1, 1e16 + 2)); adjacentY.Scatter.ShowLegend = false;
        adjacentY.Scatter.YAxis.Minimum = 1e16; adjacentY.Scatter.YAxis.Maximum = 1e16 + 2;
        adjacentY.Scatter.XAxis.Reversed = true; adjacentY.Scatter.YAxis.Reversed = true;
        adjacentY.Scatter.Trendlines.Add(new ScatterTrendline { ShowRSquared = true, ShowEquation = true });
        fixtures.Add(new Fixture { Name = "17-adjacent-large-y-reversed-trend", Spec = adjacentY });
        return fixtures;
    }

    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length != 1) { Console.Error.WriteLine("Usage: ScatterOffice.exe OUTPUT_DIRECTORY"); return 2; }
        Process[] existing = Process.GetProcessesByName("POWERPNT"); bool running = existing.Length > 0;
        foreach (Process process in existing) process.Dispose();
        if (running) { Console.Error.WriteLine("Close PowerPoint before isolated scatter integration tests."); return 2; }
        string output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output);
        string filename = Path.Combine(output, "OpenSlideCharts-scatter.pptx");
        List<Fixture> fixtures = Fixtures(); dynamic ppt = null, deck = null;
        StringBuilder evidence = new StringBuilder("fixture\tstage\titem\trole\tentity\twidth_pt\theight_pt\ttext_bound_width_pt\ttext_bound_height_pt\n");
        try
        {
            foreach (Fixture fixture in fixtures) CheckLayout(fixture);
            ppt = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
            if ((int)ppt.Presentations.Count != 0) throw new InvalidOperationException("Another presentation exists; test aborted.");
            ppt.Visible = -1; deck = ppt.Presentations.Add(-1);
            deck.PageSetup.SlideWidth = 960f; deck.PageSetup.SlideHeight = 540f;
            for (int i = 0; i < fixtures.Count; i++)
            {
                Fixture fixture = fixtures[i]; dynamic slide = null, group = null; Stopwatch elapsed = Stopwatch.StartNew();
                try
                {
                    slide = deck.Slides.Add(i + 1, 12); slide.FollowMasterBackground = 0;
                    slide.Background.Fill.Solid(); slide.Background.Fill.ForeColor.RGB = 0xE8F0F4;
                    ChartRectangle frame = fixture.Frame;
                    group = OfficeRenderer.Render((object)slide, fixture.Spec, frame.X, frame.Y, frame.Width, frame.Height);
                    CheckGroup((object)group, fixture, "render", evidence);
                    if (fixture.Name == "07-bubble-area")
                    {
                        ChartSpec changed = ChartStorage.Clone(fixture.Spec);
                        changed.Scatter.Points[0].X = 3; changed.Scatter.Points[0].Y = 4; changed.Scatter.Points[2].Size = 16;
                        changed.Scatter.MaxBubbleDiameter = 60;
                        int previous = (int)group.Id;
                        object fresh = OfficeRenderer.Replace((object)slide, (object)group, changed);
                        Release((object)group); group = fresh; fixture.Spec = changed;
                        Check((int)group.Id != previous && (int)slide.Shapes.Count == 1, "valid replacement creates one new group");
                        CheckGroup((object)group, fixture, "replace", evidence); InvalidReplacements((object)slide, (object)group, fixture, evidence);
                    }
                    slide.Export(Path.Combine(output, fixture.Name + ".png"), "PNG", 1440, 810);
                    Console.WriteLine("Rendered/verified/exported " + fixture.Name + " in " + elapsed.ElapsedMilliseconds + " ms");
                }
                finally { Release((object)group); Release((object)slide); }
            }
            deck.SaveAs(filename, 24); deck.Close(); Release((object)deck); deck = null;
            CheckPackage(filename, fixtures);
            deck = ppt.Presentations.Open(filename, 0, 0, 0);
            Check((int)deck.Slides.Count == fixtures.Count, "slide count after reopening");
            for (int i = 0; i < fixtures.Count; i++)
            {
                dynamic slide = null, group = null; Stopwatch elapsed = Stopwatch.StartNew();
                try
                {
                    slide = deck.Slides.Item(i + 1); Check((int)slide.Shapes.Count == 1, "one group after reopening");
                    group = slide.Shapes.Item(1); CheckGroup((object)group, fixtures[i], "reopen", evidence);
                    slide.Export(Path.Combine(output, fixtures[i].Name + "-reopened.png"), "PNG", 1440, 810);
                    Console.WriteLine("Reopened/verified/exported " + fixtures[i].Name + " in " + elapsed.ElapsedMilliseconds + " ms");
                }
                finally { Release((object)group); Release((object)slide); }
            }
            Console.WriteLine("PASS: " + assertions + " scatter Office assertions; " + fixtures.Count + " synthetic fixtures, native geometry/text, metadata, save/reopen and rollback.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            try { File.WriteAllText(Path.Combine(output, "scatter-native-measurements.tsv"), evidence.ToString(), new UTF8Encoding(false)); }
            catch (Exception error) { Console.Error.WriteLine("Could not save measurements: " + error.Message); }
            if ((object)deck != null) { try { deck.Close(); } catch { } Release((object)deck); }
            if ((object)ppt != null)
            {
                try { if ((int)ppt.Presentations.Count == 0) ppt.Quit(); else Console.Error.WriteLine("Other presentations exist: PowerPoint left open."); } catch { }
                Release((object)ppt);
            }
        }
    }
    private static Layout CheckLayout(Fixture fixture)
    {
        Layout layout = ChartEngine.Build(fixture.Spec, fixture.Frame.Width, fixture.Frame.Height);
        bool warning = false, micro = false; double previousSize = Double.PositiveInfinity;
        int circles = 0, primaryLabels = 0, secondaryLabels = 0, partitions = 0, trends = 0;
        bool weightedRSquared = false; HashSet<int> pointIds = new HashSet<int>(); double maxSize = 0;
        foreach (ScatterPoint p in fixture.Spec.Scatter.Points) if (p.Size.HasValue) maxSize = Math.Max(maxSize, p.Size.Value);
        foreach (Primitive p in layout.Items)
        {
            if (p.Role == "SCATTER_WARNING") warning = true;
            if (p.Role == "SCATTER_LABEL") primaryLabels++;
            if (p.Role == "SCATTER_SECOND_LABEL") secondaryLabels++;
            if (p.Role == "SCATTER_PARTITION") partitions++;
            if (p.Role == "SCATTER_TREND") trends++;
            if (p.Role == "SCATTER_TREND_LABEL" && p.Text != null && p.Text.Contains("R\u00b2 = 0.04706")) weightedRSquared = true;
            if (p.Role != "SCATTER_POINT") continue;
            int index; Check(Int32.TryParse(p.EntityId, out index) && index >= 0 && index < fixture.Spec.Scatter.Points.Count, "stable scatter source identity");
            pointIds.Add(index);
            ScatterPoint point = fixture.Spec.Scatter.Points[index];
            if (fixture.BubbleScale)
            {
                Check(p.Kind == "ellipse", "bubble uses an editable circle"); circles++;
                Check(point.Size.HasValue && point.Size.Value > 0, "zero/missing bubble size omitted");
                double fraction = point.Size.Value / maxSize;
                double expected = fixture.Spec.Scatter.MaxBubbleDiameter * (fixture.Spec.Scatter.BubbleSizing == "Area" ? Math.Sqrt(fraction) : fraction);
                Check(Math.Abs(p.W - expected) <= Math.Max(1e-12, expected * 1e-12) && Math.Abs(p.H - expected) <= Math.Max(1e-12, expected * 1e-12), "bubble diameter depends on Size only");
                Check(point.Size.Value <= previousSize, "large bubbles emitted first, small bubbles in front"); previousSize = point.Size.Value;
                if (fixture.MicroBubble && index == 0)
                { Check(Math.Abs(p.W - 5e-5) < 1e-12, "microbubble has 0.00005pt diameter without core floor"); micro = true; }
            }
        }
        if (fixture.Warning) Check(warning, "nonpositive logarithmic coordinates produce visible warning");
        if (fixture.BubbleScale)
        {
            int expected = 0; foreach (ScatterPoint point in fixture.Spec.Scatter.Points) if (point.Size.HasValue && point.Size.Value > 0) expected++;
            Check(circles == expected, "one native circle for each positive-size bubble");
        }
        if (fixture.MicroBubble) Check(micro, "microbubble primitive retained");
        if (fixture.Name == "01-explicit-markers") Check(pointIds.Count == 6 && !pointIds.Contains(6), "six marker styles rendered; None omitted");
        if (fixture.Name == "10-two-long-labels") Check(primaryLabels == 2 && secondaryLabels == 2, "primary and secondary labels both present");
        if (fixture.Name == "11-three-partitions") Check(partitions == 3, "vertical, horizontal and oblique partitions present");
        if (fixture.Name == "12-scatter-regression") Check(trends == 2, "overall and group regression lines present");
        if (fixture.Name == "13-weighted-bubble-regression") Check(trends == 1 && weightedRSquared, "known weighted R squared 4/85 displayed");
        if (fixture.Name == "14-missing-coordinates") Check(pointIds.Count == 2 && pointIds.Contains(0) && pointIds.Contains(1), "missing coordinates omitted; zero remains valid");
        if (fixture.Name == "15-thousand-points-metadata")
        {
            Check(pointIds.Count == 1000, "all 1000 points retain source identities");
            Check(primaryLabels == 0 && secondaryLabels == 0, "automatic labels disabled above 300 points");
            Check(MetadataTags(fixture.Spec) > 889, "1000-point metadata exceeds previous 889-tag limit");
        }
        if (fixture.Name == "16-adjacent-large-x-trend" || fixture.Name == "17-adjacent-large-y-reversed-trend")
        {
            Primitive line = null; foreach (Primitive p in layout.Items) if (p.Role == "SCATTER_TREND") line = p;
            Check(line != null && trends == 1 && pointIds.Count == 2, "adjacent-coordinate two-point regression remains visible");
            double length = Math.Sqrt(line.W * line.W + line.H * line.H); Check(length > 0, "adjacent-coordinate trend segment nondegenerate");
            foreach (Primitive p in layout.Items) if (p.Role == "SCATTER_POINT")
            {
                double cx = p.X + p.W / 2, cy = p.Y + p.H / 2;
                double distance = Math.Abs(line.W * (cy - line.Y) - line.H * (cx - line.X)) / length;
                Check(distance < 1e-7, "exact trend passes adjacent-coordinate marker center");
            }
        }
        return layout;
    }
    private static int MetadataTags(ChartSpec spec)
    { return (Encoding.UTF8.GetByteCount(ChartStorage.Serialize(spec)) * 2 + 179) / 180; }
    private static void CheckGroup(object value, Fixture fixture, string stage, StringBuilder evidence)
    {
        dynamic group = value; Layout layout = CheckLayout(fixture); ChartRectangle frame = fixture.Frame;
        Check((int)group.Type == 6 && (string)group.Tags.Item("OSC_VERSION") == "1", stage + ": native tagged group");
        string actualModel = ChartStorage.Serialize(ChartStorage.Read(value)), expectedModel = ChartStorage.Serialize(fixture.Spec);
        int difference = 0; while (difference < actualModel.Length && difference < expectedModel.Length && actualModel[difference] == expectedModel[difference]) difference++;
        Check(actualModel == expectedModel, fixture.Name + " " + stage + ": complete model roundtrip; first difference " + difference +
            "; actual=" + actualModel.Substring(Math.Max(0, difference - 20), Math.Min(100, actualModel.Length - Math.Max(0, difference - 20))) +
            "; expected=" + expectedModel.Substring(Math.Max(0, difference - 20), Math.Min(100, expectedModel.Length - Math.Max(0, difference - 20))));
        if (fixture.Name == "15-thousand-points-metadata")
        {
            int tags = Int32.Parse((string)group.Tags.Item("OSC_COUNT"), CultureInfo.InvariantCulture);
            Check(tags > 889 && tags == MetadataTags(fixture.Spec), stage + ": expanded metadata tags retained without truncation");
        }
        Check((int)group.GroupItems.Count == layout.Items.Count + 1, stage + ": native primitives plus transparent frame only");
        Check(Math.Abs((double)group.Left - frame.X) < .05 && Math.Abs((double)group.Top - frame.Y) < .05 && Math.Abs((double)group.Width - frame.Width) < .05 && Math.Abs((double)group.Height - frame.Height) < .05, stage + ": frame preserved");
        fixture.ShapeNames.Clear();
        dynamic bounds = group.GroupItems.Item(1); fixture.ShapeNames.Add((string)bounds.Name); Release((object)bounds);
        for (int i = 0; i < layout.Items.Count; i++)
        {
            Primitive p = layout.Items[i]; dynamic child = null;
            try
            {
                child = group.GroupItems.Item(i + 2); fixture.ShapeNames.Add((string)child.Name);
                if (!String.IsNullOrEmpty(p.Role))
                {
                    Check((string)child.Tags.Item("OSC_ROLE") == p.Role, stage + ": role tag");
                    Check((string)child.Tags.Item("OSC_ITEM") == i.ToString(CultureInfo.InvariantCulture), stage + ": layout index tag");
                }
                double width = (double)child.Width, height = (double)child.Height;
                double? boundWidth = null, boundHeight = null;
                if (p.Kind == "ellipse" || p.Kind == "rect" || p.Kind == "text")
                    Check(Math.Abs((double)child.Left - frame.X - p.X) < .05 && Math.Abs((double)child.Top - frame.Y - p.Y) < .05, stage + ": native position matches plotted coordinates");
                if (p.Role == "SCATTER_POINT") Check((int)child.Fill.ForeColor.RGB == p.Color, stage + ": point/group color retained");
                if (p.Kind == "ellipse")
                {
                    Check((int)child.Type == 1 && (int)child.AutoShapeType == 9, stage + ": native editable ellipse");
                    double tolerance = p.Role == "SCATTER_POINT" ? NativeQuantum * 1.1 + 1e-6 : .05;
                    Check(Math.Abs(width - p.W) <= tolerance && Math.Abs(height - p.H) <= tolerance, stage + ": native ellipse diameter matches layout within EMU quantization");
                    if (fixture.MicroBubble && p.Role == "SCATTER_POINT" && p.EntityId == "0")
                        Check(width < .001 && height < .001, stage + ": native microbubble has no 0.05pt display floor");
                }
                else if (p.Kind == "path") Check((int)child.Type == 5, stage + ": native freeform marker");
                else if (p.Kind == "rect") Check((int)child.Type == 1, stage + ": native rectangular marker");
                else if (p.Kind == "line")
                {
                    Check((int)child.Type == 9 || (int)child.Type == 17, stage + ": native line");
                    if (p.Role == "SCATTER_TREND")
                        Check(Math.Abs(width - Math.Abs(p.W)) < .05 && Math.Abs(height - Math.Abs(p.H)) < .05 &&
                            Math.Abs((double)child.Left - frame.X - Math.Min(p.X, p.X + p.W)) < .05 &&
                            Math.Abs((double)child.Top - frame.Y - Math.Min(p.Y, p.Y + p.H)) < .05, stage + ": native trend geometry preserves exact clipped segment");
                }
                else if (p.Kind == "text")
                {
                    string text = Normalize((string)child.TextFrame.TextRange.Text);
                    Check(text == Normalize(p.Text ?? ""), stage + ": native editable text content");
                    Check(Math.Abs(width - Math.Max(1, p.W)) < .05 && Math.Abs(height - Math.Max(1, p.H)) < .05, stage + ": text box geometry");
                    if (p.Role == "SCATTER_LABEL" || p.Role == "SCATTER_SECOND_LABEL")
                    {
                        boundWidth = (double)child.TextFrame.TextRange.BoundWidth; boundHeight = (double)child.TextFrame.TextRange.BoundHeight;
                        Check(boundWidth <= width + .5 && boundHeight <= height + .5, stage + ": text glyph bounds " + N(boundWidth.Value) + " x " + N(boundHeight.Value) + " fit label box " + N(width) + " x " + N(height) + ": " + text);
                    }
                }
                else Check(false, stage + ": unexpected primitive kind " + p.Kind);
                if (evidence != null) evidence.Append(fixture.Name).Append('\t').Append(stage).Append('\t').Append(i).Append('\t').Append(p.Role).Append('\t').Append(p.EntityId).Append('\t').Append(N(width)).Append('\t').Append(N(height)).Append('\t').Append(boundWidth.HasValue ? N(boundWidth.Value) : "").Append('\t').Append(boundHeight.HasValue ? N(boundHeight.Value) : "").Append('\n');
            }
            finally { Release((object)child); }
        }
    }
    private static string Normalize(string value) { return value.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n'); }
    private static void InvalidReplacements(object slideObject, object groupObject, Fixture fixture, StringBuilder evidence)
    {
        dynamic slide = slideObject, group = groupObject; int count = (int)slide.Shapes.Count, id = (int)group.Id; string name = (string)group.Name;
        List<ChartSpec> invalid = new List<ChartSpec>();
        ChartSpec nonfinite = ChartStorage.Clone(fixture.Spec); nonfinite.Scatter.Points[0].X = Double.NaN; invalid.Add(nonfinite);
        ChartSpec size = ChartStorage.Clone(fixture.Spec); size.Scatter.Points[0].Size = -1; invalid.Add(size);
        ChartSpec axis = ChartStorage.Clone(fixture.Spec); axis.Scatter.XAxis.Minimum = 10; axis.Scatter.XAxis.Maximum = 0; invalid.Add(axis);
        ChartSpec logarithmic = ChartStorage.Clone(fixture.Spec); logarithmic.Scatter.XAxis.Logarithmic = true; logarithmic.Scatter.XAxis.Minimum = 0; invalid.Add(logarithmic);
        ChartSpec mode = ChartStorage.Clone(fixture.Spec); mode.Scatter.BubbleSizing = "Unknown"; invalid.Add(mode);
        foreach (ChartSpec bad in invalid)
        {
            bool rejected = false;
            try { OfficeRenderer.Replace(slideObject, groupObject, bad); }
            catch (ArgumentException) { rejected = true; }
            catch (NotSupportedException) { rejected = true; }
            Check(rejected, "invalid replacement rejected before deleting old chart");
            Check((int)slide.Shapes.Count == count && (int)group.Id == id && (string)group.Name == name, "invalid replacement preserves old native shape identity");
            CheckGroup(groupObject, fixture, "rollback", evidence);
        }
    }
    private static void CheckPackage(string filename, List<Fixture> fixtures)
    {
        using (Package package = Package.Open(filename, FileMode.Open, FileAccess.Read))
        {
            foreach (PackagePart part in package.GetParts())
            {
                string path = part.Uri.OriginalString;
                Check(!path.StartsWith("/ppt/charts/", StringComparison.OrdinalIgnoreCase) && !path.StartsWith("/ppt/embeddings/", StringComparison.OrdinalIgnoreCase) && !path.StartsWith("/ppt/media/", StringComparison.OrdinalIgnoreCase), "native chart contains no embedded chart, OLE or image asset: " + path);
            }
            for (int i = 0; i < fixtures.Count; i++)
            {
                XmlDocument xml = new XmlDocument();
                using (Stream stream = package.GetPart(new Uri("/ppt/slides/slide" + (i + 1) + ".xml", UriKind.Relative)).GetStream(FileMode.Open, FileAccess.Read)) xml.Load(stream);
                XmlNamespaceManager ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main"); ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
                Check(xml.SelectNodes("//p:pic|//p:graphicFrame|//p:oleObj", ns).Count == 0, "slide stores editable shapes instead of pictures/Office charts/OLE");
                XmlNode group = xml.SelectSingleNode("/p:sld/p:cSld/p:spTree/p:grpSp", ns); Check(group != null, "saved chart group exists");
                XmlNodeList shapes = group.SelectNodes("p:sp|p:cxnSp", ns);
                Check(shapes.Count == fixtures[i].ShapeNames.Count, "saved native shape count");
                for (int s = 0; s < shapes.Count; s++)
                {
                    XmlNode identity = shapes[s].SelectSingleNode("p:nvSpPr/p:cNvPr|p:nvCxnSpPr/p:cNvPr", ns);
                    Check(identity != null && identity.Attributes["name"].Value == fixtures[i].ShapeNames[s], "saved shape z-order preserves bubble ordering");
                }
            }
        }
    }
    private static void Release(object value)
    { if (value != null && Marshal.IsComObject(value)) { try { Marshal.FinalReleaseComObject(value); } catch { } } }
}
