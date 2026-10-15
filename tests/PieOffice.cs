using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using OpenSlideCharts;

// Renderer-only integration test. Synthetic fixtures; no PIA, Excel or add-in registration.
public static class PieOffice
{
    private static int assertions;
    private const int SlideWidth = 960, SlideHeight = 540, ExportWidth = 1440, ExportHeight = 810;
    private static readonly Color Background = Color.FromArgb(217, 199, 163);
    private sealed class Fixture
    {
        internal string Name;
        internal ChartSpec Spec;
        internal ChartRectangle Frame = new ChartRectangle { X = 30, Y = 45, Width = 900, Height = 450 };
        internal bool ClearHole, FullCircle;
        internal List<string> PathNames = new List<string>();
    }
    private struct Point
    {
        internal double X, Y;
        internal Point(double x, double y) { X = x; Y = y; }
    }
    private static void Check(bool condition, string message)
    { assertions++; if (!condition) throw new Exception(message); }
    private static string Invariant(double value) { return value.ToString("R", CultureInfo.InvariantCulture); }
    private static int Ole(Color color) { return color.R | color.G << 8 | color.B << 16; }
    private static ChartSpec Spec(string kind, params double[] values)
    {
        ChartSpec spec = new ChartSpec { Kind = kind, Title = "Synthetic " + kind, Pie = new PieOptions() };
        for (int i = 0; i < values.Length; i++) spec.Rows.Add(new ChartRow { Label = "Category " + (i + 1), Value = values[i] });
        return spec;
    }
    private static void HideLabels(ChartSpec spec)
    { spec.Pie.LabelContent = "None"; spec.Pie.ShowCategory = false; }
    private static List<Fixture> Fixtures()
    {
        List<Fixture> result = new List<Fixture>();
        ChartSpec full = Spec("Pie", 100); HideLabels(full);
        result.Add(new Fixture { Name = "01-full-circle-pie", Spec = full, FullCircle = true });
        ChartSpec ring = Spec("Doughnut", 100); HideLabels(ring); ring.Pie.HoleRatio = .5;
        result.Add(new Fixture { Name = "02-full-circle-doughnut-hole", Spec = ring, ClearHole = true, FullCircle = true });
        ChartSpec center = Spec("Doughnut", 42); HideLabels(center); center.Pie.CenterTitle = "42 units";
        result.Add(new Fixture { Name = "03-center-title", Spec = center, FullCircle = true });
        ChartSpec multi = Spec("Doughnut", 4, 3, 2, 1); HideLabels(multi); multi.Pie.HoleRatio = .58;
        result.Add(new Fixture { Name = "04-multiple-slices-true-hole", Spec = multi, ClearHole = true });
        ChartSpec counter = Spec("Pie", 15, 35, 50); counter.Pie.StartAngle = 37; counter.Pie.Order = "SheetCounterclockwise";
        counter.Pie.LabelContent = "Both"; counter.Pie.PercentDigits = 1;
        result.Add(new Fixture { Name = "05-rotated-counterclockwise", Spec = counter });
        ChartSpec ascending = Spec("Doughnut", 40, 10, 30, 20); ascending.Pie.Order = "Ascending"; ascending.Pie.LabelContent = "Value";
        ascending.Pie.Slices.Add(new PieSliceStyle { Index = 0, Explode = .25, Color = "#D1495B" });
        ascending.Pie.Slices.Add(new PieSliceStyle { Index = 1, Explode = .1, Color = "#00798C" });
        result.Add(new Fixture { Name = "06-sorted-exploded", Spec = ascending });
        ChartSpec descending = Spec("Pie", 8, 3, 5); descending.Pie.Order = "Descending";
        result.Add(new Fixture { Name = "07-descending", Spec = descending });
        ChartSpec tiny = Spec("Pie", 1, .0001, 0, 2); HideLabels(tiny);
        result.Add(new Fixture { Name = "08-tiny-and-zero", Spec = tiny });
        ChartSpec minimum = Spec("Doughnut", 6, 4); HideLabels(minimum);
        result.Add(new Fixture { Name = "09-minimum-frame", Spec = minimum, Frame = new ChartRectangle { X = 320, Y = 150, Width = 320, Height = 240 }, ClearHole = true });
        ChartSpec longLabels = Spec("Pie", 6, 4); longLabels.Pie.LabelContent = "Both";
        longLabels.Rows[0].Label = "Synthetic category with a deliberately long descriptive label";
        longLabels.Rows[1].Label = "Another long category for wrapping and frame verification";
        result.Add(new Fixture { Name = "10-long-labels", Spec = longLabels });
        ChartSpec wide = Spec("Doughnut", 66, 34); wide.Pie.LabelContent = "Both";
        wide.Rows[0].Label = new string('W', 100); wide.Rows[1].Label = new string('\u4E00', 100); wide.Pie.CenterTitle = new string('W', 30);
        result.Add(new Fixture { Name = "12-wide-glyphs", Spec = wide });
        ChartSpec nativePrecision = Spec("Pie", 1e-12, 1); HideLabels(nativePrecision);
        result.Add(new Fixture { Name = "11-native-float-tiny", Spec = nativePrecision });
        return result;
    }

    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length != 1) { Console.Error.WriteLine("Usage: PieOffice.exe OUTPUT_DIRECTORY"); return 2; }
        Process[] existing = Process.GetProcessesByName("POWERPNT");
        bool running = existing.Length != 0; foreach (Process process in existing) process.Dispose();
        if (running) { Console.Error.WriteLine("Close PowerPoint before running isolated pie integration tests."); return 2; }
        string output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output);
        string pptx = Path.Combine(output, "OpenSlideCharts-pies.pptx");
        dynamic ppt = null, deck = null;
        List<Fixture> fixtures = Fixtures();
        try
        {
            // Validate all fixtures before launching the native application.
            foreach (Fixture fixture in fixtures) ChartEngine.Build(fixture.Spec, fixture.Frame.Width, fixture.Frame.Height);
            ppt = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
            if ((int)ppt.Presentations.Count != 0) throw new InvalidOperationException("PowerPoint contains another presentation; test aborted.");
            ppt.Visible = -1; deck = ppt.Presentations.Add(-1);
            deck.PageSetup.SlideWidth = (float)SlideWidth; deck.PageSetup.SlideHeight = (float)SlideHeight;
            StringBuilder manifest = new StringBuilder("fixture\timage\tcenter_x_px\tcenter_y_px\tradius_px\texpect_clear_hole\tbackground_rgb\n");
            for (int i = 0; i < fixtures.Count; i++)
            {
                Fixture fixture = fixtures[i]; dynamic slide = null, group = null;
                try
                {
                    slide = deck.Slides.Add(i + 1, 12);
                    slide.FollowMasterBackground = 0; slide.Background.Fill.Solid(); slide.Background.Fill.ForeColor.RGB = Ole(Background);
                    ChartRectangle frame = fixture.Frame;
                    group = OfficeRenderer.Render((object)slide, fixture.Spec, frame.X, frame.Y, frame.Width, frame.Height);
                    CheckGroup((object)group, fixture, "render");
                    if (i == 4)
                    {
                        ChartSpec replacement = ChartStorage.Clone(fixture.Spec);
                        replacement.Rows[0].Value = 25; replacement.Pie.StartAngle = 73;
                        int oldId = (int)group.Id;
                        object replaced = OfficeRenderer.Replace((object)slide, (object)group, replacement);
                        Release((object)group); group = replaced; fixture.Spec = replacement;
                        Check((int)slide.Shapes.Count == 1 && (int)group.Id != oldId, "valid replacement creates exactly one fresh group");
                        CheckGroup((object)group, fixture, "replace");
                        CheckInvalidReplacement((object)slide, (object)group, fixture);
                    }
                    string imageName = fixture.Name + ".png";
                    slide.Export(Path.Combine(output, imageName), "PNG", ExportWidth, ExportHeight);
                    PixelEvidence(fixture, Path.Combine(output, imageName), manifest);
                    Console.WriteLine("Rendered " + fixture.Name);
                }
                finally { Release((object)group); Release((object)slide); }
            }
            File.WriteAllText(Path.Combine(output, "pie-pixel-probes.tsv"), manifest.ToString(), new UTF8Encoding(false));
            deck.SaveAs(pptx, 24); deck.Close(); Release((object)deck); deck = null;
            CheckPackage(pptx, fixtures);
            deck = ppt.Presentations.Open(pptx, 0, 0, 0);
            Check((int)deck.Slides.Count == fixtures.Count, "slide count after reopening");
            for (int i = 0; i < fixtures.Count; i++)
            {
                dynamic slide = null, group = null;
                try
                {
                    slide = deck.Slides.Item(i + 1); Check((int)slide.Shapes.Count == 1, "one native chart after reopening");
                    group = slide.Shapes.Item(1); CheckGroup((object)group, fixtures[i], "reopen");
                    string reopened = Path.Combine(output, fixtures[i].Name + "-reopened.png");
                    slide.Export(reopened, "PNG", ExportWidth, ExportHeight);
                    PixelEvidence(fixtures[i], reopened, null);
                }
                finally { Release((object)group); Release((object)slide); }
            }
            Console.WriteLine("PASS: " + assertions + " pie Office assertions; " + fixtures.Count + " synthetic fixtures, native freeforms, XML, hole pixels, save/reopen, replacement rollback.");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
        finally
        {
            // Close only the presentation created/opened by this test. If a
            // user opened another deck meanwhile, leave PowerPoint running.
            if ((object)deck != null) { try { deck.Close(); } catch { } Release((object)deck); }
            if ((object)ppt != null)
            {
                try { if ((int)ppt.Presentations.Count == 0) ppt.Quit(); else Console.Error.WriteLine("Other presentations exist: PowerPoint left open."); } catch { }
                Release((object)ppt);
            }
        }
    }

    private static void CheckGroup(object value, Fixture fixture, string stage)
    {
        dynamic group = value; Layout layout = ChartEngine.Build(fixture.Spec, fixture.Frame.Width, fixture.Frame.Height);
        Check((int)group.Type == 6, stage + ": native group");
        Check((string)group.Tags.Item("OSC_VERSION") == "1", stage + ": chart metadata tag");
        Check(ChartStorage.Serialize(ChartStorage.Read(value)) == ChartStorage.Serialize(fixture.Spec), stage + ": exact embedded model");
        Check((int)group.GroupItems.Count == layout.Items.Count + 1, stage + ": no extra masking shape");
        Check(Math.Abs((double)group.Left - fixture.Frame.X) < .05 && Math.Abs((double)group.Top - fixture.Frame.Y) < .05 &&
            Math.Abs((double)group.Width - fixture.Frame.Width) < .05 && Math.Abs((double)group.Height - fixture.Frame.Height) < .05, stage + ": frame preserved");
        fixture.PathNames.Clear(); int slices = 0;
        for (int i = 0; i < layout.Items.Count; i++)
        {
            Primitive primitive = layout.Items[i]; dynamic child = null;
            try
            {
                child = group.GroupItems.Item(i + 2);
                if (!String.IsNullOrEmpty(primitive.Role))
                {
                    Check((string)child.Tags.Item("OSC_ROLE") == primitive.Role, stage + ": primitive role");
                    Check((string)child.Tags.Item("OSC_ITEM") == i.ToString(CultureInfo.InvariantCulture), stage + ": primitive index tag");
                }
                if (primitive.Kind == "path")
                {
                    Check((int)child.Type == 5, stage + ": editable msoFreeform");
                    Check((int)child.Nodes.Count >= 3, stage + ": native freeform nodes");
                    Check((int)child.Fill.ForeColor.RGB == primitive.Color && (int)child.Line.Visible == 0, stage + ": slice appearance");
                    Check(primitive.Role == "PIE_SLICE", stage + ": slice role");
                    int index; Check(Int32.TryParse(primitive.EntityId, out index) && index >= 0 && index < fixture.Spec.Rows.Count && fixture.Spec.Rows[index].Value > 0, stage + ": stable source row identity");
                    fixture.PathNames.Add((string)child.Name); slices++;
                }
                else if (primitive.Kind == "text")
                {
                    string actual = ((string)child.TextFrame.TextRange.Text).Replace("\r\n", "\n").Replace('\r', '\n');
                    Check(actual == (primitive.Text ?? "").Replace("\r\n", "\n").Replace('\r', '\n'), stage + ": native editable text");
                    Check(Math.Abs((double)child.Width - Math.Max(1, primitive.W)) < .05 && Math.Abs((double)child.Height - Math.Max(1, primitive.H)) < .05, stage + ": text frame preserved");
                    if (primitive.Role == "PIE_LABEL" || primitive.Role == "PIE_CENTER_TITLE")
                    {
                        Check((double)child.TextFrame.TextRange.BoundWidth <= (double)child.Width + .5, stage + ": label glyphs fit native textbox width");
                        Check((double)child.TextFrame.TextRange.BoundHeight <= (double)child.Height + .5, stage + ": label glyphs fit native textbox height");
                    }
                }
            }
            finally { Release((object)child); }
        }
        int positive = 0; foreach (ChartRow row in fixture.Spec.Rows) if (row.Value > 0) positive++;
        Check(slices == positive, stage + ": one slice per positive row; zero omitted");
    }
    private static void CheckInvalidReplacement(object slideObject, object groupObject, Fixture fixture)
    {
        dynamic slide = slideObject, group = groupObject;
        int id = (int)group.Id, count = (int)slide.Shapes.Count; string name = (string)group.Name;
        List<ChartSpec> invalid = new List<ChartSpec>();
        ChartSpec negative = ChartStorage.Clone(fixture.Spec); negative.Rows[0].Value = -1; invalid.Add(negative);
        ChartSpec zero = ChartStorage.Clone(fixture.Spec); foreach (ChartRow row in zero.Rows) row.Value = 0; invalid.Add(zero);
        ChartSpec nonfinite = ChartStorage.Clone(fixture.Spec); nonfinite.Rows[0].Value = Double.NaN; invalid.Add(nonfinite);
        ChartSpec hole = ChartStorage.Clone(fixture.Spec); hole.Pie.HoleRatio = 1; invalid.Add(hole);
        ChartSpec order = ChartStorage.Clone(fixture.Spec); order.Pie.Order = "Unknown"; invalid.Add(order);
        foreach (ChartSpec bad in invalid)
        {
            bool rejected = false;
            try { OfficeRenderer.Replace(slideObject, groupObject, bad); }
            catch (ArgumentException) { rejected = true; }
            catch (NotSupportedException) { rejected = true; }
            Check(rejected, "invalid pie replacement rejected");
            Check((int)slide.Shapes.Count == count && (int)group.Id == id && (string)group.Name == name, "invalid replacement preserves native group identity");
            CheckGroup(groupObject, fixture, "rollback");
        }
    }

    private static List<Point> Sample(ShapePath path)
    {
        List<Point> points = new List<Point>(); Point current = new Point(path.StartX, path.StartY); points.Add(current);
        foreach (PathSegment segment in path.Segments)
        {
            if (segment.Curve)
            {
                Point start = current;
                for (int i = 1; i <= 64; i++)
                {
                    double t = i / 64.0, u = 1 - t;
                    points.Add(new Point(u * u * u * start.X + 3 * u * u * t * segment.C1X + 3 * u * t * t * segment.C2X + t * t * t * segment.X,
                        u * u * u * start.Y + 3 * u * u * t * segment.C1Y + 3 * u * t * t * segment.C2Y + t * t * t * segment.Y));
                }
            }
            else points.Add(new Point(segment.X, segment.Y));
            current = new Point(segment.X, segment.Y);
        }
        return points;
    }
    private static double[] Bounds(List<Point> points)
    {
        double x0 = Double.PositiveInfinity, y0 = x0, x1 = Double.NegativeInfinity, y1 = x1;
        foreach (Point p in points) { x0 = Math.Min(x0, p.X); y0 = Math.Min(y0, p.Y); x1 = Math.Max(x1, p.X); y1 = Math.Max(y1, p.Y); }
        return new double[] { x0, y0, x1, y1 };
    }
    private static void PixelEvidence(Fixture fixture, string png, StringBuilder manifest)
    {
        if (!fixture.ClearHole && !(fixture.FullCircle && fixture.Spec.Kind == "Pie")) return;
        List<Point> points = new List<Point>();
        foreach (Primitive p in ChartEngine.Build(fixture.Spec, fixture.Frame.Width, fixture.Frame.Height).Items) if (p.Kind == "path") points.AddRange(Sample(p.Path));
        double[] bounds = Bounds(points); double scale = ExportWidth / (double)SlideWidth;
        double cx = (fixture.Frame.X + (bounds[0] + bounds[2]) / 2) * scale;
        double cy = (fixture.Frame.Y + (bounds[1] + bounds[3]) / 2) * scale;
        double radius = Math.Min(bounds[2] - bounds[0], bounds[3] - bounds[1]) * scale / 2;
        if (manifest != null) manifest.Append(fixture.Name).Append('\t').Append(Path.GetFileName(png)).Append('\t').Append(Invariant(cx)).Append('\t').Append(Invariant(cy)).Append('\t').Append(Invariant(radius)).Append('\t').Append(fixture.ClearHole ? "true" : "false").Append("\t#D9C7A3\n");
        using (Bitmap image = new Bitmap(png))
        {
            Check(image.Width == ExportWidth && image.Height == ExportHeight, "PNG dimensions");
            if (fixture.ClearHole)
            {
                double inner = fixture.Spec.Pie.HoleRatio ?? .5;
                for (int i = 0; i < 9; i++)
                {
                    double angle = i * Math.PI / 4, distance = i == 0 ? 0 : radius * inner * .4;
                    Color color = image.GetPixel((int)Math.Round(cx + Math.Cos(angle) * distance), (int)Math.Round(cy + Math.Sin(angle) * distance));
                    Check(ColorDistance(color, Background) <= 5, "true transparent hole shows distinctive slide background: " + fixture.Name);
                }
            }
            else Check(ColorDistance(image.GetPixel((int)Math.Round(cx), (int)Math.Round(cy)), Background) > 30, "full pie has filled center");
            Color filled = image.GetPixel((int)Math.Round(cx + radius * .8 * Math.Cos(.3)), (int)Math.Round(cy + radius * .8 * Math.Sin(.3)));
            Check(ColorDistance(filled, Background) > 30, "slice body remains filled around hole");
        }
    }
    private static int ColorDistance(Color a, Color b) { return Math.Abs(a.R - b.R) + Math.Abs(a.G - b.G) + Math.Abs(a.B - b.B); }

    private static void CheckPackage(string filename, List<Fixture> fixtures)
    {
        using (Package package = Package.Open(filename, FileMode.Open, FileAccess.Read))
        {
            for (int i = 0; i < fixtures.Count; i++)
            {
                Fixture fixture = fixtures[i]; XmlDocument xml = new XmlDocument();
                using (Stream stream = package.GetPart(new Uri("/ppt/slides/slide" + (i + 1) + ".xml", UriKind.Relative)).GetStream(FileMode.Open, FileAccess.Read)) xml.Load(stream);
                XmlNamespaceManager ns = new XmlNamespaceManager(xml.NameTable); ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main"); ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
                Check(xml.SelectNodes("//p:pic", ns).Count == 0, "XML contains no raster slice image");
                Check(xml.SelectNodes("//a:custGeom", ns).Count == fixture.PathNames.Count, "XML custom geometry count matches native slices");
                foreach (string name in fixture.PathNames)
                {
                    XmlNode shape = null;
                    foreach (XmlNode candidate in xml.SelectNodes("//p:sp", ns))
                    { XmlNode identity = candidate.SelectSingleNode("p:nvSpPr/p:cNvPr", ns); if (identity != null && identity.Attributes["name"].Value == name) { shape = candidate; break; } }
                    Check(shape != null, "XML native slice name survives save");
                    XmlNodeList paths = shape.SelectNodes("p:spPr/a:custGeom/a:pathLst/a:path", ns);
                    Check(paths.Count > 0, "slice stores custom paths");
                    foreach (XmlNode path in paths)
                    {
                        Check(path.SelectNodes("a:cubicBezTo", ns).Count > 0, "native slice retains Bezier arcs");
                        List<Point> polygon = XmlPath(path, ns);
                        Check(polygon.Count > 3, "custom contour has geometry");
                        Point first = polygon[0], last = polygon[polygon.Count - 1];
                        Check(path.SelectSingleNode("a:close", ns) != null || Math.Abs(first.X - last.X) <= 2 && Math.Abs(first.Y - last.Y) <= 2, "custom contour explicitly or geometrically closed");
                        if (fixture.FullCircle && fixture.ClearHole)
                        {
                            double[] bounds = Bounds(polygon); Point center = new Point((bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2);
                            Check(Winding(polygon, center) == 0, "full annulus XML has an unfilled center by contour winding");
                            Point body = new Point(center.X + (bounds[2] - bounds[0]) * .4 * Math.Cos(.3), center.Y + (bounds[3] - bounds[1]) * .4 * Math.Sin(.3));
                            Check(Winding(polygon, body) != 0, "annulus XML contour fills its body");
                        }
                    }
                }
            }
        }
    }
    private static Point XmlPoint(XmlNode node)
    { return new Point(Double.Parse(node.Attributes["x"].Value, CultureInfo.InvariantCulture), Double.Parse(node.Attributes["y"].Value, CultureInfo.InvariantCulture)); }
    private static List<Point> XmlPath(XmlNode path, XmlNamespaceManager ns)
    {
        ShapePath sampled = new ShapePath(); bool started = false;
        foreach (XmlNode command in path.ChildNodes)
        {
            if (command.LocalName == "moveTo")
            {
                Check(!started, "one contour per generated freeform"); Point point = XmlPoint(command.SelectSingleNode("a:pt", ns)); sampled.StartX = point.X; sampled.StartY = point.Y; started = true;
            }
            else if (command.LocalName == "lnTo") { Point point = XmlPoint(command.SelectSingleNode("a:pt", ns)); sampled.Segments.Add(new PathSegment { X = point.X, Y = point.Y }); }
            else if (command.LocalName == "cubicBezTo")
            {
                XmlNodeList points = command.SelectNodes("a:pt", ns); Check(points.Count == 3, "cubic contains two controls and endpoint");
                Point a = XmlPoint(points[0]), b = XmlPoint(points[1]), c = XmlPoint(points[2]);
                sampled.Segments.Add(new PathSegment { Curve = true, C1X = a.X, C1Y = a.Y, C2X = b.X, C2Y = b.Y, X = c.X, Y = c.Y });
            }
            else Check(command.LocalName == "close", "only expected native path commands");
        }
        Check(started, "custom path has starting point"); return Sample(sampled);
    }
    private static int Winding(List<Point> polygon, Point point)
    {
        int winding = 0;
        for (int i = 0; i < polygon.Count; i++)
        {
            Point a = polygon[i], b = polygon[(i + 1) % polygon.Count];
            double cross = (b.X - a.X) * (point.Y - a.Y) - (point.X - a.X) * (b.Y - a.Y);
            if (a.Y <= point.Y) { if (b.Y > point.Y && cross > 0) winding++; }
            else if (b.Y <= point.Y && cross < 0) winding--;
        }
        return winding;
    }
    private static void Release(object value)
    { if (value != null && Marshal.IsComObject(value)) { try { Marshal.FinalReleaseComObject(value); } catch { } } }
}
