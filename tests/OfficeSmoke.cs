using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using OpenSlideCharts;

public static class OfficeSmoke
{
    private static int assertions;
    private static void Check(bool condition, string message) { assertions++; if (!condition) throw new Exception(message); }
    private static void CheckTextBounds(object shape, ChartSpec spec, ChartRectangle frame)
    {
        dynamic group = shape;
        Layout layout = ChartEngine.Build(spec, frame.Width, frame.Height);
        for (int i = 0; i < layout.Items.Count; i++)
        {
            Primitive p = layout.Items[i];
            if (p.Kind != "text") continue;
            dynamic text = group.GroupItems.Item(i + 2); // First item is the transparent frame.
            Check(Math.Abs((double)text.Width - Math.Max(1, p.W)) < .02 && Math.Abs((double)text.Height - Math.Max(1, p.H)) < .02,
                "Office resized text box: " + p.Text);
            Check(Math.Abs((double)text.Left - frame.X - p.X) < .02 && Math.Abs((double)text.Top - frame.Y - p.Y) < .02,
                "Office moved text box: " + p.Text);
        }
    }
    [STAThread]
    public static int Main(string[] args)
    {
        string output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output);
        dynamic ppt = null, deck = null, excel = null, workbook = null, books = null, sheets = null, sheet = null, cells = null;
        var matrices = new Dictionary<int, string>();
        var frames = new Dictionary<int, ChartRectangle>();
        try
        {
            new XmlDocument().LoadXml(new Connect().GetCustomUI("Microsoft.PowerPoint.Presentation"));
            ppt = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application", true));
            ppt.Visible = -1;
            deck = ppt.Presentations.Add(-1);
            System.Threading.Thread.Sleep(2000);
            ppt.COMAddIns.Update();
            Console.WriteLine("COM add-ins: " + ppt.COMAddIns.Count);
            if (args.Length < 2 || args[1] != "--renderer-only")
            {
                dynamic addin = ppt.COMAddIns.Item("OpenSlideCharts.Connect");
                if (!(bool)addin.Connect) addin.Connect = true;
                Check((bool)addin.Connect, "COM add-in did not connect.");
                string status = addin.Object.DiagnosticStatus();
                Console.WriteLine(status);
                Check(status.Contains("connected=True"), "OnConnection was not called.");
            }
            else Console.WriteLine("Renderer-only test: add-in loading is NOT tested.");
            deck.PageSetup.SlideWidth = 960f; deck.PageSetup.SlideHeight = 540f;
            string[] kinds = {"Waterfall","Column","Bar","Line","Mekko","Gantt"};
            for (int i = 0; i < kinds.Length; i++)
            {
                dynamic slide = deck.Slides.Add(i + 1, 12);
                ChartSpec spec = ChartSamples.Create(kinds[i]); spec.Title = "OpenSlideCharts · " + kinds[i] + " · données éditables";
                dynamic shape = OfficeRenderer.Render((object)slide, spec, 30, 45, 900, 450);
                Check((int)shape.GroupItems.Count > 3, "Chart is not a native editable group.");
                Check(ChartStorage.Serialize(ChartStorage.Read((object)shape)) == ChartStorage.Serialize(spec), "Metadata roundtrip failed for " + kinds[i]);
                slide.Export(Path.Combine(output, kinds[i] + ".png"), "PNG", 1440, 810);
            }
            string xlsx = Path.Combine(output, "example-data.xlsx");
            excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
            excel.Visible = false; excel.DisplayAlerts = false; excel.EnableEvents = false;
            books = excel.Workbooks; workbook = books.Add();
            sheets = workbook.Worksheets; sheet = sheets.Item(1); sheet.Name = "Donnees";
            object[,] data = new object[,] {{"Départ",100,0,true},{"Croissance",40,0,false},{"Coûts",-25,0,false},{"Total",115,0,true}};
            cells = sheet.Range["A1:D4"]; cells.Value2 = data;
            Marshal.FinalReleaseComObject((object)cells); cells = sheet.Range["F1:G3"];
            cells.Value2 = new object[,] { { "North", 10 }, { "South", 20 }, { null, null } };
            Marshal.FinalReleaseComObject((object)cells); cells = sheet.Range["I1:J40"];
            object[,] tallData = new object[40, 2];
            for (int i = 0; i < 40; i++) { tallData[i, 0] = "C" + i; tallData[i, 1] = i == 0 ? 1000 : .1; }
            cells.Value2 = tallData;
            Marshal.FinalReleaseComObject((object)cells); cells = sheet.Range["L1:R4"];
            cells.Value2 = new object[,] { { "A", 1, 2, 4, "North", "#123456", "Circle" }, { "B", 3, 5, 1, "North", null, null }, { "Scale", null, null, 100, null, null, null }, { "Missing", 4, null, 1, "South", null, null } };
            workbook.SaveAs(xlsx, 51);
            Marshal.FinalReleaseComObject((object)cells); cells = null;
            Marshal.FinalReleaseComObject((object)sheet); sheet = null;
            Marshal.FinalReleaseComObject((object)sheets); sheets = null;
            workbook.Close(false); Marshal.FinalReleaseComObject((object)workbook); workbook = null;
            Marshal.FinalReleaseComObject((object)books); books = null;
            excel.Quit(); Marshal.FinalReleaseComObject((object)excel); excel = null;
            ChartSpec linked = ChartSamples.Create("Waterfall");
            linked.Title = "Lien Excel · actualisation manuelle"; linked.WorkbookPath = xlsx; linked.Worksheet = "Donnees"; linked.RangeAddress = "A1:D4";
            linked = ExcelLink.Read(linked);
            Check(linked.Rows.Count == 4 && linked.Rows[0].Label == "Départ" && linked.Rows[2].Value == -25 && linked.Rows[3].Total, "Excel import failed.");
            ChartSpec pieLink = ChartSamples.Create("Pie"); pieLink.WorkbookPath = xlsx; pieLink.Worksheet = "Donnees"; pieLink.RangeAddress = "F1:G3";
            pieLink.Pie = new PieOptions { StartAngle = 37, HoleRatio = .4 };
            pieLink.Pie.Slices.Add(new PieSliceStyle { Index = 2, Color = "#123456", Explode = .2 });
            string pieBefore = ChartStorage.Serialize(pieLink);
            ChartSpec pieImported = ExcelLink.Read(pieLink);
            Check(pieImported.Rows.Count == 2 && pieImported.Pie.Slices.Count == 0 && pieImported.ImportWarnings.Count == 1, "Changed pie row count must reset positional styles with a warning.");
            Check(pieImported.Pie.StartAngle == 37 && pieImported.Pie.HoleRatio == .4 && ChartStorage.Serialize(pieLink) == pieBefore, "Excel refresh must preserve chart settings without mutating source.");
            pieImported.Pie.Slices.Add(new PieSliceStyle { Index = 1, Color = "#123456", Explode = .2 });
            ChartSpec sameCount = ExcelLink.Read(pieImported);
            Check(sameCount.Pie.Slices.Count == 1 && sameCount.Pie.Slices[0].Index == 1 && sameCount.ImportWarnings.Count == 1, "Unchanged pie row count must retain styles and avoid duplicate warnings.");
            pieLink.RangeAddress = "I1:J40"; pieLink.Pie.Slices.Clear();
            ChartSpec tallPie = ExcelLink.Read(pieLink);
            Check(tallPie.Rows.Count == 40 && ChartEngine.Build(tallPie, 720, 1000).Items.Count > 40, "Excel must validate pie data independently of a fixed 720 by 400 layout.");
            ChartSpec scatterLink = ChartSamples.Create("Scatter"); scatterLink.WorkbookPath = xlsx; scatterLink.Worksheet = "Donnees"; scatterLink.RangeAddress = "L1:R4";
            string scatterBefore = ChartStorage.Serialize(scatterLink);
            ChartSpec bubbles = ExcelLink.Read(scatterLink);
            Check(bubbles.Kind == "Bubble" && bubbles.Scatter.Points.Count == 4 && bubbles.Scatter.Points[2].Size == 100 && !bubbles.Scatter.Points[2].X.HasValue, "Excel Size column must create bubbles and retain size-only reference rows.");
            Check(bubbles.Scatter.Points[0].Color == "#123456" && bubbles.Rows.Count == 0 && ChartStorage.Serialize(scatterLink) == scatterBefore, "Scatter Excel parsing must preserve styles without mutating its source.");
            bubbles.RangeAddress = "L1:N4";
            ChartSpec scatterAgain = ExcelLink.Read(bubbles);
            Check(scatterAgain.Kind == "Scatter" && scatterAgain.Scatter.Points.Count == 4 && !scatterAgain.Scatter.Points[0].Size.HasValue, "Removing Excel Size values must return to Scatter.");
            Check(scatterAgain.Scatter.Points[0].Color == "#123456" && scatterAgain.Scatter.Points[0].Marker == "Circle", "Same-count refresh must preserve point styles absent from the source columns.");
            dynamic linkedSlide = deck.Slides.Add(7, 12);
            dynamic old = OfficeRenderer.Render((object)linkedSlide, linked, 30, 45, 900, 450);
            linked.Rows[1].Value = 50; linked.Rows[3].Value = 125;
            dynamic fresh = OfficeRenderer.Replace((object)linkedSlide, (object)old, linked);
            Check((int)linkedSlide.Shapes.Count == 1 && ChartStorage.Read((object)fresh).Rows[1].Value == 50, "Atomic replacement failed.");
            ChartSpec invalid = ChartStorage.Clone(linked); invalid.Rows[0].Value = double.NaN;
            bool rejected = false;
            try { OfficeRenderer.Replace((object)linkedSlide, (object)fresh, invalid); } catch (ArgumentException) { rejected = true; }
            Check(rejected && (int)linkedSlide.Shapes.Count == 1, "Invalid replacement destroyed chart.");
            fresh = OfficeRenderer.Replace((object)linkedSlide, (object)fresh, ExcelLink.Read(linked));
            Check(ChartStorage.Read((object)fresh).Rows[1].Value == 40, "Refresh failed to restore workbook value.");
            linkedSlide.Export(Path.Combine(output, "ExcelLink.png"), "PNG", 1440, 810);
            foreach (string kind in new string[] { "Column", "Bar", "Line" })
            foreach (string mode in new string[] { "Clustered", "Stacked", "Percent" })
            {
                if (kind == "Line" && mode != "Clustered") continue;
                var spec = new ChartSpec { Kind = kind, LayoutMode = mode, Title = kind + " · " + mode };
                spec.Series.AddRange(new string[] { "Produit A", "Produit B", "Produit C" });
                spec.Rows.Add(new ChartRow { Label = "2024", Values = new List<double?> { 40, 25, mode == "Percent" ? 15 : -15 } });
                spec.Rows.Add(new ChartRow { Label = "2025", Values = new List<double?> { 45, null, 20 } });
                spec.Rows.Add(new ChartRow { Label = "2026", Values = new List<double?> { 50, 40, 30 } });
                int index = (int)deck.Slides.Count + 1;
                dynamic slide = deck.Slides.Add(index, 12);
                dynamic shape = OfficeRenderer.Render((object)slide, spec, 30, 45, 900, 450);
                matrices[index] = ChartStorage.Serialize(spec);
                Check(ChartStorage.Serialize(ChartStorage.Read((object)shape)) == matrices[index], "Matrix metadata changed on render.");
                slide.Export(Path.Combine(output, kind + "-" + mode + ".png"), "PNG", 1440, 810);
            }
            var styled = new ChartSpec { Kind = "Column", LayoutMode = "Stacked", Title = "Synthetic appearance" };
            styled.Series.AddRange(new string[] { "North", "South" });
            styled.Rows.Add(new ChartRow { Label = "Q1", Values = new List<double?> { 4, 6 } });
            styled.Rows.Add(new ChartRow { Label = "Q2", Values = new List<double?> { 8, 2 } });
            styled.Appearance = new ChartAppearance {
                SourceFrame = new ChartRectangle { X = 80, Y = 48, Width = 220, Height = 104 },
                PlotRect = new ChartRectangle { X = 40.0 / 220, Y = 12.0 / 104, Width = 160.0 / 220, Height = 80.0 / 104 },
                ShowTitle = false, ShowLegend = false, ShowValueAxis = false, ShowGridlines = false,
                ShowTotals = true, SeriesLabelsAtFirstStack = true, Minimum = 0, Maximum = 10,
                SeriesColors = new List<string> { "#445566", "#112233" }, DisplayOrder = new List<int> { 1, 0 },
                ValueFontSize = 10, CategoryFontSize = 10, SeriesFontSize = 10, TotalFontSize = 10,
                CategoryLabelGap = 2, TotalLabelGap = 2, SeriesLabelGap = 23.3333333333,
                SeriesLabelHeight = 10, CategoryLabelHeight = 10, TotalLabelHeight = 10, GapRatio = .5
            };
            var cascade = WaterfallProfiles.Sample();
            var leftTotal = ChartStorage.Clone(cascade);
            leftTotal.Waterfall.Columns[0].Segments[0].Value = null; leftTotal.Waterfall.Columns[0].Segments[0].Calculated = true;
            leftTotal.Waterfall.Columns[3].Segments[0].Value = 15; leftTotal.Waterfall.Columns[3].Segments[0].Calculated = false;
            leftTotal.Waterfall.Columns[5].Segments[0].Value = 4; leftTotal.Waterfall.Columns[5].Segments[0].Calculated = false;
            leftTotal.Title = "Total à gauche"; WaterfallProfiles.Arrange(leftTotal.Waterfall, "LeftTotal");
            var mixed = new ChartSpec { Kind = "Waterfall", Title = "Signes mixtes", Waterfall = new WaterfallSpec { ShowSigns = true } };
            for (int c = 0; c < 3; c++)
            {
                var column = new WaterfallColumn { Id = "mixed_" + c, Label = "Catégorie " + c };
                column.Segments.Add(new WaterfallSegment { Id = "mixed_value_" + c, SheetRow = 0, Value = c == 2 ? (double?)null : c == 0 ? 3 : 5, Calculated = c == 2 });
                if (c == 1) column.Segments.Add(new WaterfallSegment { Id = "negative", SheetRow = 1, Value = -2 });
                mixed.Waterfall.Columns.Add(column);
            }
            WaterfallProfiles.Arrange(mixed.Waterfall, "Forward");
            mixed.Waterfall.Connectors.Add(new WaterfallConnector { Id = "skew", Left = WaterfallAnchor.AtBase(mixed.Waterfall.Columns[0]), Right = WaterfallAnchor.AtOutput(mixed.Waterfall.Columns[2]), Constrain = false });
            var imports = new List<ChartSpec> { styled, cascade, leftTotal, mixed };
            for (int argument = 1; argument + 1 < args.Length; argument++)
            {
                if (args[argument] != "--input-json") continue;
                imports.Add(new JavaScriptSerializer().Deserialize<ChartSpec>(File.ReadAllText(args[argument + 1])));
            }
            foreach (ChartSpec imported in imports)
            {
                int index = (int)deck.Slides.Count + 1;
                dynamic slide = deck.Slides.Add(index, 12);
                ChartRectangle frame = OfficeRenderer.InsertionFrame(imported, 960, 540);
                frames[index] = frame;
                dynamic shape = OfficeRenderer.Render((object)slide, imported, frame.X, frame.Y, frame.Width, frame.Height);
                matrices[index] = ChartStorage.Serialize(imported);
                Check(ChartStorage.Serialize(ChartStorage.Read((object)shape)) == matrices[index], "Imported matrix metadata mismatch.");
                CheckTextBounds((object)shape, imported, frame);
                if (imported == cascade)
                {
                    imported.Waterfall.Columns[1].Segments[0].Value = 40;
                    shape = OfficeRenderer.Replace((object)slide, (object)shape, imported);
                    Check(WaterfallSolver.Solve(ChartStorage.Read((object)shape).Waterfall).Columns[5].Output == 106, "Waterfall e was not recalculated after replacement.");
                    var conflicting = ChartStorage.Clone(imported);
                    conflicting.Waterfall.Pins.Add(WaterfallProfiles.Pin(WaterfallAnchor.AtBase(conflicting.Waterfall.Columns[0]), 1));
                    bool conflict = false;
                    try { OfficeRenderer.Replace((object)slide, (object)shape, conflicting); } catch (WaterfallSolveException) { conflict = true; }
                    Check(conflict && (int)slide.Shapes.Count == 1, "Conflicting waterfall replacement destroyed chart.");
                    matrices[index] = ChartStorage.Serialize(imported);
                }
                slide.Export(Path.Combine(output, "ImportedData-" + index + ".png"), "PNG", 1440, 810);
            }
            int slideCount = (int)deck.Slides.Count;
            string pptx = Path.Combine(output, "OpenSlideCharts-demo.pptx");
            deck.SaveAs(pptx, 24); deck.Close(); Marshal.FinalReleaseComObject((object)deck); deck = null;
            deck = ppt.Presentations.Open(pptx, 0, 0, 0);
            Check((int)deck.Slides.Count == slideCount, "Slide count after reopen.");
            for (int i = 1; i <= slideCount; i++)
            {
                dynamic slide = deck.Slides.Item(i); dynamic group = slide.Shapes.Item(1);
                ChartSpec spec = ChartStorage.Read((object)group);
                Check(spec.Waterfall != null ? spec.Waterfall.Columns.Count > 0 : spec.Rows.Count > 0, "Missing data after reopening.");
                Check(Math.Abs((double)group.Width - (frames.ContainsKey(i) ? frames[i].Width : 900)) < .02, "Group bounds changed.");
                if (matrices.ContainsKey(i))
                {
                    // JSON object member order is not data. Reflection-backed serializers
                    // may enumerate fields in a different order after importing a model.
                    var expected = new JavaScriptSerializer().Deserialize<ChartSpec>(matrices[i]);
                    Check(ChartStorage.Serialize(spec) == ChartStorage.Serialize(expected), "Matrix data changed after reopening slide " + i + ".");
                }
                if (frames.ContainsKey(i)) CheckTextBounds((object)group, spec, frames[i]);
            }
            Console.WriteLine("PASS: " + assertions + " Office integration assertions; native charts, metadata, save/reopen, Excel import/refresh, rollback.");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex.ToString()); return 1; }
        finally
        {
            if ((object)workbook != null) { try { workbook.Close(false); } catch { } }
            if ((object)excel != null) { try { excel.Quit(); } catch { } }
            if ((object)deck != null) { try { deck.Close(); } catch { } }
            if ((object)ppt != null) { try { if ((int)ppt.Presentations.Count == 0) ppt.Quit(); } catch { } }
            if ((object)deck != null) { try { Marshal.FinalReleaseComObject((object)deck); } catch { } }
            if ((object)ppt != null) { try { Marshal.FinalReleaseComObject((object)ppt); } catch { } }
        }
    }
}
