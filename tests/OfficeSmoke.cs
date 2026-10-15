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
    [STAThread]
    public static int Main(string[] args)
    {
        string output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output);
        dynamic ppt = null, deck = null, excel = null, workbook = null, books = null, sheets = null, sheet = null, cells = null;
        var matrices = new Dictionary<int, string>();
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
            for (int argument = 1; argument + 1 < args.Length; argument++)
            {
                if (args[argument] != "--input-json") continue;
                var imported = new JavaScriptSerializer().Deserialize<ChartSpec>(File.ReadAllText(args[argument + 1]));
                int index = (int)deck.Slides.Count + 1;
                dynamic slide = deck.Slides.Add(index, 12);
                dynamic shape = OfficeRenderer.Render((object)slide, imported, 30, 45, 900, 450);
                matrices[index] = ChartStorage.Serialize(imported);
                Check(ChartStorage.Serialize(ChartStorage.Read((object)shape)) == matrices[index], "Imported matrix metadata mismatch.");
                slide.Export(Path.Combine(output, "ImportedData.png"), "PNG", 1440, 810);
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
                Check(spec.Rows.Count > 0, "Missing data after reopening.");
                Check(Math.Abs((double)group.Width - 900) < 1, "Group bounds changed.");
                if (matrices.ContainsKey(i)) Check(ChartStorage.Serialize(spec) == matrices[i], "Matrix data changed after reopening.");
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
