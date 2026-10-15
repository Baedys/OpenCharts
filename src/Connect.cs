using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Web.Script.Serialization;

[assembly: AssemblyTitle("OpenSlideCharts")]
[assembly: AssemblyDescription("Open source editable charts for PowerPoint")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: ComVisible(false)]

namespace OpenSlideCharts
{
    [ComImport, ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode, [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int removeMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }
    [ComImport, ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }
    [ComVisible(true), Guid("F42D7316-0B29-4DB1-A11D-6D889F02DDBA"), ProgId("OpenSlideCharts.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private dynamic app;
        private bool ribbonRequested;
        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            app = application;
            try { ((dynamic)addInInst).Object = this; } catch { }
        }
        public void OnDisconnection(int removeMode, ref Array custom) { app = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }
        public string DiagnosticStatus() { return "OpenSlideCharts 0.1; connected=" + (app != null) + "; ribbon=" + ribbonRequested; }
        public string GetCustomUI(string ribbonId)
        {
            ribbonRequested = true;
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='oscTab' label='OpenSlideCharts'>
              <group id='oscInsert' label='Graphiques éditables'>
                <button id='oscWaterfall' label='Cascade' size='large' tag='Waterfall' onAction='InsertChart'/>
                <button id='oscColumn' label='Colonnes' size='large' tag='Column' onAction='InsertChart'/>
                <button id='oscBar' label='Barres' size='large' tag='Bar' onAction='InsertChart'/>
                <button id='oscLine' label='Courbe' size='large' tag='Line' onAction='InsertChart'/>
                <button id='oscMekko' label='Mekko' size='large' tag='Mekko' onAction='InsertChart'/>
                <button id='oscGantt' label='Gantt' size='large' tag='Gantt' onAction='InsertChart'/>
                <button id='oscSeries' label='Multi-séries' onAction='InsertSeries'/>
              </group>
              <group id='oscData' label='Données et liens Excel'>
                <button id='oscEdit' label='Modifier les données' onAction='EditChart'/>
                <button id='oscImport' label='Importer JSON' onAction='ImportChart'/>
                <button id='oscRefresh' label='Actualiser la sélection' onAction='RefreshChart'/>
                <button id='oscRefreshAll' label='Actualiser la présentation' onAction='RefreshAll'/>
              </group>
              <group id='oscInfo' label='Version 0.1'><button id='oscAbout' label='À propos / limites' onAction='About'/></group>
            </tab></tabs></ribbon></customUI>";
        }
        public void InsertChart(object control)
        {
            Run(delegate
            {
                dynamic slide = ActiveSlide();
                string kind = (string)((dynamic)control).Tag;
                using (ChartEditor editor = new ChartEditor(ChartSamples.Create(kind)))
                {
                    if (editor.ShowDialog() != DialogResult.OK) return;
                    app.StartNewUndoEntry();
                    double w = (double)app.ActivePresentation.PageSetup.SlideWidth;
                    double h = (double)app.ActivePresentation.PageSetup.SlideHeight;
                    dynamic chart = OfficeRenderer.Render((object)slide, editor.Spec, 30, 35, Math.Max(320, w - 60), Math.Max(240, h - 70));
                    chart.Select();
                }
            });
        }
        public void EditChart(object control)
        {
            Run(delegate
            {
                dynamic slide = ActiveSlide(), selected = SelectedChart();
                ChartSpec edited = EditSpec(ChartStorage.Read((object)selected), (double)selected.Width, (double)selected.Height);
                if (edited != null) {
                    app.StartNewUndoEntry();
                    dynamic chart = OfficeRenderer.Replace((object)slide, (object)selected, edited);
                    chart.Select();
                }
            });
        }
        private static ChartSpec EditSpec(ChartSpec source, double? width = null, double? height = null)
        {
            if (source.Series != null && source.Series.Count > 0)
            {
                using (var editor = new SeriesEditor(source, width, height)) return editor.ShowDialog() == DialogResult.OK ? editor.Spec : null;
            }
            using (var editor = new ChartEditor(source)) return editor.ShowDialog() == DialogResult.OK ? editor.Spec : null;
        }
        public void InsertSeries(object control)
        {
            Run(delegate
            {
                var sample = new ChartSpec { Kind = "Column", Title = "Séries", LayoutMode = "Stacked" };
                sample.Series.AddRange(new string[] { "Produit A", "Produit B", "Produit C" });
                sample.Rows.Add(new ChartRow { Label = "2024", Values = new List<double?> { 40, 25, 15 } });
                sample.Rows.Add(new ChartRow { Label = "2025", Values = new List<double?> { 45, 35, 20 } });
                sample.Rows.Add(new ChartRow { Label = "2026", Values = new List<double?> { 50, 40, 30 } });
                InsertSpec(sample);
            });
        }
        public void ImportChart(object control)
        {
            Run(delegate
            {
                using (var dialog = new OpenFileDialog { Title = "Importer des données OpenSlideCharts", Filter = "Graphique JSON|*.json", CheckFileExists = true })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    if (new FileInfo(dialog.FileName).Length > 1048576) throw new InvalidDataException("Fichier JSON limité à 1 Mio.");
                    ChartSpec spec = new JavaScriptSerializer().Deserialize<ChartSpec>(File.ReadAllText(dialog.FileName));
                    InsertSpec(spec);
                }
            });
        }
        private void InsertSpec(ChartSpec initial)
        {
            dynamic slide = ActiveSlide();
            double w = (double)app.ActivePresentation.PageSetup.SlideWidth;
            double h = (double)app.ActivePresentation.PageSetup.SlideHeight;
            ChartRectangle initialFrame = OfficeRenderer.InsertionFrame(initial, w, h);
            ChartEngine.Build(initial, initialFrame.Width, initialFrame.Height);
            ChartSpec spec = EditSpec(initial, initialFrame.Width, initialFrame.Height);
            if (spec == null) return;
            app.StartNewUndoEntry();
            ChartRectangle frame = OfficeRenderer.InsertionFrame(spec, w, h);
            dynamic chart = OfficeRenderer.Render((object)slide, spec, frame.X, frame.Y, frame.Width, frame.Height);
            chart.Select();
        }
        public void RefreshChart(object control)
        {
            Run(delegate
            {
                dynamic slide = ActiveSlide(), selected = SelectedChart();
                ChartSpec updated = ExcelLink.Read(ChartStorage.Read((object)selected));
                app.StartNewUndoEntry();
                dynamic chart = OfficeRenderer.Replace((object)slide, (object)selected, updated);
                chart.Select();
            });
        }
        public void RefreshAll(object control)
        {
            Run(delegate
            {
                if (app == null || app.Presentations.Count == 0) throw new InvalidOperationException("Ouvrez une présentation.");
                int updated = 0; List<string> errors = new List<string>();
                app.StartNewUndoEntry();
                dynamic presentation = app.ActivePresentation;
                for (int i = 1; i <= presentation.Slides.Count; i++)
                {
                    dynamic slide = presentation.Slides.Item(i);
                    List<object> targets = new List<object>();
                    for (int j = 1; j <= slide.Shapes.Count; j++)
                    {
                        dynamic shape = slide.Shapes.Item(j);
                        if ((string)shape.Tags.Item("OSC_VERSION") == "1") targets.Add((object)shape);
                    }
                    foreach (object target in targets)
                    {
                        try
                        {
                            ChartSpec spec = ChartStorage.Read(target);
                            if (String.IsNullOrWhiteSpace(spec.WorkbookPath)) continue;
                            OfficeRenderer.Replace((object)slide, target, ExcelLink.Read(spec)); updated++;
                        }
                        catch (Exception ex) { errors.Add("Diapositive " + i + " : " + ex.Message); }
                    }
                }
                MessageBox.Show(updated + " graphique(s) actualisé(s)." + (errors.Count > 0 ? "\n\n" + String.Join("\n", errors.ToArray()) : ""), "OpenSlideCharts", MessageBoxButtons.OK, errors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            });
        }
        public void About(object control)
        {
            MessageBox.Show("OpenSlideCharts 0.1 — sources sous licence MIT.\n\nPrototype indépendant, non affilié à think-cell.\n6 types de graphiques en formes PowerPoint éditables.\nColonnes, barres et courbes : jusqu'à 8 séries.\nExcel : lecture des données enregistrées, actualisation manuelle.\n\nLimites : Mekko à 2 segments ; Gantt en jours relatifs ; pas d'anticollision des étiquettes ni d'édition des objets think-cell.\nLa réédition régénère les formes : personnalisations manuelles et animations ne sont pas conservées.\nLa parité think-cell reste à réaliser.", "OpenSlideCharts", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private object ActiveSlide()
        {
            if (app == null || app.Presentations.Count == 0 || app.ActiveWindow == null) throw new InvalidOperationException("Ouvrez une présentation et sélectionnez une diapositive en mode Normal.");
            try { return (object)app.ActiveWindow.View.Slide; }
            catch { throw new InvalidOperationException("Sélectionnez une diapositive en mode Normal."); }
        }
        private object SelectedChart()
        {
            dynamic selection = app.ActiveWindow.Selection;
            if ((int)selection.Type != 2 || (int)selection.ShapeRange.Count != 1) throw new InvalidOperationException("Sélectionnez un seul groupe OpenSlideCharts.");
            dynamic shape = selection.ShapeRange.Item(1);
            if ((string)shape.Tags.Item("OSC_VERSION") != "1") throw new InvalidOperationException("Sélectionnez le groupe complet OpenSlideCharts.");
            return shape;
        }
        private static void Run(Action action)
        {
            try { action(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "OpenSlideCharts", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
    }
}
