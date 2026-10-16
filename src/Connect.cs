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
    public sealed class Connect : StandardOleMarshalObject, IDTExtensibility2, IRibbonExtensibility
    {
        private dynamic app;
        private bool ribbonRequested;
        private dynamic ribbon;
        private PowerPointEvents events;
        private WaterfallSlideEditor slideEditor;
        private string interactionError;
        private bool disconnectPending;
        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            app = application;
            try { ((dynamic)addInInst).Object = this; } catch { }
            try
            {
                slideEditor = new WaterfallSlideEditor(application, delegate { EditChart(null); });
                events = new PowerPointEvents(application, SelectionChanged, null,
                    BeforePresentationSave, BeforePresentationSave, delegate(object presentation, object window) { slideEditor.SafeStop(); InvalidateRibbon(); });
            }
            catch (Exception error) { interactionError = error.Message; if (slideEditor != null) { slideEditor.Dispose(); slideEditor = null; } }
        }
        public void OnDisconnection(int removeMode, ref Array custom) { disconnectPending = true; DisconnectInteractions(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { disconnectPending = true; DisconnectInteractions(); }
        private void DisconnectInteractions()
        {
            try
            {
                if (slideEditor != null) slideEditor.Dispose();
                if (events != null) events.Dispose();
                events = null; slideEditor = null; ribbon = null; app = null; disconnectPending = false;
            }
            catch (Exception error) { interactionError = error.Message; }
        }
        private void BeforePresentationSave(object presentation, ref bool cancel)
        {
            try { if (slideEditor != null) slideEditor.Stop(true); InvalidateRibbon(); if (disconnectPending) DisconnectInteractions(); }
            catch (Exception error) { cancel = true; interactionError = error.Message; MessageBox.Show("L'édition des connecteurs n'a pas pu être terminée. Réessayez après avoir terminé l'édition.\n" + error.Message, "OpenSlideCharts", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
        public string DiagnosticStatus() { return "OpenSlideCharts 0.1; connected=" + (app != null) + "; ribbon=" + ribbonRequested + "; events=" + (events != null) + "; connector-editing=" + (slideEditor != null && slideEditor.Active) + "; connector-adding=" + (slideEditor != null && slideEditor.AddingConnector) + "; interface-handles=" + (slideEditor != null && slideEditor.InterfaceHandlesVisible) + "; edits=" + (slideEditor == null ? 0 : slideEditor.AppliedEdits) + "; polls=" + (slideEditor == null ? 0 : slideEditor.PollCount) + "; interaction=" + (interactionError ?? (slideEditor == null ? "" : slideEditor.LastMessage)); }
        public string GetCustomUI(string ribbonId)
        {
            ribbonRequested = true;
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'><ribbon><tabs><tab id='oscTab' label='OpenSlideCharts'>
              <group id='oscInsert' label='Graphiques éditables'>
                <button id='oscWaterfall' label='Cascade' size='large' tag='Waterfall' onAction='InsertChart'/>
                <button id='oscColumn' label='Colonnes' size='large' tag='Column' onAction='InsertChart'/>
                <button id='oscBar' label='Barres' size='large' tag='Bar' onAction='InsertChart'/>
                <button id='oscLine' label='Courbe' size='large' tag='Line' onAction='InsertChart'/>
                <button id='oscMekko' label='Mekko' size='large' tag='Mekko' onAction='InsertChart'/>
                <button id='oscGantt' label='Gantt' size='large' tag='Gantt' onAction='InsertChart'/>
                <button id='oscPie' label='Secteurs' tag='Pie' onAction='InsertChart'/>
                <button id='oscDoughnut' label='Anneau' tag='Doughnut' onAction='InsertChart'/>
                <button id='oscScatter' label='Dispersion' tag='Scatter' onAction='InsertChart'/>
                <button id='oscBubble' label='Bulles' tag='Bubble' onAction='InsertChart'/>
                <button id='oscSeries' label='Multi-séries' onAction='InsertSeries'/>
              </group>
              <group id='oscData' label='Données et liens Excel'>
                <button id='oscEdit' label='Modifier les données' onAction='EditChart'/>
                <button id='oscImport' label='Importer JSON' onAction='ImportChart'/>
                <button id='oscConnectors' label='Éditer les connecteurs' onAction='EditConnectors'/>
                <button id='oscApplyConnector' label='Appliquer le déplacement' onAction='ApplyConnectorMove'/>
                <button id='oscEndConnectors' label='Terminer les connecteurs' onAction='EndConnectorEditing'/>
                <button id='oscRefresh' label='Actualiser la sélection' onAction='RefreshChart'/>
                <button id='oscRefreshAll' label='Actualiser la présentation' onAction='RefreshAll'/>
              </group>
              <group id='oscInfo' label='Version 0.1'><button id='oscAbout' label='À propos / limites' onAction='About'/></group>
            </tab></tabs></ribbon>" + ContextMenus() + "</customUI>";
        }
        public void OnRibbonLoad(object ribbonUi) { ribbon = ribbonUi; }
        private void SelectionChanged(object selection)
        {
            if (slideEditor != null) slideEditor.SelectionChanged(selection);
            InvalidateRibbon();
        }
        private void InvalidateRibbon()
        {
            if (ribbon != null) try { ribbon.Invalidate(); } catch (COMException) { }
        }
        private static string ContextMenus()
        {
            // Official Office Fluent IDs; context menus also cover keyboard invocation.
            string[] ids = { "ContextMenuObjectsGroup", "ContextMenuShape", "ContextMenuShapeConnector" };
            var xml = new System.Text.StringBuilder("<contextMenus>");
            for (int i = 0; i < ids.Length; i++)
            {
                string suffix = i.ToString(System.Globalization.CultureInfo.InvariantCulture);
                xml.Append("<contextMenu idMso='" + ids[i] + "'><menu id='oscContext" + suffix + "' label='OpenCharts' getVisible='GetChartContextVisible'>");
                xml.Append("<button id='oscContextData" + suffix + "' label='Modifier les données' onAction='EditChart'/>");
                xml.Append("<button id='oscContextLinks" + suffix + "' label='Modifier les connecteurs' getVisible='GetWaterfallContextVisible' onAction='EditConnectors'/>");
                xml.Append("<button id='oscContextDelete" + suffix + "' label='Supprimer ce connecteur' getVisible='GetConnectorContextVisible' onAction='DeleteConnector'/>");
                xml.Append("<menu id='oscContextCreate" + suffix + "' label='Ajouter un connecteur depuis…' getVisible='GetSegmentContextVisible'>");
                xml.Append("<button id='oscContextCreateTop" + suffix + "' label='Sommet du segment' onAction='BeginConnectorFromTop'/>");
                xml.Append("<button id='oscContextCreateBottom" + suffix + "' label='Pied du segment' onAction='BeginConnectorFromBottom'/>");
                xml.Append("<button id='oscContextCreateOutput" + suffix + "' label='Cumul de la colonne' onAction='BeginConnectorFromOutput'/></menu>");
                xml.Append("<button id='oscContextZeroTop" + suffix + "' label='Placer le sommet à zéro' getVisible='GetSegmentContextVisible' onAction='PinSegmentTopZero'/>");
                xml.Append("<button id='oscContextZeroBottom" + suffix + "' label='Placer le pied à zéro' getVisible='GetSegmentContextVisible' onAction='PinSegmentBottomZero'/>");
                xml.Append("<button id='oscContextCancelCreate" + suffix + "' label='Annuler l’ajout du connecteur' getVisible='GetPendingConnectorVisible' onAction='CancelConnectorCreation'/>");
                xml.Append("<button id='oscContextFinish" + suffix + "' label='Terminer les connecteurs' getVisible='GetFinishContextVisible' onAction='EndConnectorEditing'/>");
                xml.Append("</menu></contextMenu>");
            }
            return xml.Append("</contextMenus>").ToString();
        }
        private OfficeChartSelection ContextSelection()
        {
            try { return app == null ? null : OfficeChartSelection.Resolve((object)app.ActiveWindow.Selection); }
            catch { return null; }
        }
        public bool GetChartContextVisible(object control) { return ContextSelection() != null; }
        public bool GetWaterfallContextVisible(object control)
        { var selected = ContextSelection(); return events != null && selected != null && selected.Spec.Waterfall != null; }
        public bool GetConnectorContextVisible(object control)
        { var selected = ContextSelection(); return events != null && selected != null && selected.Spec.Waterfall != null && selected.Primitive != null && selected.Primitive.Role == "WATERFALL_CONNECTOR"; }
        public bool GetFinishContextVisible(object control)
        { return slideEditor != null && slideEditor.Active && GetWaterfallContextVisible(control); }
        public bool GetSegmentContextVisible(object control)
        { var selected = ContextSelection(); return events != null && selected != null && selected.Spec.Waterfall != null && selected.Primitive != null && selected.Primitive.Role == "WATERFALL_SEGMENT" && (slideEditor == null || !slideEditor.AddingConnector); }
        public bool GetPendingConnectorVisible(object control)
        { return slideEditor != null && slideEditor.AddingConnector && GetWaterfallContextVisible(control); }
        public void BeginConnectorFromTop(object control) { SegmentCommand("Top", false); }
        public void BeginConnectorFromBottom(object control) { SegmentCommand("Bottom", false); }
        public void BeginConnectorFromOutput(object control) { SegmentCommand("Output", false); }
        public void PinSegmentTopZero(object control) { SegmentCommand("Top", true); }
        public void PinSegmentBottomZero(object control) { SegmentCommand("Bottom", true); }
        public void CompleteConnectorAtTop(object control) { CompleteConnector("Top", true); }
        public void CompleteConnectorAtBottom(object control) { CompleteConnector("Bottom", true); }
        public void CompleteConnectorAtOutput(object control) { CompleteConnector("Output", true); }
        public void CompleteFreeConnectorAtTop(object control) { CompleteConnector("Top", false); }
        public void CompleteFreeConnectorAtBottom(object control) { CompleteConnector("Bottom", false); }
        public void CompleteFreeConnectorAtOutput(object control) { CompleteConnector("Output", false); }
        public void CancelConnectorCreation(object control)
        { Run(delegate { if (slideEditor != null) slideEditor.CancelConnectorCreation(); InvalidateRibbon(); }); }
        private OfficeChartSelection SelectedSegment()
        {
            OfficeChartSelection selected = ContextSelection();
            if (events == null || slideEditor == null || selected == null || selected.Primitive == null || selected.Primitive.Role != "WATERFALL_SEGMENT")
                throw new InvalidOperationException("Sélectionnez un segment de cascade.");
            return selected;
        }
        private void SegmentCommand(string edge, bool pin)
        {
            Run(delegate
            {
                OfficeChartSelection selected = SelectedSegment(); Primitive item = selected.Primitive;
                if (pin) slideEditor.PinSegmentZero(ActiveSlide(), selected.Chart, item.ColumnId, item.SegmentId, edge);
                else slideEditor.BeginSegmentConnector(ActiveSlide(), selected.Chart, item.ColumnId, item.SegmentId, edge);
                InvalidateRibbon();
            });
        }
        private void CompleteConnector(string edge, bool constrain)
        {
            Run(delegate { OfficeChartSelection selected = SelectedSegment(); Primitive item = selected.Primitive; slideEditor.CompleteSegmentConnector(selected.Chart, item.ColumnId, item.SegmentId, edge, constrain); InvalidateRibbon(); });
        }
        public void DeleteConnector(object control)
        {
            Run(delegate
            {
                OfficeChartSelection selected = ContextSelection();
                if (slideEditor == null || events == null || selected == null || selected.Primitive == null || selected.Primitive.Role != "WATERFALL_CONNECTOR")
                    throw new InvalidOperationException("Sélectionnez un connecteur de cascade.");
                slideEditor.DeleteConnector(ActiveSlide(), selected.Chart, selected.Primitive.EntityId);
                InvalidateRibbon();
            });
        }
        public void InsertChart(object control)
        {
            Run(delegate
            {
                dynamic slide = ActiveSlide();
                string kind = (string)((dynamic)control).Tag;
                if (kind == "Waterfall") { InsertSpec(WaterfallProfiles.Sample()); return; }
                if (kind == "Gantt") { InsertSpec(GanttSamples.Create()); return; }
                if (kind == "Pie" || kind == "Doughnut") { InsertSpec(ChartSamples.Create(kind)); return; }
                if (kind == "Scatter" || kind == "Bubble") { InsertSpec(ChartSamples.Create(kind)); return; }
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
                if (slideEditor != null) slideEditor.Stop(true, true);
                dynamic slide = ActiveSlide(), selected = SelectedChart();
                ChartSpec edited = EditSpec(ChartStorage.Read((object)selected), (double)selected.Width, (double)selected.Height);
                if (edited != null) {
                    app.StartNewUndoEntry();
                    dynamic chart = OfficeRenderer.Replace((object)slide, (object)selected, edited);
                    chart.Select();
                }
            });
        }
        public void EditConnectors(object control)
        {
            Run(delegate
            {
                if (slideEditor == null || events == null) throw new InvalidOperationException("Événements PowerPoint indisponibles : " + interactionError);
                OfficeChartSelection selection = OfficeChartSelection.Resolve((object)app.ActiveWindow.Selection);
                if (selection == null) throw new InvalidOperationException("Sélectionnez une cascade ou un de ses connecteurs.");
                string id = selection.Primitive != null && selection.Primitive.Role == "WATERFALL_CONNECTOR" ? selection.Primitive.EntityId : null;
                slideEditor.Start(ActiveSlide(), selection.Chart, id);
                InvalidateRibbon();
            });
        }
        public void ApplyConnectorMove(object control) { Run(delegate { if (slideEditor != null) slideEditor.ProcessPendingEdits(true); InvalidateRibbon(); }); }
        public void EndConnectorEditing(object control) { Run(delegate { if (slideEditor != null) slideEditor.Stop(true, true); InvalidateRibbon(); }); }
        private static ChartSpec EditSpec(ChartSpec source, double? width = null, double? height = null)
        {
            if (source.Gantt != null)
            {
                using (var editor = new GanttEditor(source, width ?? 900, height ?? 470)) return editor.ShowDialog() == DialogResult.OK ? editor.Spec : null;
            }
            if (String.Equals(source.Kind, "Scatter", StringComparison.OrdinalIgnoreCase) || String.Equals(source.Kind, "Bubble", StringComparison.OrdinalIgnoreCase))
            {
                using (var editor = new ScatterEditor(source, width ?? 720, height ?? 400)) return editor.ShowDialog() == DialogResult.OK ? editor.Spec : null;
            }
            if (String.Equals(source.Kind, "Pie", StringComparison.OrdinalIgnoreCase) || String.Equals(source.Kind, "Doughnut", StringComparison.OrdinalIgnoreCase))
            {
                using (var editor = new PieEditor(source, width ?? 720, height ?? 400)) return editor.ShowDialog() == DialogResult.OK ? editor.Spec : null;
            }
            if (source.Waterfall != null)
            {
                using (var editor = new WaterfallEditor(source, width, height)) return editor.ShowDialog() == DialogResult.OK ? editor.Spec : null;
            }
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
            MessageBox.Show("OpenSlideCharts 0.1 — sources sous licence MIT.\n\nPrototype indépendant, non affilié à think-cell.\n10 types de graphiques en formes PowerPoint éditables.\nColonnes, barres et courbes : jusqu'à 8 séries.\nSecteurs/anneaux : rotation, ordre, trou et éclatement depuis l'éditeur.\nDispersion/bulles : axes, groupes, tailles, partitions et tendances linéaires.\nExcel : lecture des données enregistrées, actualisation manuelle.\n\nLimites : Mekko à 2 segments ; Gantt en jours relatifs ; placement automatique général et édition des objets think-cell absents.\nLa réédition régénère les formes : personnalisations manuelles et animations ne sont pas conservées.\nLa parité think-cell reste à réaliser.", "OpenSlideCharts", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private object ActiveSlide()
        {
            if (app == null || app.Presentations.Count == 0 || app.ActiveWindow == null) throw new InvalidOperationException("Ouvrez une présentation et sélectionnez une diapositive en mode Normal.");
            try { return (object)app.ActiveWindow.View.Slide; }
            catch { throw new InvalidOperationException("Sélectionnez une diapositive en mode Normal."); }
        }
        private object SelectedChart()
        {
            OfficeChartSelection selected = OfficeChartSelection.Resolve((object)app.ActiveWindow.Selection);
            if (selected == null) throw new InvalidOperationException("Sélectionnez un graphique OpenSlideCharts ou une de ses formes.");
            return selected.Chart;
        }
        private static void Run(Action action)
        {
            try { action(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "OpenSlideCharts", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }
    }
}
