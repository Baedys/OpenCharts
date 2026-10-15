using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    public sealed class WaterfallEditor : Form
    {
        private sealed class Choice
        {
            public string Id { get; set; }
            public string Label { get; set; }
        }
        private readonly ChartSpec source;
        private readonly double renderWidth, renderHeight;
        private readonly TextBox title = new TextBox();
        private readonly DataGridView data = Grid(), links = Grid(), pins = Grid();
        private readonly ComboBox profile = new ComboBox(), order = new ComboBox();
        private readonly CheckBox signs = new CheckBox { Text = "Afficher les signes", AutoSize = true };
        private readonly Label status = new Label { Dock = DockStyle.Fill, AutoEllipsis = true };
        public ChartSpec Spec { get; private set; }

        public WaterfallEditor(ChartSpec initial, double? width = null, double? height = null)
        {
            if (initial == null || initial.Waterfall == null) throw new ArgumentException("Une cascade à connecteurs est requise.");
            WaterfallSolver.Solve(initial.Waterfall);
            source = ChartStorage.Clone(initial); renderWidth = width ?? 720; renderHeight = height ?? 400;
            Text = "OpenSlideCharts — Cascade"; StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi; Font = new Font("Segoe UI", 9);
            ClientSize = new Size(1040, 650); MinimumSize = new Size(850, 550);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), RowCount = 5, ColumnCount = 1 };
            foreach (float h in new float[] { 34, 54, 0, 45, 48 }) root.RowStyles.Add(new RowStyle(h == 0 ? SizeType.Percent : SizeType.Absolute, h == 0 ? 100 : h));
            Controls.Add(root);
            title.Text = initial.Title ?? ""; title.Dock = DockStyle.Fill; root.Controls.Add(title, 0, 0);
            root.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Colonnes = catégories ; lignes = segments. Une cellule e par colonne au maximum. Cellule vide = segment absent.\nLes liens relient une base ou un niveau cumulé. Chaque somme indépendante doit avoir un ancrage." }, 0, 1);
            var tabs = new TabControl { Dock = DockStyle.Fill }; root.Controls.Add(tabs, 0, 2);
            var dataPage = new TabPage("Données"); tabs.TabPages.Add(dataPage);
            var dataPanel = PanelWithActions(dataPage, data);
            dataPanel.Controls.Add(Button("Coller matrice", delegate { PasteClipboard(); }));
            dataPanel.Controls.Add(Button("Ajouter colonne", delegate { if (data.Columns.Count >= 81) { Error("80 colonnes au maximum."); return; } AddColumn("Catégorie " + data.Columns.Count, NewId("column")); UpdateChoices(); }));
            dataPanel.Controls.Add(Button("Retirer colonne", delegate {
                if (data.CurrentCell == null || data.CurrentCell.ColumnIndex == 0 || data.Columns.Count <= 2) return;
                string id = (string)data.Columns[data.CurrentCell.ColumnIndex].Tag;
                data.Columns.RemoveAt(data.CurrentCell.ColumnIndex);
                RemoveReferences(links, id, 0, 2); RemoveReferences(pins, id, 0); UpdateChoices();
            }));
            dataPanel.Controls.Add(Button("Ajouter segment", delegate { if (data.Rows.Count >= 8) { Error("Huit lignes de segments au maximum."); return; } data.Rows.Add("Segment " + (data.Rows.Count + 1)); }));
            dataPanel.Controls.Add(Button("Retirer segment", delegate { if (data.CurrentCell != null && data.Rows.Count > 1) data.Rows.RemoveAt(data.CurrentCell.RowIndex); }));
            data.ColumnHeaderMouseDoubleClick += delegate(object sender, DataGridViewCellMouseEventArgs e) {
                if (e.ColumnIndex == 0) return;
                string label = AskName(data.Columns[e.ColumnIndex].HeaderText); if (label != null) { data.Columns[e.ColumnIndex].HeaderText = label; UpdateChoices(); }
            };
            data.KeyDown += delegate(object sender, KeyEventArgs e) { if (e.Control && e.KeyCode == Keys.V) { e.Handled = true; e.SuppressKeyPress = true; PasteClipboard(); } };
            data.EditingControlShowing += delegate(object sender, DataGridViewEditingControlShowingEventArgs e) { e.Control.KeyDown -= EditingKeyDown; e.Control.KeyDown += EditingKeyDown; };
            var linkPage = new TabPage("Connecteurs"); tabs.TabPages.Add(linkPage);
            var linkActions = PanelWithActions(linkPage, links);
            links.Columns.Add(ColumnChoice("Left", "Colonne de départ")); links.Columns.Add(AnchorChoice("LeftAnchor", "Niveau de départ"));
            links.Columns.Add(ColumnChoice("Right", "Colonne d'arrivée")); links.Columns.Add(AnchorChoice("RightAnchor", "Niveau d'arrivée"));
            links.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Constrain", HeaderText = "Égaliser", Width = 75 });
            links.Columns[4].DefaultCellStyle.NullValue = true;
            linkActions.Controls.Add(Button("Ajouter lien", delegate { AddLinkRow(null); }));
            linkActions.Controls.Add(Button("Retirer lien", delegate { RemoveCurrent(links); }));
            linkActions.Controls.Add(new Label { AutoSize = true, Text = "Égaliser décoché : lien libre, oblique si niveaux différents.", Padding = new Padding(5, 8, 0, 0) });
            var pinPage = new TabPage("Ancrages"); tabs.TabPages.Add(pinPage);
            var pinActions = PanelWithActions(pinPage, pins);
            pins.Columns.Add(ColumnChoice("Column", "Colonne")); pins.Columns.Add(AnchorChoice("Anchor", "Niveau")); pins.Columns.Add("Level", "Valeur d'ancrage");
            pinActions.Controls.Add(Button("Ajouter ancrage", delegate { AddPinRow(null); }));
            pinActions.Controls.Add(Button("Retirer ancrage", delegate { RemoveCurrent(pins); }));
            pinActions.Controls.Add(new Label { AutoSize = true, Text = "Après suppression d'un lien, ancrez la nouvelle somme à 0.", Padding = new Padding(5, 8, 0, 0) });
            var options = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            profile.DropDownStyle = order.DropDownStyle = ComboBoxStyle.DropDownList;
            profile.Items.AddRange(new object[] { "Cumul vers droite", "Cumul vers gauche", "Total à gauche", "Fin à zéro" }); profile.SelectedIndex = 0; profile.Width = 145;
            order.Items.AddRange(new object[] { "Ordre des lignes", "Ordre inverse" }); order.SelectedIndex = source.Waterfall.SegmentOrder == "SheetBottomToTop" ? 1 : 0; order.Width = 135;
            signs.Checked = source.Waterfall.ShowSigns;
            options.Controls.Add(profile); options.Controls.Add(Button("Recréer les liens", delegate { Arrange(); })); options.Controls.Add(order); options.Controls.Add(signs);
            options.Controls.Add(Button("Calculer", delegate { try { ShowSolution(ReadSpec()); } catch (Exception e) { Error(e); } }));
            root.Controls.Add(options, 0, 3);
            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); footer.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 205)); footer.Controls.Add(status, 0, 0);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            var apply = Button("Appliquer", delegate { try { var candidate = ReadSpec(); ChartEngine.Build(candidate, renderWidth, renderHeight); Spec = candidate; DialogResult = DialogResult.OK; Close(); } catch (Exception e) { Error(e); } });
            var cancel = Button("Annuler", delegate { DialogResult = DialogResult.Cancel; Close(); }); buttons.Controls.Add(apply); buttons.Controls.Add(cancel); footer.Controls.Add(buttons, 1, 0);
            root.Controls.Add(footer, 0, 4); AcceptButton = apply; CancelButton = cancel;
            LoadData(source.Waterfall); UpdateChoices(); LoadConstraints(source.Waterfall);
        }
        private static DataGridView Grid()
        {
            var grid = new DataGridView { Dock = DockStyle.Fill, BackgroundColor = Color.White, AllowUserToAddRows = false, AllowUserToDeleteRows = false, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill, RowHeadersWidth = 30 };
            grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            return grid;
        }
        private static FlowLayoutPanel PanelWithActions(TabPage page, DataGridView grid)
        {
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            panel.Controls.Add(grid, 0, 0); var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, AutoScroll = true }; panel.Controls.Add(actions, 0, 1); page.Controls.Add(panel); return actions;
        }
        private static Button Button(string text, EventHandler action) { var b = new Button { Text = text, AutoSize = true }; b.Click += action; return b; }
        private static string NewId(string prefix) { return prefix + "_" + Guid.NewGuid().ToString("N"); }
        private static DataGridViewComboBoxColumn ColumnChoice(string name, string label)
        { return new DataGridViewComboBoxColumn { Name = name, HeaderText = label, DisplayMember = "Label", ValueMember = "Id", MinimumWidth = 130 }; }
        private static DataGridViewComboBoxColumn AnchorChoice(string name, string label)
        { var c = new DataGridViewComboBoxColumn { Name = name, HeaderText = label }; c.Items.AddRange("Base", "Cumul", "Personnalisé"); return c; }
        private void AddColumn(string label, string id)
        { int i = data.Columns.Add(NewId("data"), label); data.Columns[i].Tag = id; data.Columns[i].MinimumWidth = 85; data.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable; }
        private void LoadData(WaterfallSpec model)
        {
            data.Columns.Clear(); data.Rows.Clear(); data.Columns.Add("Segment", "Segment"); data.Columns[0].ReadOnly = true; data.Columns[0].Frozen = true; data.Columns[0].MinimumWidth = 85;
            int rows = 1; foreach (WaterfallColumn c in model.Columns) { AddColumn(c.Label, c.Id); for (int i = 0; i < c.Segments.Count; i++) rows = Math.Max(rows, (c.Segments[i].SheetRow >= 0 ? c.Segments[i].SheetRow : i) + 1); }
            for (int r = 0; r < rows; r++) data.Rows.Add("Segment " + (r + 1));
            for (int c = 0; c < model.Columns.Count; c++)
            for (int r = 0; r < model.Columns[c].Segments.Count; r++)
            {
                WaterfallSegment s = model.Columns[c].Segments[r]; int sheetRow = s.SheetRow >= 0 ? s.SheetRow : r;
                data.Rows[sheetRow].Cells[c + 1].Value = s.Calculated ? "e" : ChartEditor.DisplayNumber(s.Value.Value); data.Rows[sheetRow].Cells[c + 1].Tag = s.Id;
            }
        }
        private List<Choice> Choices()
        {
            var choices = new List<Choice>(); for (int c = 1; c < data.Columns.Count; c++) choices.Add(new Choice { Id = (string)data.Columns[c].Tag, Label = c + ": " + data.Columns[c].HeaderText }); return choices;
        }
        private void UpdateChoices()
        { foreach (DataGridView g in new DataGridView[] { links, pins }) foreach (DataGridViewColumn column in g.Columns) if (column is DataGridViewComboBoxColumn && ((DataGridViewComboBoxColumn)column).ValueMember == "Id") ((DataGridViewComboBoxColumn)column).DataSource = Choices(); }
        private static void RemoveCurrent(DataGridView grid) { if (grid.CurrentCell != null) grid.Rows.RemoveAt(grid.CurrentCell.RowIndex); }
        private static void RemoveReferences(DataGridView grid, string id, params int[] columns)
        { for (int r = grid.Rows.Count - 1; r >= 0; r--) foreach (int c in columns) if (Convert.ToString(grid.Rows[r].Cells[c].Value) == id) { grid.Rows.RemoveAt(r); break; } }
        private static bool IsOutput(WaterfallAnchor a)
        { foreach (double w in a.Weights) if (w != 1) return false; return true; }
        private static bool IsBase(WaterfallAnchor a)
        { foreach (double w in a.Weights) if (w != 0) return false; return true; }
        private static void SetAnchor(DataGridViewRow row, int columnCell, int anchorCell, WaterfallAnchor a)
        { row.Cells[columnCell].Value = a.ColumnId; row.Cells[anchorCell].Value = IsBase(a) ? "Base" : IsOutput(a) ? "Cumul" : "Personnalisé"; row.Cells[anchorCell].Tag = a; }
        private void AddLinkRow(WaterfallConnector link)
        {
            if (data.Columns.Count < 2) return;
            int index = links.Rows.Add(); var row = links.Rows[index]; row.Tag = link == null ? NewId("link") : link.Id;
            if (link != null) { SetAnchor(row, 0, 1, link.Left); SetAnchor(row, 2, 3, link.Right); row.Cells[4].Value = link.Constrain; }
            else { row.Cells[0].Value = data.Columns[1].Tag; row.Cells[1].Value = "Cumul"; row.Cells[2].Value = data.Columns[Math.Min(2, data.Columns.Count - 1)].Tag; row.Cells[3].Value = "Base"; row.Cells[4].Value = true; }
        }
        private void AddPinRow(WaterfallPin pin)
        {
            if (data.Columns.Count < 2) return;
            int index = pins.Rows.Add(); var row = pins.Rows[index]; row.Tag = pin == null ? NewId("pin") : pin.Id;
            row.Cells[0].Tag = pin;
            if (pin != null) { SetAnchor(row, 0, 1, pin.Anchor); row.Cells[2].Value = ChartEditor.DisplayNumber(pin.Level); }
            else { row.Cells[0].Value = data.Columns[1].Tag; row.Cells[1].Value = "Base"; row.Cells[2].Value = "0"; }
        }
        private void LoadConstraints(WaterfallSpec model)
        { links.Rows.Clear(); pins.Rows.Clear(); foreach (var c in model.Connectors) AddLinkRow(c); foreach (var p in model.Pins) AddPinRow(p); }
        private WaterfallSpec ReadData()
        {
            data.EndEdit(); var model = new WaterfallSpec { ShowSigns = signs.Checked, SegmentOrder = order.SelectedIndex == 1 ? "SheetBottomToTop" : "SheetTopToBottom" };
            for (int c = 1; c < data.Columns.Count; c++)
            {
                var column = new WaterfallColumn { Id = (string)data.Columns[c].Tag, Label = data.Columns[c].HeaderText };
                for (int r = 0; r < data.Rows.Count; r++)
                {
                    var cell = data.Rows[r].Cells[c]; string raw = Convert.ToString(cell.Value).Trim(); if (raw.Length == 0) continue;
                    if (cell.Tag == null) cell.Tag = NewId("segment"); bool equals = String.Equals(raw, "e", StringComparison.OrdinalIgnoreCase);
                    column.Segments.Add(new WaterfallSegment { Id = (string)cell.Tag, SheetRow = r, Calculated = equals, Value = equals ? (double?)null : ChartEditor.Number(raw, "Colonne " + c + ", segment " + (r + 1)) });
                }
                if (column.Segments.Count == 0) throw new ArgumentException("Colonne " + c + " vide : renseignez une valeur ou retirez la colonne.");
                model.Columns.Add(column);
            }
            return model;
        }
        private WaterfallAnchor ReadAnchor(DataGridViewRow row, int col, int anchor, WaterfallSpec model)
        {
            string id = Convert.ToString(row.Cells[col].Value); WaterfallColumn column = model.Columns.Find(delegate(WaterfallColumn c) { return c.Id == id; });
            if (column == null) throw new ArgumentException("Choisissez une colonne existante pour chaque lien et ancrage.");
            string type = Convert.ToString(row.Cells[anchor].Value);
            if (type == "Base") return WaterfallAnchor.AtBase(column);
            if (type == "Cumul") return WaterfallAnchor.AtOutput(column);
            WaterfallAnchor original = row.Cells[anchor].Tag as WaterfallAnchor;
            if (type != "Personnalisé" || original == null || original.ColumnId != id) throw new ArgumentException("Choisissez Base ou Cumul pour ce niveau.");
            WaterfallColumn old = source.Waterfall.Columns.Find(delegate(WaterfallColumn c) { return c.Id == id; });
            if (old == null || original.Weights.Count != old.Segments.Count) throw new ArgumentException("Ancrage personnalisé incompatible avec les segments.");
            var result = new WaterfallAnchor { ColumnId = id };
            foreach (WaterfallSegment s in column.Segments) { int i = old.Segments.FindIndex(delegate(WaterfallSegment prior) { return prior.Id == s.Id; }); result.Weights.Add(i < 0 ? 0 : original.Weights[i]); }
            for (int i = 0; i < old.Segments.Count; i++) if (original.Weights[i] != 0 && !column.Segments.Exists(delegate(WaterfallSegment s) { return s.Id == old.Segments[i].Id; })) throw new ArgumentException("Un segment d'ancrage a été supprimé. Choisissez Base ou Cumul.");
            return result;
        }
        private ChartSpec ReadSpec()
        {
            links.EndEdit(); pins.EndEdit(); var result = ChartStorage.Clone(source); var model = ReadData(); result.Title = title.Text; result.Waterfall = model;
            foreach (DataGridViewRow row in links.Rows) model.Connectors.Add(new WaterfallConnector { Id = (string)row.Tag, Left = ReadAnchor(row, 0, 1, model), Right = ReadAnchor(row, 2, 3, model), Constrain = row.Cells[4].Value == null || Convert.ToBoolean(row.Cells[4].Value) });
            foreach (DataGridViewRow row in pins.Rows)
            {
                var pin = new WaterfallPin { Id = (string)row.Tag, Anchor = ReadAnchor(row, 0, 1, model), Level = ChartEditor.Number(Convert.ToString(row.Cells[2].Value), "Ancrage") };
                WaterfallPin old = row.Cells[0].Tag as WaterfallPin;
                if (old == null) old = source.Waterfall.Pins.Find(delegate(WaterfallPin p) { return p.Id == pin.Id; });
                bool unchanged = old != null && old.Level == pin.Level && old.Anchor.ColumnId == pin.Anchor.ColumnId && old.Anchor.Weights.Count == pin.Anchor.Weights.Count;
                for (int i = 0; unchanged && i < pin.Anchor.Weights.Count; i++) if (old.Anchor.Weights[i] != pin.Anchor.Weights[i]) unchanged = false;
                pin.AutoGenerated = unchanged && old.AutoGenerated;
                model.Pins.Add(pin);
            }
            return result;
        }
        private void Arrange()
        {
            try { var model = ReadData(); WaterfallProfiles.Arrange(model, new string[] { "Forward", "Backward", "LeftTotal", "EndAtZero" }[profile.SelectedIndex]); WaterfallSolver.Solve(model); LoadConstraints(model); status.Text = "Liens et ancrages recréés. Les données sont conservées."; }
            catch (Exception e) { Error(e); }
        }
        private void ShowSolution(ChartSpec spec)
        {
            var solved = WaterfallSolver.Solve(spec.Waterfall); var values = new List<string>();
            foreach (var c in solved.Columns) foreach (var s in c.Segments) if (s.Calculated) values.Add(c.Label + " : e = " + ChartEditor.DisplayNumber(s.Value));
            status.Text = values.Count == 0 ? "Calcul résolu." : String.Join(" ; ", values.ToArray());
        }
        private void EditingKeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Control || e.KeyCode != Keys.V) return;
            try
            {
                string text = Clipboard.GetText();
                if (text.IndexOf('\t') < 0 && text.TrimEnd('\r', '\n').IndexOf('\n') < 0) return;
                e.Handled = true; e.SuppressKeyPress = true; PasteMatrix(text);
            }
            catch (Exception error) { Error(error); }
        }
        private void PasteClipboard() { try { PasteMatrix(Clipboard.GetText()); } catch (Exception e) { Error(e); } }
        private void PasteMatrix(string text)
        {
            var table = ChartEditor.ParseTsv(text);
            if (table.Count < 2 || table.Count > 9 || table[0].Count < 2 || table[0].Count > 81) throw new ArgumentException("Matrice attendue : en-tête, 1 à 8 lignes de segments et 1 à 80 catégories.");
            for (int r = 1; r < table.Count; r++)
            {
                if (table[r].Count != table[0].Count) throw new ArgumentException("Nombre de colonnes différent à la ligne " + (r + 1));
                for (int c = 1; c < table[r].Count; c++) if (!String.IsNullOrWhiteSpace(table[r][c]) && !String.Equals(table[r][c].Trim(), "e", StringComparison.OrdinalIgnoreCase)) ChartEditor.Number(table[r][c], "Collage");
            }
            var ids = new List<string>(); var cells = new Dictionary<string, string>();
            for (int c = 1; c < data.Columns.Count; c++) { ids.Add((string)data.Columns[c].Tag); for (int r = 0; r < data.Rows.Count; r++) if (data.Rows[r].Cells[c].Tag != null) cells[c + ":" + r] = (string)data.Rows[r].Cells[c].Tag; }
            data.Rows.Clear(); data.Columns.Clear(); data.Columns.Add("Segment", "Segment"); data.Columns[0].ReadOnly = true; data.Columns[0].Frozen = true;
            for (int c = 1; c < table[0].Count; c++) AddColumn(table[0][c], c <= ids.Count ? ids[c - 1] : NewId("column"));
            for (int r = 1; r < table.Count; r++)
            {
                int index = data.Rows.Add(table[r].ToArray());
                for (int c = 1; c < table[r].Count; c++) { string id; if (cells.TryGetValue(c + ":" + (r - 1), out id)) data.Rows[index].Cells[c].Tag = id; }
            }
            for (int i = table[0].Count - 1; i < ids.Count; i++) { RemoveReferences(links, ids[i], 0, 2); RemoveReferences(pins, ids[i], 0); }
            UpdateChoices(); status.Text = "Matrice collée. Recréez les liens si la structure a changé.";
        }
        private string AskName(string initial)
        {
            using (var dialog = new Form { Text = "Nom de catégorie", StartPosition = FormStartPosition.CenterParent, ClientSize = new Size(330, 85), FormBorderStyle = FormBorderStyle.FixedDialog, MinimizeBox = false, MaximizeBox = false })
            {
                var input = new TextBox { Text = initial, Left = 12, Top = 12, Width = 305 }; var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 235, Top = 45 };
                dialog.Controls.Add(input); dialog.Controls.Add(ok); dialog.AcceptButton = ok; return dialog.ShowDialog(this) == DialogResult.OK ? input.Text : null;
            }
        }
        private void Error(Exception error)
        {
            var solve = error as WaterfallSolveException;
            if (solve == null) { Error(error.Message); return; }
            if (solve.Failure == WaterfallSolveFailure.Inconsistent)
                Error("Ancrages incompatibles. Modifiez un lien ou une valeur d'ancrage ; décochez Égaliser pour conserver un connecteur libre.");
            else
                Error("Une somme manque d'ancrages pour calculer ses niveaux ou ses cellules e. Ajoutez un ancrage dans l'onglet Ancrages, ou recréez les liens avec une disposition.");
        }
        private void Error(string text) { MessageBox.Show(this, text, "Cascade : vérifier les contraintes", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }
}
