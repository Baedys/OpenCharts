using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    // Matrix editor preserves missing observations, category order and series names.
    public sealed class SeriesEditor : Form
    {
        private readonly ComboBox kind = new ComboBox(), mode = new ComboBox();
        private readonly TextBox title = new TextBox();
        private readonly DataGridView grid = new DataGridView();
        private readonly ChartSpec source;
        private readonly double renderWidth, renderHeight;
        public ChartSpec Spec { get; private set; }

        public SeriesEditor(ChartSpec initial, double? width = null, double? height = null)
        {
            source = ChartStorage.Clone(initial);
            ChartRectangle frame = source.Appearance == null ? null : source.Appearance.SourceFrame;
            renderWidth = width ?? (frame == null ? 720 : frame.Width);
            renderHeight = height ?? (frame == null ? 400 : frame.Height);
            Text = "OpenSlideCharts — Séries";
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(860, 560); MinimumSize = new Size(680, 440);
            Font = new Font("Segoe UI", 9);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 1, RowCount = 4 };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            Controls.Add(root);
            var fields = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3 };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            kind.DropDownStyle = mode.DropDownStyle = ComboBoxStyle.DropDownList;
            kind.Items.AddRange(new object[] { "Column", "Bar", "Line" });
            mode.Items.AddRange(new object[] { "Clustered", "Stacked", "Percent" });
            foreach (object option in kind.Items)
                if (String.Equals(option.ToString(), initial.Kind, StringComparison.OrdinalIgnoreCase)) kind.SelectedItem = option;
            if (kind.SelectedIndex < 0) kind.SelectedIndex = 0;
            mode.SelectedItem = initial.LayoutMode ?? "Clustered"; if (mode.SelectedIndex < 0) mode.SelectedIndex = 0;
            kind.Dock = mode.Dock = title.Dock = DockStyle.Fill;
            title.Text = initial.Title ?? "";
            fields.Controls.Add(kind, 0, 0); fields.Controls.Add(mode, 1, 0); fields.Controls.Add(title, 2, 0); root.Controls.Add(fields, 0, 0);
            string instructions = "Cellule vide = donnée absente. Collage : première ligne = noms de séries, première colonne = catégories.\nClustered : groupé ; Stacked : empilé ; Percent : empilé à 100 %. Double-clic sur un en-tête pour renommer.";
            if (source.ImportWarnings != null && source.ImportWarnings.Count > 0)
            {
                root.RowStyles[1].Height = 85;
                instructions += "\nImport partiel : " + String.Join(" ; ", source.ImportWarnings.ToArray());
            }
            root.Controls.Add(new Label { Dock = DockStyle.Fill, Text = instructions }, 0, 1);
            grid.Dock = DockStyle.Fill; grid.BackgroundColor = Color.White;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AllowUserToOrderColumns = false;
            grid.ColumnHeaderMouseDoubleClick += delegate(object sender, DataGridViewCellMouseEventArgs e)
            {
                if (e.ColumnIndex < 1) return;
                string name = AskName(grid.Columns[e.ColumnIndex].HeaderText);
                if (name != null) grid.Columns[e.ColumnIndex].HeaderText = name;
            };
            grid.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Control && e.KeyCode == Keys.V) { e.Handled = true; e.SuppressKeyPress = true; Paste(); }
            };
            root.Controls.Add(grid, 0, 2);
            LoadMatrix(initial);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            actions.Controls.Add(MakeButton("Coller matrice", delegate { Paste(); }));
            actions.Controls.Add(MakeButton("Ajouter série", delegate
            {
                if (grid.Columns.Count >= 9) { Error("Huit séries au maximum."); return; }
                string name = AskName("Série " + grid.Columns.Count);
                if (name != null) AddSeries(name);
            }));
            actions.Controls.Add(MakeButton("Retirer série", delegate
            {
                if (grid.CurrentCell == null || grid.CurrentCell.ColumnIndex == 0) { Error("Sélectionnez une cellule de la série à retirer."); return; }
                if (grid.Columns.Count <= 2) { Error("Conservez au moins une série."); return; }
                grid.Columns.RemoveAt(grid.CurrentCell.ColumnIndex);
            }));
            var apply = MakeButton("Appliquer", delegate
            {
                try { ChartSpec candidate = ReadMatrix(); ValidateCandidate(candidate); Spec = candidate; DialogResult = DialogResult.OK; Close(); }
                catch (Exception ex) { Error(ex.Message); }
            });
            actions.Controls.Add(apply);
            var cancel = MakeButton("Annuler", delegate { DialogResult = DialogResult.Cancel; Close(); });
            actions.Controls.Add(cancel); AcceptButton = apply; CancelButton = cancel;
            root.Controls.Add(actions, 0, 3);
        }
        private static Button MakeButton(string label, EventHandler action)
        { var b = new Button { Text = label, AutoSize = true, Height = 30 }; b.Click += action; return b; }
        private void ValidateCandidate(ChartSpec candidate)
        {
            if ((candidate.Appearance == null || candidate.Appearance.PlotRect == null) && (renderWidth < 320 || renderHeight < 240))
                throw new ArgumentException("Ce changement nécessite un cadre d'au moins 320 × 240 points. Conservez le type actuel, puis agrandissez le groupe dans PowerPoint avant de le modifier.");
            ChartEngine.Build(candidate, renderWidth, renderHeight);
        }
        private void AddSeries(string name, int sourceIndex = -1)
        {
            int index = grid.Columns.Add("S_" + Guid.NewGuid().ToString("N"), name);
            grid.Columns[index].SortMode = DataGridViewColumnSortMode.NotSortable;
            grid.Columns[index].Tag = sourceIndex;
        }
        private void LoadMatrix(ChartSpec spec)
        {
            grid.CancelEdit(); grid.Rows.Clear(); grid.Columns.Clear();
            grid.Columns.Add("Category", "Catégorie"); grid.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            for (int s = 0; s < spec.Series.Count; s++) AddSeries(spec.Series[s], s < source.Series.Count ? s : -1);
            foreach (ChartRow row in spec.Rows)
            {
                var cells = new object[spec.Series.Count + 1]; cells[0] = row.Label;
                for (int s = 0; s < spec.Series.Count; s++) cells[s + 1] = row.Values[s].HasValue ? ChartEditor.DisplayNumber(row.Values[s].Value) : "";
                grid.Rows.Add(cells);
            }
        }
        private ChartSpec ReadMatrix()
        {
            grid.EndEdit();
            var result = ChartStorage.Clone(source);
            result.Kind = (string)kind.SelectedItem; result.LayoutMode = (string)mode.SelectedItem; result.Title = title.Text;
            result.Series.Clear(); result.Rows.Clear();
            for (int s = 1; s < grid.Columns.Count; s++) result.Series.Add(grid.Columns[s].HeaderText);
            if (result.Appearance != null)
            {
                var oldColors = result.Appearance.SeriesColors;
                var oldOrder = result.Appearance.DisplayOrder;
                var newIndices = new Dictionary<int, int>();
                result.Appearance.SeriesColors = new List<string>();
                result.Appearance.DisplayOrder = new List<int>();
                for (int s = 1; s < grid.Columns.Count; s++)
                {
                    int old = grid.Columns[s].Tag is int ? (int)grid.Columns[s].Tag : -1;
                    if (old >= 0) newIndices[old] = s - 1;
                    if (oldColors != null && oldColors.Count > 0)
                        result.Appearance.SeriesColors.Add(old >= 0 && old < oldColors.Count ? oldColors[old] : "#808080");
                }
                if (oldOrder != null && oldOrder.Count > 0)
                {
                    foreach (int old in oldOrder) if (newIndices.ContainsKey(old)) result.Appearance.DisplayOrder.Add(newIndices[old]);
                    for (int s = 0; s < result.Series.Count; s++) if (!result.Appearance.DisplayOrder.Contains(s)) result.Appearance.DisplayOrder.Add(s);
                }
                if (result.Appearance.ValueLabelOffsets != null)
                {
                    var retained = new List<ChartLabelOffset>();
                    foreach (ChartLabelOffset offset in result.Appearance.ValueLabelOffsets)
                        if (newIndices.ContainsKey(offset.SeriesIndex)) { offset.SeriesIndex = newIndices[offset.SeriesIndex]; retained.Add(offset); }
                    result.Appearance.ValueLabelOffsets = retained;
                }
            }
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;
                var item = new ChartRow { Label = Convert.ToString(row.Cells[0].Value) };
                for (int s = 1; s < grid.Columns.Count; s++)
                {
                    string value = Convert.ToString(row.Cells[s].Value);
                    item.Values.Add(String.IsNullOrWhiteSpace(value) ? (double?)null : ChartEditor.Number(value, "Ligne " + (row.Index + 1) + ", " + grid.Columns[s].HeaderText));
                }
                result.Rows.Add(item);
            }
            if (result.Appearance != null && ChartStorage.Serialize(result) != ChartStorage.Serialize(source))
            {
                // Bounds in exported OOXML describe the saved snapshot. They do
                // not establish whether the original user selected fixed bounds.
                // Recompute after numeric edits instead of freezing stale limits.
                bool changed = result.Rows.Count != source.Rows.Count || result.Series.Count != source.Series.Count;
                for (int r = 0; !changed && r < result.Rows.Count; r++)
                {
                    if (result.Rows[r].Values.Count != source.Rows[r].Values.Count) { changed = true; break; }
                    for (int s = 0; s < result.Rows[r].Values.Count; s++) if (result.Rows[r].Values[s] != source.Rows[r].Values[s]) { changed = true; break; }
                }
                if (changed) { result.Appearance.Minimum = null; result.Appearance.Maximum = null; }
                if (result.Rows.Count != source.Rows.Count) result.Appearance.ValueLabelOffsets = new List<ChartLabelOffset>();
            }
            if (result.Appearance != null && (!String.Equals(result.Kind, source.Kind, StringComparison.OrdinalIgnoreCase) || result.LayoutMode != source.LayoutMode))
            {
                // The observed plot and side labels belong to the original chart
                // topology. Preserve colors/fonts, but lay out the new type anew.
                result.Appearance.PlotRect = null;
                result.Appearance.Minimum = null; result.Appearance.Maximum = null;
                result.Appearance.SeriesLabelsAtFirstStack = false;
                result.Appearance.ShowTotals = false; result.Appearance.ShowLegend = true;
                result.Appearance.ValueLabelOffsets = new List<ChartLabelOffset>();
            }
            return result;
        }
        private void Paste()
        {
            try { PasteMatrix(Clipboard.GetText()); }
            catch (Exception ex) { Error(ex.Message); }
        }
        private void PasteMatrix(string text)
        {
            List<List<string>> table = ChartEditor.ParseTsv(text);
            if (table.Count < 2 || table[0].Count < 2 || table[0].Count > 9) throw new ArgumentException("En-tête requis : Catégorie puis un à huit noms de séries.");
            var candidate = new ChartSpec { Kind = (string)kind.SelectedItem, LayoutMode = (string)mode.SelectedItem, Title = title.Text };
            for (int s = 1; s < table[0].Count; s++) candidate.Series.Add(table[0][s]);
            for (int r = 1; r < table.Count; r++)
            {
                if (table[r].Count != table[0].Count) throw new ArgumentException("Ligne " + (r + 1) + " : nombre de colonnes différent de l'en-tête.");
                var row = new ChartRow { Label = table[r][0] };
                for (int s = 1; s < table[r].Count; s++) row.Values.Add(String.IsNullOrWhiteSpace(table[r][s]) ? (double?)null : ChartEditor.Number(table[r][s], "Ligne " + (r + 1)));
                candidate.Rows.Add(row);
            }
            // Paste validates the data before changing the grid. ReadMatrix merges
            // retained appearance later; Apply validates that final model at its real size.
            ChartEngine.Build(candidate, 720, 400); LoadMatrix(candidate);
        }
        private string AskName(string initial)
        {
            using (var dialog = new Form { Text = "Nom de série", StartPosition = FormStartPosition.CenterParent, ClientSize = new Size(330, 85), FormBorderStyle = FormBorderStyle.FixedDialog, MinimizeBox = false, MaximizeBox = false })
            {
                var input = new TextBox { Text = initial, Left = 12, Top = 12, Width = 305 };
                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 235, Top = 45 };
                dialog.Controls.Add(input); dialog.Controls.Add(ok); dialog.AcceptButton = ok;
                return dialog.ShowDialog(this) == DialogResult.OK ? input.Text : null;
            }
        }
        private void Error(string text) { MessageBox.Show(this, text, "Vérifier les données", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }
}
