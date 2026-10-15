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
        public ChartSpec Spec { get; private set; }

        public SeriesEditor(ChartSpec initial)
        {
            source = ChartStorage.Clone(initial);
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
            root.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Cellule vide = donnée absente. Collage : première ligne = noms de séries, première colonne = catégories.\nClustered : groupé ; Stacked : empilé ; Percent : empilé à 100 %. Double-clic sur un en-tête pour renommer." }, 0, 1);
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
                try { ChartSpec candidate = ReadMatrix(); ChartEngine.Build(candidate, 720, 400); Spec = candidate; DialogResult = DialogResult.OK; Close(); }
                catch (Exception ex) { Error(ex.Message); }
            });
            actions.Controls.Add(apply);
            var cancel = MakeButton("Annuler", delegate { DialogResult = DialogResult.Cancel; Close(); });
            actions.Controls.Add(cancel); AcceptButton = apply; CancelButton = cancel;
            root.Controls.Add(actions, 0, 3);
        }
        private static Button MakeButton(string label, EventHandler action)
        { var b = new Button { Text = label, AutoSize = true, Height = 30 }; b.Click += action; return b; }
        private void AddSeries(string name)
        {
            int index = grid.Columns.Add("S_" + Guid.NewGuid().ToString("N"), name);
            grid.Columns[index].SortMode = DataGridViewColumnSortMode.NotSortable;
        }
        private void LoadMatrix(ChartSpec spec)
        {
            grid.CancelEdit(); grid.Rows.Clear(); grid.Columns.Clear();
            grid.Columns.Add("Category", "Catégorie"); grid.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            foreach (string name in spec.Series) AddSeries(name);
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
            return result;
        }
        private void Paste()
        {
            try
            {
                List<List<string>> table = ChartEditor.ParseTsv(Clipboard.GetText());
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
                ChartEngine.Build(candidate, 720, 400); LoadMatrix(candidate);
            }
            catch (Exception ex) { Error(ex.Message); }
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
