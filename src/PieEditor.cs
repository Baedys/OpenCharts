using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    public sealed class PieEditor : Form
    {
        private readonly ComboBox kind = new ComboBox(), order = new ComboBox(), content = new ComboBox();
        private readonly TextBox title = new TextBox(), angle = new TextBox(), hole = new TextBox(), center = new TextBox();
        private readonly TextBox workbook = new TextBox(), worksheet = new TextBox(), range = new TextBox();
        private readonly CheckBox categories = new CheckBox();
        private readonly TextBox digits = new TextBox();
        private readonly DataGridView grid = new DataGridView();
        private readonly double frameWidth, frameHeight;
        private readonly string initialHole;
        private readonly string[] orders = { "SheetClockwise", "SheetCounterclockwise", "Ascending", "Descending" };
        private readonly string[] contents = { "Percent", "Value", "Both", "None" };
        public ChartSpec Spec { get; private set; }

        public PieEditor(ChartSpec initial) : this(initial, 720, 400) { }
        public PieEditor(ChartSpec initial, double width, double height)
        {
            // Build the entire form before WinForms computes the host's DPI
            // scale. Otherwise controls added after an early layout keep their
            // 96-DPI row heights inside PowerPoint's high-DPI process.
            SuspendLayout();
            Spec = ChartStorage.Clone(initial ?? ChartSamples.Create("Pie"));
            ChartEngine.Build(Spec, width, height);
            frameWidth = width; frameHeight = height;
            PieOptions options = Spec.Pie ?? new PieOptions();
            Text = "OpenSlideCharts — Secteurs et anneaux";
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(900, 640); MinimumSize = new Size(720, 530);
            Font = new Font("Segoe UI", 9); BackColor = Color.FromArgb(247, 249, 252);
            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(18), ColumnCount = 1, RowCount = 3 };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44)); Controls.Add(layout);
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170)); header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.Controls.Add(Caption("Type"), 0, 0); header.Controls.Add(Caption("Titre du graphique"), 1, 0);
            Configure(kind, new string[] { "Secteurs", "Anneau" }); kind.SelectedIndex = String.Equals(Spec.Kind, "Doughnut", StringComparison.OrdinalIgnoreCase) ? 1 : 0;
            title.Text = Spec.Title ?? ""; title.Dock = DockStyle.Fill;
            header.Controls.Add(kind, 0, 1); header.Controls.Add(title, 1, 1); layout.Controls.Add(header, 0, 0);
            var tabs = new TabControl { Dock = DockStyle.Fill }; layout.Controls.Add(tabs, 0, 1);
            var dataTab = new TabPage("Données"); var settingsTab = new TabPage("Réglages"); var excelTab = new TabPage("Lien Excel");
            tabs.TabPages.AddRange(new TabPage[] { dataTab, settingsTab, excelTab });
            var dataLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(8) };
            dataLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44)); dataLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); dataLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38)); dataTab.Controls.Add(dataLayout);
            dataLayout.Controls.Add(new Label { Dock = DockStyle.Fill, Text = "Valeurs positives ou nulles, au moins une positive. Couleur : #RRGGBB ou vide. Éclatement : 0 à 50 % du rayon. Collage/import remplace les lignes et leurs réglages." }, 0, 0);
            grid.Dock = DockStyle.Fill; grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            grid.AllowUserToOrderColumns = false; grid.RowHeadersWidth = 34; grid.BackgroundColor = Color.White;
            grid.Columns.Add("label", "Libellé"); grid.Columns.Add("value", "Valeur"); grid.Columns.Add("color", "Couleur"); grid.Columns.Add("explode", "Éclatement (%)");
            foreach (DataGridViewColumn column in grid.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
            dataLayout.Controls.Add(grid, 0, 1);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill };
            actions.Controls.Add(Button("Coller Excel", delegate { PasteClipboard(); }));
            actions.Controls.Add(Button("Ajouter une ligne", delegate { grid.Rows.Add("", "0", "", "0"); }));
            actions.Controls.Add(Button("Supprimer les lignes", delegate {
                var rows = new List<DataGridViewRow>(); foreach (DataGridViewCell cell in grid.SelectedCells) if (!cell.OwningRow.IsNewRow && !rows.Contains(cell.OwningRow)) rows.Add(cell.OwningRow);
                foreach (DataGridViewRow row in rows) grid.Rows.Remove(row);
            })); dataLayout.Controls.Add(actions, 0, 2);
            grid.KeyDown += EditingKeyDown;
            grid.EditingControlShowing += delegate(object sender, DataGridViewEditingControlShowingEventArgs e) {
                e.Control.KeyDown -= EditingKeyDown; e.Control.KeyDown += EditingKeyDown;
            };
            var settings = Fields(settingsTab, 8);
            Configure(order, new string[] { "Feuille, sens horaire", "Feuille, sens antihoraire", "Valeurs croissantes, sens horaire", "Valeurs décroissantes, sens horaire" });
            Configure(content, new string[] { "Pourcentage", "Valeur", "Valeur et pourcentage", "Aucun montant" });
            order.SelectedIndex = Array.IndexOf(orders, options.Order); content.SelectedIndex = Array.IndexOf(contents, options.LabelContent);
            angle.Text = ChartEditor.DisplayNumber(options.StartAngle);
            hole.Text = ChartEditor.DisplayNumber(100 * (options.HoleRatio ?? (kind.SelectedIndex == 1 ? .5 : 0))); initialHole = hole.Text;
            center.Text = options.CenterTitle ?? ""; categories.Text = "Afficher les libellés"; categories.Checked = options.ShowCategory;
            digits.Text = options.PercentDigits.ToString(CultureInfo.InvariantCulture);
            Field(settings, 0, "Départ (degrés ; -90 = haut)", angle); Field(settings, 1, "Diamètre du trou (%)", hole);
            Field(settings, 2, "Ordre des segments", order); Field(settings, 3, "Montants affichés", content);
            Field(settings, 4, "Décimales des pourcentages", digits); Field(settings, 5, "Catégories", categories); Field(settings, 6, "Titre au centre (trou requis)", center);
            var note = new Label { Text = "Rotation positive : sens horaire. Changer le type rétablit le trou par défaut (0 % ou 50 %). Libellés externes ; espace insuffisant signalé avant application.", Dock = DockStyle.Fill };
            settings.Controls.Add(note, 0, 7); settings.SetColumnSpan(note, 2);
            kind.SelectedIndexChanged += delegate { hole.Text = kind.SelectedIndex == 1 ? "50" : "0"; if (kind.SelectedIndex == 0) center.Text = ""; };
            var excel = Fields(excelTab, 5);
            workbook.Text = Spec.WorkbookPath ?? ""; worksheet.Text = Spec.Worksheet ?? ""; range.Text = Spec.RangeAddress ?? "";
            Field(excel, 0, "Classeur enregistré", workbook); Field(excel, 1, "Feuille", worksheet); Field(excel, 2, "Plage sans en-tête (2 colonnes)", range);
            var excelActions = new FlowLayoutPanel { Dock = DockStyle.Fill };
            excelActions.Controls.Add(Button("Parcourir…", delegate {
                using (var dialog = new OpenFileDialog { Filter = "Classeurs Excel|*.xlsx;*.xlsm;*.xls;*.xlsb", CheckFileExists = true }) if (dialog.ShowDialog(this) == DialogResult.OK) workbook.Text = dialog.FileName;
            }));
            excelActions.Controls.Add(Button("Importer Excel", delegate { ImportExcel(); })); excel.Controls.Add(excelActions, 0, 3); excel.SetColumnSpan(excelActions, 2);
            var excelNote = new Label { Dock = DockStyle.Fill, Text = "Lecture des valeurs enregistrées uniquement. Importer remplace les données et les couleurs/éclatements. Le lien est conservé pour l'actualisation manuelle." };
            excel.Controls.Add(excelNote, 0, 4); excel.SetColumnSpan(excelNote, 2);
            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var apply = Button("Appliquer", delegate { Apply(); }); var cancel = Button("Annuler", delegate { DialogResult = DialogResult.Cancel; Close(); });
            cancel.DialogResult = DialogResult.Cancel; footer.Controls.Add(apply); footer.Controls.Add(cancel); layout.Controls.Add(footer, 0, 2); AcceptButton = apply; CancelButton = cancel;
            LoadRows(Spec.Rows, options.Slices);
            ResumeLayout(true);
        }

        private static Label Caption(string text) { return new Label { Text = text, AutoSize = true, Margin = new Padding(3, 6, 3, 3) }; }
        private static Button Button(string text, EventHandler action) { var button = new Button { Text = text, AutoSize = true, MinimumSize = new Size(105, 28) }; button.Click += action; return button; }
        private static void Configure(ComboBox control, string[] choices) { control.DropDownStyle = ComboBoxStyle.DropDownList; control.Items.AddRange(choices); control.Dock = DockStyle.Fill; }
        private static TableLayoutPanel Fields(Control parent, int rows)
        {
            var fields = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = rows, Padding = new Padding(12) };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 245)); fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < rows; i++) fields.RowStyles.Add(new RowStyle(i == rows - 1 ? SizeType.Percent : SizeType.Absolute, i == rows - 1 ? 100 : 38));
            parent.Controls.Add(fields); return fields;
        }
        private static void Field(TableLayoutPanel fields, int row, string caption, Control control) { fields.Controls.Add(Caption(caption), 0, row); control.Dock = DockStyle.Fill; fields.Controls.Add(control, 1, row); }
        private void LoadRows(List<ChartRow> rows, List<PieSliceStyle> styles)
        {
            grid.CancelEdit(); grid.Rows.Clear();
            for (int i = 0; i < rows.Count; i++)
            {
                PieSliceStyle style = styles == null ? null : styles.Find(delegate(PieSliceStyle s) { return s.Index == i; });
                grid.Rows.Add(rows[i].Label ?? "", ChartEditor.DisplayNumber(rows[i].Value), style == null ? "" : style.Color ?? "", ChartEditor.DisplayNumber(style == null ? 0 : style.Explode * 100));
            }
        }
        private ChartSpec ReadSpec(bool readRows)
        {
            grid.EndEdit();
            ChartSpec result = ChartStorage.Clone(Spec);
            int percentDigits;
            if (!Int32.TryParse(digits.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out percentDigits))
                throw new ArgumentException("Décimales des pourcentages : entier entre 0 et 6 attendu.");
            result.Kind = kind.SelectedIndex == 1 ? "Doughnut" : "Pie"; result.Title = title.Text.Trim(); result.Rows = new List<ChartRow>();
            result.WorkbookPath = workbook.Text.Trim(); result.Worksheet = worksheet.Text.Trim(); result.RangeAddress = range.Text.Trim();
            result.Pie = new PieOptions { StartAngle = ChartEditor.Number(angle.Text, "Angle"), HoleRatio = ChartEditor.Number(hole.Text, "Trou") / 100,
                Order = orders[order.SelectedIndex], LabelContent = contents[content.SelectedIndex], ShowCategory = categories.Checked,
                PercentDigits = percentDigits, CenterTitle = center.Text.Trim() };
            if ((Spec.Pie == null || !Spec.Pie.HoleRatio.HasValue) && result.Kind.Equals(Spec.Kind, StringComparison.OrdinalIgnoreCase) && hole.Text == initialHole) result.Pie.HoleRatio = null;
            if (!readRows) return result;
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;
                string label = Cell(row, 0), value = Cell(row, 1), color = Cell(row, 2).Trim(), explode = Cell(row, 3);
                if (String.IsNullOrWhiteSpace(label) && String.IsNullOrWhiteSpace(value) && color.Length == 0 && String.IsNullOrWhiteSpace(explode)) continue;
                int index = result.Rows.Count; string location = "Ligne " + (row.Index + 1);
                result.Rows.Add(new ChartRow { Label = label, Value = ChartEditor.Number(value, location + ", Valeur") });
                double distance = String.IsNullOrWhiteSpace(explode) ? 0 : ChartEditor.Number(explode, location + ", Éclatement") / 100;
                if (color.Length > 0 || distance != 0) result.Pie.Slices.Add(new PieSliceStyle { Index = index, Color = color.Length == 0 ? null : color, Explode = distance });
            }
            return result;
        }
        private static string Cell(DataGridViewRow row, int column) { return Convert.ToString(row.Cells[column].Value, CultureInfo.CurrentCulture) ?? ""; }
        private void ValidateCandidate(ChartSpec candidate) { ChartEngine.Build(candidate, frameWidth, frameHeight); }
        private void Apply()
        {
            try { ChartSpec candidate = ReadSpec(true); ValidateCandidate(candidate); Spec = candidate; DialogResult = DialogResult.OK; Close(); }
            catch (Exception error) { ShowError(error); }
        }
        private void PasteClipboard()
        {
            try { PasteData(Clipboard.GetText()); } catch (Exception error) { ShowError(error); }
        }
        private void EditingKeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Control || e.KeyCode != Keys.V) return;
            try
            {
                if (!Clipboard.ContainsText()) return;
                string text = Clipboard.GetText();
                if (text.IndexOf('\t') < 0 && text.TrimEnd('\r', '\n').IndexOf('\n') < 0) return;
                e.Handled = true; e.SuppressKeyPress = true; PasteData(text);
            }
            catch (Exception error) { ShowError(error); }
        }
        private void PasteData(string text)
        {
            List<List<string>> table = ChartEditor.ParseTsv(text); var rows = new List<ChartRow>();
            foreach (List<string> cells in table)
            {
                if (cells.TrueForAll(delegate(string s) { return String.IsNullOrWhiteSpace(s); })) continue;
                string first = cells[0].Trim().ToLowerInvariant();
                if (rows.Count == 0 && cells.Count == 2 && (first == "label" || first == "libellé" || first == "libelle" || first == "category" || first == "catégorie") && (cells[1].Trim().Equals("value", StringComparison.OrdinalIgnoreCase) || cells[1].Trim().Equals("valeur", StringComparison.OrdinalIgnoreCase))) continue;
                if (cells.Count != 2) throw new ArgumentException("Collage : deux colonnes attendues, Libellé et Valeur.");
                rows.Add(new ChartRow { Label = cells[0], Value = ChartEditor.Number(cells[1], "Ligne collée " + (rows.Count + 1)) });
            }
            if (rows.Count == 0 || rows.Count > 80) throw new ArgumentException("Collage : 1 à 80 lignes attendues.");
            ChartSpec candidate = ReadSpec(false); candidate.Rows = rows; ValidateCandidate(candidate);
            LoadRows(rows, null);
        }
        private void ImportExcel()
        {
            try { UseWaitCursor = true; ChartSpec candidate = ExcelLink.Read(ReadSpec(false)); ValidateCandidate(candidate); LoadRows(candidate.Rows, null); }
            catch (Exception error) { ShowError(error); }
            finally { UseWaitCursor = false; }
        }
        private void ShowError(Exception error) { MessageBox.Show(this, error.Message, "Vérifier les données", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }
}
