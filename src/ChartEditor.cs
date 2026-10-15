using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    public class ChartEditor : Form
    {
        private readonly ComboBox kind = new ComboBox();
        private readonly TextBox title = new TextBox();
        private readonly TextBox workbook = new TextBox();
        private readonly TextBox worksheet = new TextBox();
        private readonly TextBox range = new TextBox();
        private readonly Label help = new Label();
        private readonly DataGridView grid = new DataGridView();
        public ChartSpec Spec { get; private set; }

        public ChartEditor(ChartSpec initial)
        {
            Spec = Clone(initial ?? ChartSamples.Create("Waterfall"));
            Text = "OpenSlideCharts — Éditer le graphique";
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(900, 680);
            MinimumSize = new Size(760, 620);
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.FromArgb(247, 249, 252);
            var layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.Padding = new Padding(20);
            layout.ColumnCount = 1;
            layout.RowCount = 7;
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 68));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 55));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
            Controls.Add(layout);

            var heading = new Label();
            heading.Text = "Données et présentation";
            heading.Font = new Font(Font.FontFamily, 16F, FontStyle.Bold);
            heading.ForeColor = Color.FromArgb(26, 47, 77);
            heading.Dock = DockStyle.Fill;
            layout.Controls.Add(heading, 0, 0);

            var fields = new TableLayoutPanel();
            fields.Dock = DockStyle.Fill;
            fields.ColumnCount = 2;
            fields.RowCount = 2;
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            fields.Controls.Add(Caption("Type de graphique"), 0, 0);
            fields.Controls.Add(Caption("Titre"), 1, 0);
            kind.DropDownStyle = ComboBoxStyle.DropDownList;
            kind.Items.AddRange(new object[] { "Waterfall", "Column", "Bar", "Line", "Mekko", "Gantt" });
            kind.Dock = DockStyle.Fill;
            title.Dock = DockStyle.Fill;
            fields.Controls.Add(kind, 0, 1);
            fields.Controls.Add(title, 1, 1);
            layout.Controls.Add(fields, 0, 1);

            help.Dock = DockStyle.Fill;
            help.ForeColor = Color.FromArgb(66, 84, 108);
            help.Padding = new Padding(0, 8, 0, 4);
            layout.Controls.Add(help, 0, 2);
            kind.SelectedIndexChanged += delegate { UpdateHelp(); };

            grid.Dock = DockStyle.Fill;
            grid.BackgroundColor = Color.White;
            grid.BorderStyle = BorderStyle.FixedSingle;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AllowUserToAddRows = true;
            grid.AllowUserToDeleteRows = true;
            grid.AllowUserToOrderColumns = false;
            grid.RowHeadersWidth = 35;
            grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            grid.MultiSelect = true;
            grid.ColumnHeadersHeight = 32;
            grid.RowTemplate.Height = 27;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(244, 247, 251);
            grid.Columns.Add("Label", "Libellé");
            grid.Columns.Add("Value", "Valeur");
            grid.Columns.Add("Secondary", "Secondaire");
            grid.Columns[0].FillWeight = 190;
            var total = new DataGridViewCheckBoxColumn();
            total.Name = "Total";
            total.HeaderText = "Total";
            total.FillWeight = 65;
            grid.Columns.Add(total);
            foreach (DataGridViewColumn column in grid.Columns) column.SortMode = DataGridViewColumnSortMode.NotSortable;
            grid.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Control && e.KeyCode == Keys.V)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    PasteTable();
                }
            };
            grid.EditingControlShowing += delegate(object sender, DataGridViewEditingControlShowingEventArgs e)
            {
                e.Control.KeyDown -= EditingKeyDown;
                e.Control.KeyDown += EditingKeyDown;
            };
            grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e)
            {
                e.ThrowException = false;
                ShowError("Valeur de cellule non valide.");
            };
            layout.Controls.Add(grid, 0, 3);

            var actions = new FlowLayoutPanel();
            actions.Dock = DockStyle.Fill;
            actions.WrapContents = false;
            actions.Controls.Add(Button("Coller tableau Excel", delegate { PasteTable(); }));
            actions.Controls.Add(Button("Charger un exemple", delegate
            {
                LoadRows(ChartSamples.Create(CurrentKind()).Rows);
            }));
            var pasteHint = Caption("Collage TSV : remplace toutes les lignes. Ctrl+V dans la grille.");
            pasteHint.Margin = new Padding(10, 10, 0, 0);
            actions.Controls.Add(pasteHint);
            layout.Controls.Add(actions, 0, 4);

            var excel = new GroupBox();
            excel.Text = "Source Excel (facultatif)";
            excel.Dock = DockStyle.Fill;
            excel.Padding = new Padding(10, 20, 10, 7);
            var excelFields = new TableLayoutPanel();
            excelFields.Dock = DockStyle.Fill;
            excelFields.ColumnCount = 4;
            excelFields.RowCount = 2;
            excelFields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 55));
            excelFields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            excelFields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            excelFields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            excelFields.Controls.Add(Caption("Fichier"), 0, 0);
            workbook.Dock = DockStyle.Fill;
            excelFields.Controls.Add(workbook, 1, 0);
            excelFields.SetColumnSpan(workbook, 2);
            var browse = Button("Parcourir…", BrowseWorkbook);
            excelFields.Controls.Add(browse, 3, 0);
            excelFields.Controls.Add(Caption("Feuille"), 0, 1);
            worksheet.Dock = DockStyle.Fill;
            excelFields.Controls.Add(worksheet, 1, 1);
            excelFields.Controls.Add(Caption("Plage (A1:D8)"), 2, 1);
            var rangePanel = new TableLayoutPanel();
            rangePanel.Dock = DockStyle.Fill;
            rangePanel.ColumnCount = 2;
            rangePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            rangePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            range.Dock = DockStyle.Fill;
            rangePanel.Controls.Add(range, 0, 0);
            rangePanel.Controls.Add(Button("Importer Excel", ImportExcel), 1, 0);
            excelFields.Controls.Add(rangePanel, 3, 1);
            excel.Controls.Add(excelFields);
            layout.Controls.Add(excel, 0, 5);

            var footer = new FlowLayoutPanel();
            footer.Dock = DockStyle.Fill;
            footer.FlowDirection = FlowDirection.RightToLeft;
            footer.Padding = new Padding(0, 7, 0, 0);
            var ok = Button("Appliquer", Apply);
            ok.BackColor = Color.FromArgb(30, 86, 160);
            ok.ForeColor = Color.White;
            ok.FlatStyle = FlatStyle.Flat;
            var cancel = Button("Annuler", delegate { DialogResult = DialogResult.Cancel; Close(); });
            cancel.DialogResult = DialogResult.Cancel;
            footer.Controls.Add(ok);
            footer.Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;
            layout.Controls.Add(footer, 0, 6);
            LoadSpec(Spec);
        }

        private static Label Caption(string text)
        {
            return new Label { Text = text, AutoSize = true, Margin = new Padding(3, 5, 3, 3) };
        }

        private static Button Button(string text, EventHandler handler)
        {
            var result = new Button { Text = text, AutoSize = true, MinimumSize = new Size(105, 29), Margin = new Padding(3) };
            result.Click += handler;
            return result;
        }

        private string CurrentKind() { return kind.SelectedItem == null ? "Waterfall" : kind.SelectedItem.ToString(); }

        private void UpdateHelp()
        {
            switch (CurrentKind())
            {
                case "Waterfall": help.Text = "Valeur = variation. Total coché = montant absolu (pas une variation). Secondaire ignoré."; break;
                case "Mekko": help.Text = "Valeur = largeur de catégorie (> 0). Secondaire = pourcentage de hauteur, entre 0 et 100. Total ignoré."; break;
                case "Gantt": help.Text = "Valeur = départ, en jours depuis l’origine. Secondaire = durée, en jours. Total ignoré."; break;
                default: help.Text = "Valeur = montant de chaque catégorie. Secondaire et Total ignorés. Libellés et titre restent éditables."; break;
            }
        }

        private void LoadSpec(ChartSpec spec)
        {
            foreach (object option in kind.Items)
                if (String.Equals(option.ToString(), spec.Kind, StringComparison.OrdinalIgnoreCase)) kind.SelectedItem = option;
            if (kind.SelectedIndex < 0) kind.SelectedIndex = 0;
            title.Text = spec.Title ?? "";
            workbook.Text = spec.WorkbookPath ?? "";
            worksheet.Text = spec.Worksheet ?? "";
            range.Text = spec.RangeAddress ?? "";
            LoadRows(spec.Rows);
        }

        private void LoadRows(List<ChartRow> rows)
        {
            grid.CancelEdit();
            grid.Rows.Clear();
            if (rows == null) return;
            foreach (var row in rows)
                grid.Rows.Add(row.Label ?? "", DisplayNumber(row.Value), DisplayNumber(row.Secondary), row.Total);
        }

        internal static string DisplayNumber(double value)
        {
            string result = value.ToString("R", CultureInfo.InvariantCulture);
            // Explicit exponent disambiguates generated three-digit decimal fractions.
            if (Regex.IsMatch(result, @"^[+-]?\d{1,3}\.\d{3}$")) result += "e0";
            return result;
        }

        private ChartSpec ReadSpec(bool readRows)
        {
            grid.EndEdit();
            var result = new ChartSpec();
            result.Kind = CurrentKind();
            result.Title = title.Text.Trim();
            result.WorkbookPath = workbook.Text.Trim();
            result.Worksheet = worksheet.Text.Trim();
            result.RangeAddress = range.Text.Trim();
            result.Rows = new List<ChartRow>();
            if (!readRows) return result;
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row.IsNewRow) continue;
                string label = Cell(row, 0);
                string value = Cell(row, 1);
                string secondary = Cell(row, 2);
                bool isTotal = row.Cells[3].Value is bool && (bool)row.Cells[3].Value;
                if (String.IsNullOrWhiteSpace(label) && String.IsNullOrWhiteSpace(value) && String.IsNullOrWhiteSpace(secondary) && !isTotal) continue;
                string location = "Ligne " + (row.Index + 1).ToString(CultureInfo.InvariantCulture);
                result.Rows.Add(new ChartRow { Label = label, Value = Number(value, location + ", Valeur"),
                    Secondary = String.IsNullOrWhiteSpace(secondary) ? 0 : Number(secondary, location + ", Secondaire"), Total = isTotal });
            }
            return result;
        }

        private static string Cell(DataGridViewRow row, int column)
        {
            return Convert.ToString(row.Cells[column].Value, CultureInfo.CurrentCulture) ?? "";
        }

        private void Apply(object sender, EventArgs e)
        {
            try
            {
                ChartSpec candidate = ReadSpec(true);
                ChartEngine.Build(candidate, 720, 400);
                Spec = candidate;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception error) { ShowError(error.Message); }
        }

        private void BrowseWorkbook(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Choisir le classeur Excel";
                dialog.Filter = "Classeurs Excel|*.xlsx;*.xlsm;*.xls;*.xlsb|Tous les fichiers|*.*";
                dialog.CheckFileExists = true;
                if (dialog.ShowDialog(this) == DialogResult.OK) workbook.Text = dialog.FileName;
            }
        }

        private void ImportExcel(object sender, EventArgs e)
        {
            try
            {
                ChartSpec candidate = ReadSpec(false);
                if (String.IsNullOrWhiteSpace(candidate.WorkbookPath) || String.IsNullOrWhiteSpace(candidate.Worksheet) || String.IsNullOrWhiteSpace(candidate.RangeAddress))
                    throw new ArgumentException("Renseignez le fichier Excel, la feuille et la plage avant l’import.");
                UseWaitCursor = true;
                ChartSpec imported = ExcelLink.Read(candidate);
                if (imported == null) throw new InvalidOperationException("L’import Excel n’a retourné aucune donnée.");
                LoadRows(imported.Rows);
            }
            catch (Exception error) { ShowError(error.Message); }
            finally { UseWaitCursor = false; }
        }

        private void EditingKeyDown(object sender, KeyEventArgs e)
        {
            if (!e.Control || e.KeyCode != Keys.V) return;
            // Keep ordinary single-cell text paste; intercept a full Excel range.
            try
            {
                string text = Clipboard.GetText();
                if (text.IndexOf('\t') < 0 && text.TrimEnd('\r', '\n').IndexOf('\n') < 0) return;
                e.Handled = true;
                e.SuppressKeyPress = true;
                PasteTable();
            }
            catch (Exception error) { ShowError(error.Message); }
        }

        private void PasteTable()
        {
            try
            {
                if (!Clipboard.ContainsText()) throw new ArgumentException("Presse-papiers vide : copiez une plage Excel avec libellés et valeurs.");
                var table = ParseTsv(Clipboard.GetText());
                var rows = new List<ChartRow>();
                for (int i = 0; i < table.Count; i++)
                {
                    List<string> cells = table[i];
                    bool empty = true;
                    foreach (string cell in cells) if (!String.IsNullOrWhiteSpace(cell)) empty = false;
                    if (empty) continue;
                    if (i == 0 && cells.Count >= 2 && IsHeader(cells[0], cells[1])) continue;
                    if (cells.Count < 2 || cells.Count > 4) throw new ArgumentException("Ligne " + (i + 1) + " : 2 à 4 colonnes attendues (Libellé, Valeur, Secondaire, Total).");
                    string location = "Ligne collée " + (i + 1);
                    rows.Add(new ChartRow { Label = cells[0], Value = Number(cells[1], location + ", Valeur"),
                        Secondary = cells.Count > 2 && !String.IsNullOrWhiteSpace(cells[2]) ? Number(cells[2], location + ", Secondaire") : 0,
                        Total = cells.Count > 3 && Boolean(cells[3], location) });
                }
                if (rows.Count == 0) throw new ArgumentException("Aucune ligne de données dans le presse-papiers.");
                LoadRows(rows);
            }
            catch (Exception error) { ShowError(error.Message); }
        }

        private static bool IsHeader(string first, string second)
        {
            string a = first.Trim().ToLowerInvariant();
            string b = second.Trim().ToLowerInvariant();
            return (a == "libellé" || a == "libelle" || a == "label" || a == "catégorie" || a == "category") &&
                   (b == "valeur" || b == "value" || b == "départ" || b == "depart" || b == "largeur");
        }

        private static bool Boolean(string input, string location)
        {
            switch (input.Trim().ToLowerInvariant())
            {
                case "": case "0": case "false": case "faux": case "non": return false;
                case "1": case "true": case "vrai": case "oui": case "x": return true;
                default: throw new ArgumentException(location + ", Total : utilisez VRAI/FAUX, oui/non ou 1/0.");
            }
        }

        internal static double Number(string input, string location)
        {
            string s = input.Trim().Replace('\u00a0', ' ').Replace('\u202f', ' ').Replace('\u2212', '-');
            string normalized;
            if (Regex.IsMatch(s, @"^[+-]?\d{1,3}[.,]\d{3}$"))
                throw new ArgumentException(location + " : nombre ambigu « " + input + " ». Utilisez un espace pour les milliers (1 234), ou une décimale explicite (1,2340 / 1.2340).");
            if (Regex.IsMatch(s, @"^[+-]?\d{1,3}( \d{3})+([.,]\d+)?$"))
                normalized = s.Replace(" ", "").Replace(',', '.');
            else if (Regex.IsMatch(s, @"^[+-]?\d{1,3}(,\d{3})+\.\d+$"))
                normalized = s.Replace(",", "");
            else if (Regex.IsMatch(s, @"^[+-]?\d{1,3}(\.\d{3})+,\d+$"))
                normalized = s.Replace(".", "").Replace(',', '.');
            else if (Regex.IsMatch(s, @"^[+-]?\d{1,3}(,\d{3}){2,}$"))
                normalized = s.Replace(",", "");
            else if (Regex.IsMatch(s, @"^[+-]?\d{1,3}(\.\d{3}){2,}$"))
                normalized = s.Replace(".", "");
            else if (Regex.IsMatch(s, @"^[+-]?(\d+([.,]\d*)?|[.,]\d+)([eE][+-]?\d+)?$"))
                normalized = s.Replace(',', '.');
            else throw new ArgumentException(location + " : nombre non reconnu « " + input + " ».");
            double result;
            if (!Double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out result) || Double.IsNaN(result) || Double.IsInfinity(result))
                throw new ArgumentException(location + " : nombre hors limites.");
            return result;
        }

        internal static List<List<string>> ParseTsv(string input)
        {
            var result = new List<List<string>>();
            var row = new List<string>();
            var cell = new System.Text.StringBuilder();
            bool quoted = false;
            bool closed = false;
            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                if (quoted)
                {
                    if (c == '"')
                    {
                        if (i + 1 < input.Length && input[i + 1] == '"') { cell.Append('"'); i++; }
                        else { quoted = false; closed = true; }
                    }
                    else cell.Append(c);
                    continue;
                }
                if (c == '"' && cell.Length == 0 && !closed) { quoted = true; continue; }
                if (c == '\t' || c == '\r' || c == '\n')
                {
                    row.Add(cell.ToString()); cell.Length = 0; closed = false;
                    if (c != '\t')
                    {
                        result.Add(row); row = new List<string>();
                        if (c == '\r' && i + 1 < input.Length && input[i + 1] == '\n') i++;
                    }
                }
                else
                {
                    if (closed) throw new ArgumentException("Collage TSV non valide : caractère après un champ entre guillemets.");
                    cell.Append(c);
                }
            }
            if (quoted) throw new ArgumentException("Collage TSV non valide : guillemet fermant manquant.");
            if (cell.Length > 0 || row.Count > 0 || closed) { row.Add(cell.ToString()); result.Add(row); }
            return result;
        }

        private static ChartSpec Clone(ChartSpec source)
        {
            var result = new ChartSpec { Kind = source.Kind, Title = source.Title, WorkbookPath = source.WorkbookPath,
                Worksheet = source.Worksheet, RangeAddress = source.RangeAddress, Rows = new List<ChartRow>() };
            if (source.Rows != null)
                foreach (var row in source.Rows) result.Rows.Add(new ChartRow { Label = row.Label, Value = row.Value, Secondary = row.Secondary, Total = row.Total });
            return result;
        }

        private void ShowError(string message)
        {
            MessageBox.Show(this, message, "Vérifier les données", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
