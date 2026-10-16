using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace OpenSlideCharts
{
    public sealed class GanttEditor : Form
    {
        private readonly TextBox title = new TextBox(), start = new TextBox(), end = new TextBox(), culture = new TextBox();
        private readonly TextBox labelFraction = new TextBox(), fontSize = new TextBox();
        private readonly TextBox activityHeader = new TextBox(), responsibleHeader = new TextBox(), remarkHeader = new TextBox();
        private readonly TextBox workbook = new TextBox(), worksheet = new TextBox(), range = new TextBox();
        private readonly ComboBox scaleMode = new ComboBox(), weekend = new ComboBox(), weekStart = new ComboBox();
        private readonly CheckBox responsible = new CheckBox(), remark = new CheckBox();
        private readonly FlowLayoutPanel weekendDays = new FlowLayoutPanel();
        private readonly CheckBox[] weekendChecks = new CheckBox[7];
        private readonly DataGridView activities = new DataGridView(), items = new DataGridView(), scales = new DataGridView();
        private readonly DataGridView anchors = new DataGridView(), shades = new DataGridView(), links = new DataGridView();
        private readonly double width, height;
        private readonly string[] units = { "Years", "Quarters", "Months", "Weeks", "Days" };
        public ChartSpec Spec { get; private set; }

        public GanttEditor(ChartSpec source) : this(source, 900, 470) { }
        public GanttEditor(ChartSpec source, double frameWidth, double frameHeight)
        {
            SuspendLayout();
            Spec = ChartStorage.Clone(source ?? GanttSamples.Create()); ChartEngine.Build(Spec, frameWidth, frameHeight);
            width = frameWidth; height = frameHeight;
            Text = "OpenSlideCharts — Calendrier Gantt"; StartPosition = FormStartPosition.CenterScreen;
            AutoScaleDimensions = new SizeF(96, 96); AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1100, 720); MinimumSize = new Size(850, 600); Font = new Font("Segoe UI", 9);
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), RowCount = 3, ColumnCount = 1 };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64)); root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44)); Controls.Add(root);
            var header = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 2, ColumnCount = 1 };
            header.RowStyles.Add(new RowStyle(SizeType.Absolute, 26)); header.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            header.Controls.Add(new Label { Text = "Dates ISO. Plusieurs objets par activité ; la dernière ligne est réservée aux jalons.", Dock = DockStyle.Fill }, 0, 0);
            title.Dock = DockStyle.Fill; header.Controls.Add(title, 0, 1); root.Controls.Add(header, 0, 0);
            var tabs = new TabControl { Dock = DockStyle.Fill }; root.Controls.Add(tabs, 0, 1);
            TabPage calendarTab = Tab(tabs, "Calendrier"), activityTab = Tab(tabs, "Activités"), labelTab = Tab(tabs, "Libellés"), itemTab = Tab(tabs, "Objets"), anchorTab = Tab(tabs, "Ancres"), shadeTab = Tab(tabs, "Ombres"), linkTab = Tab(tabs, "Connecteurs"), excelTab = Tab(tabs, "Lien Excel");
            var calendar = Fields(calendarTab, 9, 260);
            Field(calendar, 0, "Début inclus (AAAA-MM-JJ)", start); Field(calendar, 1, "Fin incluse (AAAA-MM-JJ)", end);
            Choices(scaleMode, new string[] { "Automatique", "Manuel" }); Field(calendar, 2, "Échelles", scaleMode);
            Configure(scales, new string[] { "Unité", "Format" }, false);
            scales.Columns.Insert(0, new DataGridViewCheckBoxColumn { Name = "visible", HeaderText = "Afficher" });
            scales.Columns.Add(new DataGridViewCheckBoxColumn { Name = "separators", HeaderText = "Séparateurs" });
            scales.Columns[1].ReadOnly = true; scales.Height = 150; Field(calendar, 3, "Choix manuel (formats facultatifs)", scales);
            calendar.RowStyles[3].Height = 150;
            Choices(weekend, new string[] { "Sans ombre", "Ombrés", "Masqués" }); Field(calendar, 4, "Week-ends", weekend);
            string[] days = { "Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi" };
            Choices(weekStart, days); Field(calendar, 5, "Premier jour de semaine", weekStart);
            weekendDays.AutoSize = true; weekendDays.WrapContents = true; weekendDays.MinimumSize = new Size(0, 28);
            for (int i = 0; i < days.Length; i++) { weekendChecks[i] = new CheckBox { Text = days[i], AutoSize = true, Margin = new Padding(3, 2, 10, 2) }; weekendDays.Controls.Add(weekendChecks[i]); }
            Field(calendar, 6, "Jours de week-end", weekendDays); calendar.RowStyles[6].SizeType = SizeType.AutoSize;
            Field(calendar, 7, "Langue des dates (ex. fr-FR)", culture);
            var fit = Button("Cadrer sur les objets", delegate { Fit(); }); calendar.Controls.Add(fit, 1, 8);

            Configure(activities, new string[] { "ID", "Activité", "Responsable", "Remarque", "Retrait", "Couleur" }, false);
            activities.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Séparateur", Name = "separator" });
            activities.Columns[0].Visible = false;
            GridPage(activityTab, activities, "Plusieurs objets peuvent appartenir à une activité. La ligne des jalons reste en bas ; jalons et crochets y sont admis.",
                new Control[] { Button("Ajouter activité", delegate { activities.Rows.Insert(Math.Max(0, activities.Rows.Count - 1), "row_" + Guid.NewGuid().ToString("N"), "", "", "", "0", "", false); }), Button("Supprimer activité", delegate { DeleteActivities(); }) });
            Configure(items, new string[] { "ID", "Activité ID", "Type", "Début", "Fin", "Libellé", "Couleur", "Mode dates", "Ancre début", "Ancre fin" }, true);
            items.Columns[0].Visible = false; items.Columns[1].HeaderText = "Activité";
            items.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Dates visibles", Name = "dates" });
            GridPage(itemTab, items, "Barres, flèches de processus, jalons ou crochets. Jours inclus = fin de journée incluse. Instants exacts = fin exclue ; requis avec ancres. Couleurs : #RRGGBB. Défilement horizontal pour les ancres.",
                new Control[] { Button("Supprimer sélection", delegate { DeleteRows(items); }) });
            Configure(anchors, new string[] { "ID", "Activité ID", "Indice (0–31)", "Instant ISO" }, true);
            anchors.Columns[0].Visible = false; anchors.Columns[1].HeaderText = "Activité"; anchors.Columns[2].HeaderText = "Colonne date (1–32)";
            GridPage(anchorTab, anchors, "Le tableau Excel crée des ancres de dates, jamais des barres implicites. Choisir ces ancres dans l'onglet Objets pour relier les dates. Une ancre appartient à une activité.",
                new Control[] { Button("Coller tableau Excel", delegate { PasteAnchors(Clipboard.GetText()); }), Button("Supprimer sélection", delegate { DeleteRows(anchors); }) });
            Configure(shades, new string[] { "Début inclus", "Fin incluse", "Libellé", "Couleur" }, true);
            GridPage(shadeTab, shades, "Ombres verticales pour congés et autres périodes ; elles ne compressent pas le calendrier.", new Control[] { Button("Supprimer sélection", delegate { DeleteRows(shades); }) });
            Configure(links, new string[] { "Objet source ID", "Objet cible ID", "Bord source", "Bord cible", "Couleur" }, true);
            links.Columns[0].HeaderText = "Objet source"; links.Columns[1].HeaderText = "Objet cible";
            GridPage(linkTab, links, "Connecteurs visuels entre les bords Début ou Fin des objets. Aucun décalage ni recalcul automatique du planning.", new Control[] { Button("Supprimer sélection", delegate { DeleteRows(links); }) });

            var excel = Fields(excelTab, 5, 280);
            Field(excel, 0, "Classeur enregistré", workbook); Field(excel, 1, "Feuille", worksheet); Field(excel, 2, "Plage sans en-tête", range);
            var excelButtons = new FlowLayoutPanel { Dock = DockStyle.Fill };
            excelButtons.Controls.Add(Button("Parcourir…", delegate { using (var d = new OpenFileDialog { Filter = "Classeurs Excel|*.xlsx;*.xlsm;*.xls;*.xlsb", CheckFileExists = true }) if (d.ShowDialog(this) == DialogResult.OK) workbook.Text = d.FileName; }));
            excelButtons.Controls.Add(Button("Importer Excel", delegate { ImportExcel(); })); excel.Controls.Add(excelButtons, 1, 3);
            Note(excel, 4, "Première ligne : jalons, même vide. Première colonne : libellé ; suivantes : dates/instants exacts. L'import conserve les objets existants et leurs liens par indice d'ancre."); excel.RowStyles[4].Height = 70;
            var labelFields = Fields(labelTab, 7, 300);
            Field(labelFields, 0, "Colonne Responsable", responsible); Field(labelFields, 1, "Colonne Remarque", remark);
            Field(labelFields, 2, "Part réservée aux libellés (0,15–0,65)", labelFraction); Field(labelFields, 3, "Police (6–24 points)", fontSize);
            Field(labelFields, 4, "Titre Activité", activityHeader); Field(labelFields, 5, "Titre Responsable", responsibleHeader); Field(labelFields, 6, "Titre Remarque", remarkHeader);
            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            footer.Controls.Add(Button("Appliquer", delegate { Apply(); }));
            var cancel = Button("Annuler", delegate { DialogResult = DialogResult.Cancel; Close(); }); cancel.DialogResult = DialogResult.Cancel; CancelButton = cancel;
            footer.Controls.Add(cancel); root.Controls.Add(footer, 0, 2);
            items.DefaultValuesNeeded += delegate(object sender, DataGridViewRowEventArgs e) { bool milestone = activities.Rows.Count == 1; e.Row.Cells[0].Value = "item_" + Guid.NewGuid().ToString("N"); e.Row.Cells[1].Value = Cell(activities.Rows[0], 0); e.Row.Cells[2].Value = milestone ? "Milestone" : "Bar"; e.Row.Cells[3].Value = start.Text; e.Row.Cells[4].Value = milestone ? null : end.Text; e.Row.Cells[7].Value = "InclusiveDays"; };
            anchors.DefaultValuesNeeded += delegate(object sender, DataGridViewRowEventArgs e) { e.Row.Cells[0].Value = "date_" + Guid.NewGuid().ToString("N"); e.Row.Cells[1].Value = Spec.Gantt.MilestoneRow.Id; e.Row.Cells[2].Value = 1; };
            tabs.Selected += delegate { RefreshChoices(); };
            LoadSpec(Spec); ResumeLayout(true);
        }

        private static TabPage Tab(TabControl tabs, string name) { var tab = new TabPage(name); tabs.TabPages.Add(tab); return tab; }
        private static Button Button(string text, EventHandler action) { var b = new Button { Text = text, AutoSize = true, MinimumSize = new Size(130, 28) }; b.Click += action; return b; }
        private static void Choices(ComboBox control, string[] values) { control.DropDownStyle = ComboBoxStyle.DropDownList; control.Items.AddRange(values); }
        private static TableLayoutPanel Fields(Control parent, int rows, int labelWidth)
        {
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = rows + 1, Padding = new Padding(10), AutoScroll = true };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, labelWidth)); panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < rows; i++) panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); parent.Controls.Add(panel); return panel;
        }
        private static void Field(TableLayoutPanel panel, int row, string text, Control control)
        { panel.Controls.Add(new Label { Text = text, Dock = DockStyle.Fill, Padding = new Padding(0, 5, 0, 0) }, 0, row); control.Dock = DockStyle.Fill; panel.Controls.Add(control, 1, row); }
        private static void Note(TableLayoutPanel panel, int row, string text) { var label = new Label { Text = text, Dock = DockStyle.Fill }; panel.Controls.Add(label, 0, row); panel.SetColumnSpan(label, 2); }
        private static void Configure(DataGridView grid, string[] headers, bool addRows)
        {
            grid.Dock = DockStyle.Fill; grid.AllowUserToAddRows = addRows; grid.AllowUserToDeleteRows = addRows; grid.RowHeadersWidth = 24;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            grid.BackgroundColor = Color.White; grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            foreach (string header in headers) grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = header, MinimumWidth = 75 });
        }
        private static void GridPage(TabPage page, DataGridView grid, string help, Control[] buttons)
        {
            var panel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 3, ColumnCount = 1, Padding = new Padding(10) };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 60)); panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            panel.Controls.Add(new Label { Text = help, Dock = DockStyle.Fill }, 0, 0); panel.Controls.Add(grid, 0, 1);
            var bar = new FlowLayoutPanel { Dock = DockStyle.Fill }; bar.Controls.AddRange(buttons); panel.Controls.Add(bar, 0, 2); page.Controls.Add(panel);
        }
        private static string Cell(DataGridViewRow row, int col) { return Convert.ToString(row.Cells[col].Value, CultureInfo.CurrentCulture) ?? ""; }
        private static string Optional(string value) { return String.IsNullOrWhiteSpace(value) ? null : value.Trim(); }
        private static bool Flag(DataGridViewRow row, int col) { return row.Cells[col].Value is bool && (bool)row.Cells[col].Value; }
        private static bool Empty(DataGridViewRow row)
        { foreach (DataGridViewCell cell in row.Cells) if (cell.Value is bool ? (bool)cell.Value : !String.IsNullOrWhiteSpace(Convert.ToString(cell.Value, CultureInfo.CurrentCulture))) return false; return true; }
        private static int Integer(string value, string name) { int result; if (!Int32.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out result)) throw new FormatException(name + " : entier requis."); return result; }
        private static void DeleteRows(DataGridView grid)
        {
            var selected = new HashSet<DataGridViewRow>(); foreach (DataGridViewCell cell in grid.SelectedCells) if (!cell.OwningRow.IsNewRow) selected.Add(cell.OwningRow);
            foreach (DataGridViewRow row in selected) grid.Rows.Remove(row);
        }
        private void DeleteActivities()
        {
            var selected = new HashSet<DataGridViewRow>(); foreach (DataGridViewCell cell in activities.SelectedCells) if (Cell(cell.OwningRow, 0) != Spec.Gantt.MilestoneRow.Id) selected.Add(cell.OwningRow);
            foreach (DataGridViewRow row in selected) activities.Rows.Remove(row);
        }
        private void LoadSpec(ChartSpec spec)
        {
            // Imported rows can have new identities. Populate plain cells first,
            // then install the matching choice lists before returning to the UI.
            foreach (DataGridView grid in new[] { items, anchors, links })
            {
                grid.Rows.Clear();
                for (int i = 0; i < grid.Columns.Count; i++) if (grid.Columns[i] is DataGridViewComboBoxColumn)
                {
                    string header = grid.Columns[i].HeaderText; int minimum = grid.Columns[i].MinimumWidth;
                    grid.Columns.RemoveAt(i); grid.Columns.Insert(i, new DataGridViewTextBoxColumn { HeaderText = header, MinimumWidth = minimum });
                }
            }
            GanttSpec g = spec.Gantt;
            title.Text = spec.Title ?? ""; start.Text = g.StartDate; end.Text = g.EndDate; culture.Text = g.Culture;
            scaleMode.SelectedIndex = g.ScaleMode == "Manual" ? 1 : 0; weekend.SelectedIndex = g.WeekendDisplay == "Hidden" ? 2 : g.WeekendDisplay == "Shaded" ? 1 : 0; weekStart.SelectedIndex = g.WeekStartsOn;
            for (int i = 0; i < 7; i++) weekendChecks[i].Checked = g.WeekendDays.Contains(i);
            scales.Rows.Clear();
            var displayedUnits = new List<string>();
            foreach (GanttScale scale in g.Scales ?? new List<GanttScale>()) displayedUnits.Add(scale.Unit);
            foreach (string unit in units) if (!displayedUnits.Contains(unit)) displayedUnits.Add(unit);
            foreach (string unit in displayedUnits)
            {
                GanttScale selected = (g.Scales ?? new List<GanttScale>()).Find(delegate(GanttScale scale) { return scale.Unit == unit; });
                scales.Rows.Add(selected != null, unit, selected == null ? "" : selected.Format, selected == null || selected.Separators);
            }
            activities.Rows.Clear(); foreach (GanttActivity activity in g.Activities) AddActivity(activity);
            AddActivity(g.MilestoneRow); activities.Rows[activities.Rows.Count - 1].Cells[0].ReadOnly = true;
            items.Rows.Clear(); foreach (GanttItem item in g.Items) items.Rows.Add(item.Id, item.RowId, item.Kind, item.StartDate, item.EndDate, item.Label, item.Color, item.DateMode, item.StartAnchor, item.EndAnchor, item.ShowDates);
            anchors.Rows.Clear(); foreach (GanttDateAnchor anchor in g.DateAnchors ?? new List<GanttDateAnchor>()) anchors.Rows.Add(anchor.Id, anchor.RowId, anchor.Index + 1, anchor.Date);
            shades.Rows.Clear(); foreach (GanttShade shade in g.Shades ?? new List<GanttShade>()) shades.Rows.Add(shade.StartDate, shade.EndDate, shade.Label, shade.Color);
            links.Rows.Clear(); foreach (GanttLink link in g.Links ?? new List<GanttLink>()) links.Rows.Add(link.FromId, link.ToId, link.FromEdge, link.ToEdge, link.Color);
            workbook.Text = spec.WorkbookPath ?? ""; worksheet.Text = spec.Worksheet ?? ""; range.Text = spec.RangeAddress ?? "";
            responsible.Checked = g.ShowResponsible; remark.Checked = g.ShowRemark; labelFraction.Text = ChartEditor.DisplayNumber(g.LabelFraction); fontSize.Text = ChartEditor.DisplayNumber(g.FontSize);
            activityHeader.Text = g.ActivityHeader; responsibleHeader.Text = g.ResponsibleHeader; remarkHeader.Text = g.RemarkHeader;
            RefreshChoices();
        }
        private void AddActivity(GanttActivity a) { activities.Rows.Add(a.Id, a.Label, a.Responsible, a.Remark, a.Indent, a.Color, a.Separator); }
        private ChartSpec ReadSpec()
        {
            foreach (var grid in new[] { activities, items, scales, anchors, shades, links }) grid.EndEdit();
            ChartSpec result = ChartStorage.Clone(Spec); GanttSpec g = result.Gantt;
            result.Title = title.Text; result.WorkbookPath = workbook.Text.Trim(); result.Worksheet = worksheet.Text.Trim(); result.RangeAddress = range.Text.Trim();
            g.StartDate = start.Text.Trim(); g.EndDate = end.Text.Trim(); g.Culture = culture.Text.Trim();
            g.ScaleMode = scaleMode.SelectedIndex == 1 ? "Manual" : "Auto"; g.WeekendDisplay = new[] { "Plain", "Shaded", "Hidden" }[weekend.SelectedIndex]; g.WeekStartsOn = weekStart.SelectedIndex;
            g.WeekendDays = new List<int>(); for (int i = 0; i < 7; i++) if (weekendChecks[i].Checked) g.WeekendDays.Add(i);
            g.Scales = new List<GanttScale>(); foreach (DataGridViewRow row in scales.Rows) if (Flag(row, 0)) g.Scales.Add(new GanttScale { Unit = Cell(row, 1), Format = Optional(Cell(row, 2)), Separators = Flag(row, 3) });
            g.Activities = new List<GanttActivity>();
            foreach (DataGridViewRow row in activities.Rows)
            {
                var activity = new GanttActivity { Id = Cell(row, 0).Trim(), Label = Cell(row, 1), Responsible = Cell(row, 2), Remark = Cell(row, 3), Indent = Integer(Cell(row, 4), "Retrait"), Color = Optional(Cell(row, 5)), Separator = Flag(row, 6) };
                if (activity.Id == Spec.Gantt.MilestoneRow.Id) g.MilestoneRow = activity; else g.Activities.Add(activity);
            }
            g.Items = new List<GanttItem>();
            foreach (DataGridViewRow row in items.Rows) if (!row.IsNewRow && !Empty(row)) g.Items.Add(new GanttItem { Id = Cell(row, 0).Trim(), RowId = Cell(row, 1).Trim(), Kind = String.IsNullOrWhiteSpace(Cell(row, 2)) ? "Bar" : Cell(row, 2).Trim(), StartDate = Optional(Cell(row, 3)), EndDate = Optional(Cell(row, 4)), Label = Cell(row, 5), Color = Optional(Cell(row, 6)), DateMode = String.IsNullOrWhiteSpace(Cell(row, 7)) ? "InclusiveDays" : Cell(row, 7).Trim(), StartAnchor = Optional(Cell(row, 8)), EndAnchor = Optional(Cell(row, 9)), ShowDates = Flag(row, 10) });
            g.DateAnchors = new List<GanttDateAnchor>();
            foreach (DataGridViewRow row in anchors.Rows) if (!row.IsNewRow && !Empty(row)) g.DateAnchors.Add(new GanttDateAnchor { Id = Cell(row, 0).Trim(), RowId = Cell(row, 1).Trim(), Index = Integer(Cell(row, 2), "Colonne date") - 1, Date = Cell(row, 3).Trim() });
            g.Shades = new List<GanttShade>(); foreach (DataGridViewRow row in shades.Rows) if (!row.IsNewRow && !Empty(row)) g.Shades.Add(new GanttShade { StartDate = Cell(row, 0).Trim(), EndDate = Cell(row, 1).Trim(), Label = Cell(row, 2), Color = Optional(Cell(row, 3)) });
            g.Links = new List<GanttLink>(); foreach (DataGridViewRow row in links.Rows) if (!row.IsNewRow && !Empty(row)) g.Links.Add(new GanttLink { FromId = Cell(row, 0).Trim(), ToId = Cell(row, 1).Trim(), FromEdge = String.IsNullOrWhiteSpace(Cell(row, 2)) ? "End" : Cell(row, 2).Trim(), ToEdge = String.IsNullOrWhiteSpace(Cell(row, 3)) ? "Start" : Cell(row, 3).Trim(), Color = Optional(Cell(row, 4)) });
            g.ShowResponsible = responsible.Checked; g.ShowRemark = remark.Checked; g.LabelFraction = ChartEditor.Number(labelFraction.Text, "Part libellés"); g.FontSize = ChartEditor.Number(fontSize.Text, "Police");
            g.ActivityHeader = activityHeader.Text; g.ResponsibleHeader = responsibleHeader.Text; g.RemarkHeader = remarkHeader.Text;
            return result;
        }
        private static void ChoiceColumn(DataGridView grid, int index, List<KeyValuePair<string, string>> values)
        {
            grid.EndEdit(); var saved = new List<object>();
            foreach (DataGridViewRow row in grid.Rows)
            {
                object value = row.Cells[index].Value; saved.Add(value);
                if (value != null && !values.Exists(delegate(KeyValuePair<string, string> pair) { return pair.Key == Convert.ToString(value); })) values.Add(new KeyValuePair<string, string>(Convert.ToString(value), Convert.ToString(value)));
            }
            string header = grid.Columns[index].HeaderText;
            grid.Columns.RemoveAt(index); grid.Columns.Insert(index, new DataGridViewComboBoxColumn { HeaderText = header, DataSource = values, DisplayMember = "Value", ValueMember = "Key", MinimumWidth = 100, FlatStyle = FlatStyle.Flat });
            for (int i = 0; i < saved.Count; i++) grid.Rows[i].Cells[index].Value = saved[i];
        }
        private void RefreshChoices()
        {
            var rows = new List<KeyValuePair<string, string>>();
            foreach (DataGridViewRow row in activities.Rows) rows.Add(new KeyValuePair<string, string>(Cell(row, 0), Cell(row, 1)));
            ChoiceColumn(items, 1, new List<KeyValuePair<string, string>>(rows)); ChoiceColumn(anchors, 1, new List<KeyValuePair<string, string>>(rows));
            var dates = new List<KeyValuePair<string, string>> { new KeyValuePair<string, string>("", "Aucune") };
            foreach (DataGridViewRow row in anchors.Rows) if (!row.IsNewRow && !Empty(row)) dates.Add(new KeyValuePair<string, string>(Cell(row, 0), "Date " + Cell(row, 2) + " : " + Cell(row, 3)));
            ChoiceColumn(items, 8, new List<KeyValuePair<string, string>>(dates)); ChoiceColumn(items, 9, new List<KeyValuePair<string, string>>(dates));
            ChoiceColumn(items, 2, Pairs(new[] { "Bar", "Pentagon", "Chevron", "Milestone", "Bracket" }, new[] { "Barre", "Flèche pentagone", "Chevron", "Jalon", "Crochet" }));
            ChoiceColumn(items, 7, Pairs(new[] { "InclusiveDays", "Exact" }, new[] { "Jours inclus", "Instants exacts" }));
            ChoiceColumn(links, 2, Pairs(new[] { "Start", "End" }, new[] { "Début", "Fin" })); ChoiceColumn(links, 3, Pairs(new[] { "Start", "End" }, new[] { "Début", "Fin" }));
            var objects = new List<KeyValuePair<string, string>>();
            foreach (DataGridViewRow row in items.Rows) if (!row.IsNewRow && !Empty(row)) objects.Add(new KeyValuePair<string, string>(Cell(row, 0), Cell(row, 5) + " (" + Cell(row, 2) + ", " + Cell(row, 3) + ")"));
            ChoiceColumn(links, 0, new List<KeyValuePair<string, string>>(objects)); ChoiceColumn(links, 1, new List<KeyValuePair<string, string>>(objects));
        }
        private static List<KeyValuePair<string, string>> Pairs(string[] keys, string[] labels)
        { var result = new List<KeyValuePair<string, string>>(); for (int i = 0; i < keys.Length; i++) result.Add(new KeyValuePair<string, string>(keys[i], labels[i])); return result; }
        private void Apply()
        { try { ChartSpec candidate = ReadSpec(); ChartEngine.Build(candidate, width, height); Spec = candidate; DialogResult = DialogResult.OK; Close(); } catch (Exception error) { Error(error); } }
        private void Fit()
        { try { ChartSpec candidate = GanttData.FitScale(ReadSpec()); ChartEngine.Build(candidate, width, height); LoadSpec(candidate); } catch (Exception error) { Error(error); } }
        private void ImportExcel()
        { try { ChartSpec candidate = ExcelLink.Read(ReadSpec()); ChartEngine.Build(candidate, width, height); LoadSpec(candidate); } catch (Exception error) { Error(error); } }
        private void PasteAnchors(string text)
        { try { ChartSpec candidate = GanttData.ReadTsv(ReadSpec(), text); ChartEngine.Build(candidate, width, height); LoadSpec(candidate); } catch (Exception error) { Error(error); } }
        private void Error(Exception error) { MessageBox.Show(this, error.Message, "OpenSlideCharts", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
    }
}
