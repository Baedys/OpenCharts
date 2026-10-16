using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenSlideCharts
{
    public static class GanttData
    {
        // The reference datasheet supplies activity labels and date anchors.
        // It does not implicitly turn successive dates into bars.
        public static ChartSpec ReadCells(ChartSpec source, object[,] cells, bool date1904)
        {
            if (source == null || source.Gantt == null) throw new ArgumentException("Un Gantt daté est requis.");
            if (cells == null || cells.GetLength(0) < 1 || cells.GetLength(0) > 101 || cells.GetLength(1) < 2 || cells.GetLength(1) > 33)
                throw new ArgumentException("Gantt : 1 à 101 lignes, 2 à 33 colonnes ; première ligne réservée aux jalons, première colonne aux libellés.");
            ChartSpec result = ChartStorage.Clone(source); GanttSpec g = result.Gantt;
            int rows = cells.GetLength(0) - 1, r0 = cells.GetLowerBound(0), c0 = cells.GetLowerBound(1);
            if (rows != g.Activities.Count && g.Items.Count != 0)
                throw new NotSupportedException("Le nombre d'activités a changé. Les objets existants sont conservés : ajustez d'abord la structure dans l'éditeur.");
            var activities = new List<GanttActivity>();
            var anchors = new List<GanttDateAnchor>();
            var previous = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (GanttDateAnchor anchor in g.DateAnchors ?? new List<GanttDateAnchor>()) previous[anchor.RowId + "\n" + anchor.Index] = anchor.Id;
            var identifiers = new HashSet<string>(StringComparer.Ordinal) { g.MilestoneRow.Id };
            foreach (GanttActivity activity in g.Activities) identifiers.Add(activity.Id);
            for (int r = 0; r <= rows; r++)
            {
                GanttActivity activity;
                if (r == 0) activity = g.MilestoneRow;
                else
                {
                    if (r <= g.Activities.Count) activity = g.Activities[r - 1];
                    else
                    {
                        string id; do { id = "activity_" + Guid.NewGuid().ToString("N"); } while (!identifiers.Add(id));
                        activity = new GanttActivity { Id = id };
                    }
                    activities.Add(activity);
                }
                activity.Label = Convert.ToString(cells[r0 + r, c0], CultureInfo.CurrentCulture) ?? "";
                for (int c = 1; c < cells.GetLength(1); c++)
                {
                    object value = cells[r0 + r, c0 + c];
                    if (value == null || value is string && String.IsNullOrWhiteSpace((string)value)) continue;
                    string key = activity.Id + "\n" + (c - 1), id;
                    if (!previous.TryGetValue(key, out id)) id = "date_" + Guid.NewGuid().ToString("N");
                    anchors.Add(new GanttDateAnchor { Id = id, RowId = activity.Id, Index = c - 1, Date = GanttDates.Iso(ReadDate(value, date1904)) });
                }
            }
            g.Activities = activities; g.DateAnchors = anchors;
            ChartEngine.ValidateGantt(result);
            return result;
        }

        public static DateTime ReadDate(object value, bool date1904)
        {
            if (value is DateTime) return DateTime.SpecifyKind((DateTime)value, DateTimeKind.Unspecified);
            if (value is string) return GanttDates.ParseInstant(((string)value).Trim());
            if (!(value is double || value is float || value is decimal || value is int || value is long || value is short))
                throw new FormatException("Une ancre doit contenir une date ISO ou une date numérique Excel, sans erreur de formule.");
            double serial = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            if (Double.IsNaN(serial) || Double.IsInfinity(serial) || serial < (date1904 ? 0 : 1))
                throw new FormatException("Date numérique Excel hors domaine.");
            // Excel's fictional 1900-02-29 has no Gregorian representation.
            if (!date1904 && serial >= 60 && serial < 61) throw new FormatException("Le 29 février 1900 fictif d'Excel n'est pas une date grégorienne.");
            DateTime epoch = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 31);
            try { return epoch.AddDays(!date1904 && serial >= 61 ? serial - 1 : serial); }
            catch (ArgumentOutOfRangeException) { throw new FormatException("Date numérique Excel hors domaine."); }
        }

        public static ChartSpec ReadTsv(ChartSpec source, string text)
        {
            List<List<string>> table = ChartEditor.ParseTsv(text);
            int columns = 0; foreach (List<string> row in table) columns = Math.Max(columns, row.Count);
            if (table.Count < 1 || table.Count > 101 || columns < 2 || columns > 33)
                throw new ArgumentException("Tableau Gantt limité à 101 lignes et 33 colonnes, première ligne des jalons incluse.");
            var cells = new object[table.Count, columns];
            for (int r = 0; r < table.Count; r++) for (int c = 0; c < table[r].Count; c++) cells[r, c] = table[r][c];
            return ReadCells(source, cells, false);
        }

        public static ChartSpec FitScale(ChartSpec source)
        {
            if (source == null || source.Gantt == null) throw new ArgumentException("Un Gantt daté est requis.");
            ChartSpec result = ChartStorage.Clone(source); GanttSpec g = result.Gantt;
            ChartEngine.ValidateGantt(result);
            if (g.Items.Count == 0) throw new ArgumentException("Aucun objet temporel à cadrer.");
            var anchors = new Dictionary<string, GanttDateAnchor>(StringComparer.Ordinal);
            foreach (GanttDateAnchor anchor in g.DateAnchors ?? new List<GanttDateAnchor>()) anchors.Add(anchor.Id, anchor);
            DateTime low = DateTime.MaxValue, high = DateTime.MinValue;
            foreach (GanttItem item in g.Items)
            {
                DateTime start = GanttDates.ParseInstant(String.IsNullOrEmpty(item.StartAnchor) ? item.StartDate : anchors[item.StartAnchor].Date);
                DateTime end = item.Kind == "Milestone" ? start : GanttDates.ParseInstant(String.IsNullOrEmpty(item.EndAnchor) ? item.EndDate : anchors[item.EndAnchor].Date);
                if (item.Kind != "Milestone" && item.DateMode == "Exact" && end.TimeOfDay == TimeSpan.Zero) end = end.AddDays(-1);
                if (start.Date < low) low = start.Date;
                if (end.Date > high) high = end.Date;
            }
            g.StartDate = low.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            g.EndDate = high.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            ChartEngine.ValidateGantt(result); return result;
        }
    }
}
