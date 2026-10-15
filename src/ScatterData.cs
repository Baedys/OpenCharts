using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenSlideCharts
{
    public static class ScatterData
    {
        // Dataset replacement is transactional and never aliases the caller's points.
        // A populated Size column selects Bubble, as in the documented datasheet workflow.
        public static ChartSpec ReplacePoints(ChartSpec source, List<ScatterPoint> points)
        {
            if (source == null || source.Scatter == null) throw new ArgumentException("Modèle de dispersion/bulles requis.");
            if (points == null) throw new ArgumentNullException("points");
            ChartSpec result = ChartStorage.Clone(source);
            result.Scatter.Points = new List<ScatterPoint>(); bool sizes = false;
            foreach (ScatterPoint point in points)
            {
                if (point == null) throw new ArgumentException("Point absent.");
                result.Scatter.Points.Add(new ScatterPoint { Label = point.Label, Group = point.Group, Color = point.Color,
                    Marker = point.Marker, X = point.X, Y = point.Y, Size = point.Size });
                sizes |= point.Size.HasValue;
            }
            result.Kind = sizes ? "Bubble" : "Scatter";
            if (sizes && String.Equals(source.Kind, "Scatter", StringComparison.OrdinalIgnoreCase))
            {
                bool reset = false;
                if (result.Scatter.MarkerScheme == "Shapes") { result.Scatter.MarkerScheme = "Colors"; reset = true; }
                foreach (ScatterPoint point in result.Scatter.Points)
                    if (!String.IsNullOrEmpty(point.Marker) && point.Marker != "Circle" && point.Marker != "None") { point.Marker = null; reset = true; }
                if (reset) Warn(result, "Conversion en bulles : formes de marqueurs remplacées par des cercles.");
            }
            return result;
        }

        public static ChartSpec ReadCells(ChartSpec source, object[,] cells)
        {
            if (source == null || source.Scatter == null) throw new ArgumentException("Modèle de dispersion/bulles requis.");
            if (cells == null || cells.GetLength(0) < 1 || cells.GetLength(0) > 1000 || cells.GetLength(1) < 3 || cells.GetLength(1) > 7)
                throw new ArgumentException("Dispersion/bulles : 1 à 1000 lignes, 3 à 7 colonnes (Libellé, X, Y, Taille, Groupe, Couleur, Marqueur), sans en-tête.");
            int firstRow = cells.GetLowerBound(0), firstColumn = cells.GetLowerBound(1), columns = cells.GetLength(1);
            var points = new List<ScatterPoint>();
            for (int r = firstRow; r <= cells.GetUpperBound(0); r++)
            {
                bool empty = true; for (int c = firstColumn; c <= cells.GetUpperBound(1); c++) if (!Missing(cells[r, c])) empty = false;
                if (empty) continue;
                points.Add(new ScatterPoint { Label = Text(cells[r, firstColumn]), X = Number(cells[r, firstColumn + 1]), Y = Number(cells[r, firstColumn + 2]),
                    Size = columns > 3 ? Number(cells[r, firstColumn + 3]) : null, Group = columns > 4 ? Text(cells[r, firstColumn + 4]) : "",
                    Color = columns > 5 ? OptionalText(cells[r, firstColumn + 5]) : null, Marker = columns > 6 ? OptionalText(cells[r, firstColumn + 6]) : null });
            }
            List<ScatterPoint> old = source.Scatter.Points;
            bool changedCount = old == null || old.Count != points.Count, lostStyles = false;
            if (old != null)
            {
                if (!changedCount)
                    for (int i = 0; i < points.Count; i++)
                    {
                        if (columns <= 5) points[i].Color = old[i].Color;
                        if (columns <= 6) points[i].Marker = old[i].Marker;
                    }
                else foreach (ScatterPoint point in old)
                    lostStyles |= point != null && (columns <= 5 && !String.IsNullOrEmpty(point.Color) || columns <= 6 && !String.IsNullOrEmpty(point.Marker));
            }
            ChartSpec result = ReplacePoints(source, points);
            if (changedCount && lostStyles) Warn(result, "Le nombre de points a changé : couleurs et marqueurs individuels non fournis par Excel réinitialisés.");
            ChartEngine.ValidateScatter(result);
            return result;
        }
        private static bool Missing(object value) { return value == null || value is string && String.IsNullOrWhiteSpace((string)value); }
        private static string Text(object value) { return Convert.ToString(value, CultureInfo.CurrentCulture) ?? ""; }
        private static string OptionalText(object value) { return Missing(value) ? null : Text(value).Trim(); }
        private static double? Number(object value) { return Missing(value) ? (double?)null : ExcelLink.Number(value); }
        private static void Warn(ChartSpec spec, string warning)
        {
            if (spec.ImportWarnings == null) spec.ImportWarnings = new List<string>();
            if (!spec.ImportWarnings.Contains(warning)) spec.ImportWarnings.Add(warning);
        }
    }
}
