using System;
using System.Collections.Generic;

namespace OpenSlideCharts
{
    public static partial class ChartEngine
    {
        private static readonly int[] SeriesColors = { Blue, Green, 4438767, 10907530, 13998939, 7369077, Red, 11119017 };

        // Null means a missing observation, not zero. Stacked positive and negative
        // values have independent baselines. This is an original renderer, not a
        // reproduction of the reference product's placement algorithm.
        private static void MultiSeries(Layout layout, ChartSpec spec, double width, double height)
        {
            string kind = spec.Kind.ToLowerInvariant();
            string mode = spec.LayoutMode ?? "Clustered";
            if (kind != "column" && kind != "bar" && kind != "line")
                throw new NotSupportedException("Multiple series require Column, Bar or Line.");
            if (mode != "Clustered" && mode != "Stacked" && mode != "Percent")
                throw new ArgumentException("LayoutMode must be Clustered, Stacked or Percent.");
            if (kind == "line" && mode != "Clustered")
                throw new NotSupportedException("Stacked lines are not supported.");
            int series = spec.Series.Count, categories = spec.Rows.Count;
            if (series > 8) throw new ArgumentException("Provide at most eight series.");
            bool horizontal = kind == "bar", stacked = mode != "Clustered";
            double[,] starts = new double[categories, series], ends = new double[categories, series];
            double[,] shown = new double[categories, series];
            double low = 0, high = 0;
            for (int c = 0; c < categories; c++)
            {
                List<double?> values = spec.Rows[c].Values;
                if (values == null || values.Count != series) throw new ArgumentException("Each category must have one value or null per series.");
                double largest = 0, sum = 0;
                foreach (double? cell in values)
                {
                    if (!cell.HasValue) continue;
                    if (!Finite(cell.Value)) throw new ArgumentException("Series values must be finite.");
                    if (mode == "Percent" && cell.Value < 0) throw new ArgumentException("Percent stacks require nonnegative values.");
                    largest = Math.Max(largest, Math.Abs(cell.Value));
                }
                if (mode == "Percent" && largest > 0)
                    foreach (double? cell in values) if (cell.HasValue) sum += cell.Value / largest;
                double positive = 0, negative = 0;
                for (int s = 0; s < series; s++)
                {
                    if (!values[s].HasValue) continue;
                    double value = values[s].Value;
                    if (mode == "Percent") value = largest == 0 ? 0 : (value / largest) / sum * 100;
                    double start = stacked ? (value < 0 ? negative : positive) : 0;
                    double end = start + value;
                    if (!Finite(end)) throw new ArgumentException("Stacked cumulative value overflow.");
                    if (value < 0) negative = end; else positive = end;
                    starts[c, s] = start; ends[c, s] = end; shown[c, s] = value;
                    low = Math.Min(low, end); high = Math.Max(high, end);
                }
            }
            if (mode == "Percent") high = 100;
            Scale scale = new Scale(low, high);
            int legendCols = Math.Min(series, Math.Max(1, (int)((width - 28) / 135)));
            int legendRows = (series + legendCols - 1) / legendCols;
            double legendCell = (width - 28) / legendCols;
            for (int s = 0; s < series; s++)
            {
                double x = 14 + (s % legendCols) * legendCell, y = 42 + (s / legendCols) * 18;
                Shape(layout, "rect", x, y + 3, 9, 9, SeriesColors[s]);
                Text(layout, spec.Series[s] ?? ("Series " + (s + 1)), x + 13, y, legendCell - 16, 17, 9, Ink, "left");
            }
            double left = horizontal ? Math.Min(130, width * .28) : 68;
            double top = 50 + legendRows * 18, plotW = width - left - 24, plotH = height - top - 44;
            if (plotH < 55) throw new ArgumentException("Chart is too short for its legend; increase the height.");
            if (!horizontal) AxisY(layout, scale, left, top, plotW, plotH);
            else for (int t = 0; t <= 4; t++)
            {
                double fraction = t / 4.0, x = left + fraction * plotW;
                Line(layout, x, top, 0, plotH, Gray);
                Text(layout, Number(scale.Tick(fraction)), Math.Max(0, x - 24), top + plotH + 7, 48, 18, 9, Ink);
            }
            if (horizontal) Line(layout, left + scale.Fraction(0) * plotW, top, 0, plotH, Ink);
            else Line(layout, left, top + plotH * (1 - scale.Fraction(0)), plotW, 0, Ink);
            double step = (horizontal ? plotH : plotW) / categories;
            double fullBar = step * .7, bar = stacked ? fullBar : fullBar / series;
            double[] previousX = new double[series], previousY = new double[series];
            bool[] connected = new bool[series];
            for (int c = 0; c < categories; c++)
            {
                if (horizontal) Text(layout, spec.Rows[c].Label ?? "", 4, top + c * step, left - 10, step, Font(step), Ink, "right");
                else Text(layout, spec.Rows[c].Label ?? "", left + c * step, top + plotH + 7, step, 30, Font(step), Ink);
                for (int s = 0; s < series; s++)
                {
                    if (!spec.Rows[c].Values[s].HasValue) { connected[s] = false; continue; }
                    double start = scale.Fraction(starts[c, s]), end = scale.Fraction(ends[c, s]);
                    double center = (horizontal ? top : left) + step * (c + .5);
                    if (kind == "line")
                    {
                        double y = top + plotH * (1 - end);
                        if (connected[s]) Line(layout, previousX[s], previousY[s], center - previousX[s], y - previousY[s], SeriesColors[s]);
                        Shape(layout, "ellipse", center - 2.5, y - 2.5, 5, 5, SeriesColors[s]);
                        previousX[s] = center; previousY[s] = y; connected[s] = true;
                        continue;
                    }
                    double offset = center - fullBar / 2 + (stacked ? 0 : s * bar);
                    double x = horizontal ? left + Math.Min(start, end) * plotW : offset;
                    double y2 = horizontal ? offset : top + (1 - Math.Max(start, end)) * plotH;
                    double w = horizontal ? Math.Abs(end - start) * plotW : bar;
                    double h = horizontal ? bar : Math.Abs(end - start) * plotH;
                    if (shown[c, s] == 0) Line(layout, x, y2, horizontal ? 0 : w, horizontal ? h : 0, SeriesColors[s]);
                    else Shape(layout, "rect", x, y2, w, h, SeriesColors[s]);
                    if (w >= 30 && h >= 16)
                        Text(layout, Number(shown[c, s]) + (mode == "Percent" ? "%" : ""), x + 2, y2 + (h - 16) / 2, w - 4, 16, 9, 16777215);
                }
            }
        }
    }
}
