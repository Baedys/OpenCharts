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
            ChartAppearance appearance = spec.Appearance ?? new ChartAppearance();
            appearance.Validate(series, categories);
            var offsets = new Dictionary<int, ChartLabelOffset>();
            if (appearance.ValueLabelOffsets != null) foreach (ChartLabelOffset offset in appearance.ValueLabelOffsets) offsets[offset.CategoryIndex * series + offset.SeriesIndex] = offset;
            int textColor = spec.Appearance == null ? Ink : ChartAppearance.ParseColor(appearance.TextColor);
            var order = new List<int>();
            if (appearance.DisplayOrder != null && appearance.DisplayOrder.Count > 0) order.AddRange(appearance.DisplayOrder);
            else for (int s = 0; s < series; s++) order.Add(s);
            int[] colors = new int[series];
            for (int s = 0; s < series; s++) colors[s] = appearance.SeriesColors != null && appearance.SeriesColors.Count > 0 ? ChartAppearance.ParseColor(appearance.SeriesColors[s]) : SeriesColors[s];
            bool horizontal = kind == "bar", stacked = mode != "Clustered";
            if (stacked && appearance.OverlapPercent != 100) throw new NotSupportedException("Only fully overlapping stacks are supported.");
            double[,] starts = new double[categories, series], ends = new double[categories, series];
            double[,] shown = new double[categories, series];
            double[] positiveTops = new double[categories], totals = new double[categories];
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
                foreach (int s in order)
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
                positiveTops[c] = positive;
                totals[c] = positive + negative;
                if (!Finite(totals[c])) throw new ArgumentException("Category total overflow.");
            }
            if (mode == "Percent") high = 100;
            if (appearance.Minimum.HasValue)
            {
                if (low < appearance.Minimum.Value) throw new NotSupportedException("Clipping data below a fixed axis minimum is not supported.");
                low = appearance.Minimum.Value;
            }
            if (appearance.Maximum.HasValue)
            {
                if (high > appearance.Maximum.Value) throw new NotSupportedException("Clipping data above a fixed axis maximum is not supported.");
                high = appearance.Maximum.Value;
            }
            Scale scale = new Scale(low, high);
            int legendCols = Math.Min(series, Math.Max(1, (int)((width - 28) / 135)));
            int legendRows = appearance.ShowLegend ? (series + legendCols - 1) / legendCols : 0;
            double legendCell = (width - 28) / legendCols;
            for (int s = 0; appearance.ShowLegend && s < series; s++)
            {
                double x = 14 + (s % legendCols) * legendCell, y = 42 + (s / legendCols) * 18;
                Shape(layout, "rect", x, y + 3, 9, 9, colors[s]);
                Text(layout, spec.Series[s] ?? ("Series " + (s + 1)), x + 13, y, legendCell - 16, 17, 9, Ink, "left");
            }
            double left = horizontal ? Math.Min(130, width * .28) : 68;
            double top = 50 + legendRows * 18, plotW = width - left - 24, plotH = height - top - 44;
            if (appearance.PlotRect != null)
            {
                left = width * appearance.PlotRect.X; top = height * appearance.PlotRect.Y;
                plotW = width * appearance.PlotRect.Width; plotH = height * appearance.PlotRect.Height;
            }
            if (appearance.PlotRect == null && plotH < 55) throw new ArgumentException("Chart is too short for its legend; increase the height.");
            for (int t = 0; t <= 4; t++)
            {
                double fraction = t / 4.0;
                if (horizontal)
                {
                    double x = left + fraction * plotW;
                    if (appearance.ShowGridlines) Line(layout, x, top, 0, plotH, Gray);
                    if (appearance.ShowValueAxis) Text(layout, Number(scale.Tick(fraction)), Math.Max(0, x - 24), top + plotH + 7, 48, 18, 9, Ink);
                }
                else
                {
                    double y = top + plotH * (1 - fraction);
                    if (appearance.ShowGridlines) Line(layout, left, y, plotW, 0, Gray);
                    if (appearance.ShowValueAxis) Text(layout, Number(scale.Tick(fraction)), 4, y - 9, left - 10, 18, 9, Ink, "right");
                }
            }
            if (horizontal) Line(layout, left + scale.Fraction(0) * plotW, top, 0, plotH, Ink);
            else Line(layout, left, top + plotH * (1 - scale.Fraction(0)), plotW, 0, Ink);
            layout.Items[layout.Items.Count - 1].LineWidth = appearance.AxisLineWidth;
            if (spec.Appearance != null) layout.Items[layout.Items.Count - 1].Color = ChartAppearance.ParseColor(appearance.AxisColor);
            double step = (horizontal ? plotH : plotW) / categories;
            double fullBar = step / (1 + appearance.GapRatio), bar = stacked ? fullBar : fullBar / series;
            double[] previousX = new double[series], previousY = new double[series];
            bool[] connected = new bool[series];
            for (int c = 0; c < categories; c++)
            {
                double categoryFont = spec.Appearance == null ? Font(step) : appearance.CategoryFontSize;
                if (horizontal) Text(layout, spec.Rows[c].Label ?? "", 4, top + c * step, left - 10, step, categoryFont, Ink, "right");
                else
                {
                    double gap = appearance.CategoryLabelGap ?? 4;
                    Text(layout, spec.Rows[c].Label ?? "", left + c * step, top + plotH + gap, step, appearance.CategoryLabelHeight ?? Math.Min(30, height - top - plotH - gap), categoryFont, Ink);
                    if (spec.Appearance != null) ImportedText(layout, "top", 0);
                }
                foreach (int s in order)
                {
                    if (!spec.Rows[c].Values[s].HasValue) { connected[s] = false; continue; }
                    double start = scale.Fraction(starts[c, s]), end = scale.Fraction(ends[c, s]);
                    double center = (horizontal ? top : left) + step * (c + .5);
                    if (kind == "line")
                    {
                        double y = top + plotH * (1 - end);
                        if (connected[s]) Line(layout, previousX[s], previousY[s], center - previousX[s], y - previousY[s], colors[s]);
                        Shape(layout, "ellipse", center - 2.5, y - 2.5, 5, 5, colors[s]);
                        previousX[s] = center; previousY[s] = y; connected[s] = true;
                        continue;
                    }
                    double offset = center - fullBar / 2 + (stacked ? 0 : order.IndexOf(s) * bar);
                    double x = horizontal ? left + Math.Min(start, end) * plotW : offset;
                    double y2 = horizontal ? offset : top + (1 - Math.Max(start, end)) * plotH;
                    double w = horizontal ? Math.Abs(end - start) * plotW : bar;
                    double h = horizontal ? bar : Math.Abs(end - start) * plotH;
                    if (shown[c, s] == 0) Line(layout, x, y2, horizontal ? 0 : w, horizontal ? h : 0, colors[s]);
                    else Shape(layout, "rect", x, y2, w, h, colors[s]);
                    double labelHeight = spec.Appearance == null ? 16 : appearance.ValueFontSize * 1.2;
                    if (w >= 30 && h >= labelHeight)
                    {
                        Text(layout, ChartAppearance.Format(shown[c, s], appearance.ValueFormat, appearance.ValueDigits) + (mode == "Percent" ? "%" : ""), x + 2, y2 + (h - labelHeight) / 2, w - 4, labelHeight, appearance.ValueFontSize, ChartAppearance.ParseColor(appearance.ValueTextColor));
                        ChartLabelOffset labelOffset;
                        if (offsets.TryGetValue(c * series + s, out labelOffset))
                        {
                            Primitive label = layout.Items[layout.Items.Count - 1];
                            label.X += labelOffset.X * (appearance.SourceFrame == null ? 1 : width / appearance.SourceFrame.Width);
                            label.Y += labelOffset.Y * (appearance.SourceFrame == null ? 1 : height / appearance.SourceFrame.Height);
                        }
                    }
                    if (!horizontal && stacked && c == 0 && appearance.SeriesLabelsAtFirstStack)
                    {
                        double labelH = appearance.SeriesLabelHeight ?? appearance.SeriesFontSize * 1.2;
                        Text(layout, spec.Series[s] ?? "", 0, y2 + (h - labelH) / 2, Math.Max(0, x - (appearance.SeriesLabelGap ?? 10)), labelH, appearance.SeriesFontSize, Ink, "right");
                        ImportedText(layout, "center", 0);
                    }
                }
                if (!horizontal && stacked && appearance.ShowTotals)
                {
                    double y = top + plotH * (1 - scale.Fraction(positiveTops[c]));
                    double labelH = appearance.TotalLabelHeight ?? appearance.TotalFontSize * 1.2;
                    Text(layout, ChartAppearance.Format(totals[c], appearance.TotalFormat, appearance.ValueDigits), left + c * step, Math.Max(0, y - labelH - (appearance.TotalLabelGap ?? 3)), step, labelH, appearance.TotalFontSize, Ink);
                    ImportedText(layout, "bottom", 2);
                }
            }
            if (spec.Appearance != null)
                foreach (Primitive item in layout.Items) if (item.Kind == "text") { item.FontName = appearance.FontName; if (item.Color == Ink) item.Color = textColor; }
        }
        private static void ImportedText(Layout layout, string anchor, double margin)
        {
            Primitive text = layout.Items[layout.Items.Count - 1];
            text.VerticalAlignment = anchor; text.MarginLeft = margin; text.MarginRight = margin; text.NoWrap = true;
        }
    }
}
