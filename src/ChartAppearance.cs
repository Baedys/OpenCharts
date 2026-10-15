using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenSlideCharts
{
    public sealed class ChartRectangle
    {
        public double X, Y, Width, Height;
    }
    public sealed class ChartLabelOffset
    {
        public int SeriesIndex, CategoryIndex;
        public double X, Y;
    }

    // Optional observed formatting. Coordinates of PlotRect are fractions of
    // SourceFrame. SourceFrame is in PowerPoint points, not screen pixels.
    public sealed class ChartAppearance
    {
        public List<string> SeriesColors = new List<string>();
        public List<int> DisplayOrder = new List<int>();
        public List<ChartLabelOffset> ValueLabelOffsets = new List<ChartLabelOffset>();
        public bool ShowTitle = true, ShowLegend = true, ShowValueAxis = true, ShowGridlines = true;
        public bool ShowTotals, SeriesLabelsAtFirstStack;
        public string ValueFormat = "General", TotalFormat = "General", FontName = "Aptos";
        public string TextColor = "#000000", AxisColor = "#000000", ValueTextColor = "#FFFFFF";
        public int ValueDigits = -1;
        public double ValueFontSize = 9, CategoryFontSize = 11, SeriesFontSize = 11, TotalFontSize = 11;
        public double GapRatio = 3.0 / 7, OverlapPercent = 100, AxisLineWidth = 1.3;
        public ChartRectangle PlotRect, SourceFrame;
        public double? Minimum, Maximum;
        public double? SeriesLabelGap, CategoryLabelGap, TotalLabelGap;
        public double? SeriesLabelHeight, CategoryLabelHeight, TotalLabelHeight;

        internal static bool IsFinite(double v) { return !Double.IsNaN(v) && !Double.IsInfinity(v); }
        internal void Validate(int series, int categories)
        {
            foreach (double size in new double[] { ValueFontSize, CategoryFontSize, SeriesFontSize, TotalFontSize })
                if (!IsFinite(size) || size < 2 || size > 96) throw new ArgumentException("Appearance font sizes must be between 2 and 96 points.");
            if (!IsFinite(GapRatio) || GapRatio < 0 || GapRatio > 5) throw new ArgumentException("GapRatio must be between 0 and 5.");
            if (!IsFinite(OverlapPercent) || OverlapPercent < -100 || OverlapPercent > 100) throw new ArgumentException("OverlapPercent must be between -100 and 100.");
            if (!IsFinite(AxisLineWidth) || AxisLineWidth <= 0 || AxisLineWidth > 10) throw new ArgumentException("AxisLineWidth must be between 0 and 10 points.");
            if (ValueDigits < -1 || ValueDigits > 12) throw new ArgumentException("ValueDigits must be between -1 and 12.");
            ParseColor(TextColor); ParseColor(AxisColor); ParseColor(ValueTextColor);
            foreach (double? length in new double?[] { SeriesLabelGap, CategoryLabelGap, TotalLabelGap, SeriesLabelHeight, CategoryLabelHeight, TotalLabelHeight })
                if (length.HasValue && (!IsFinite(length.Value) || length.Value < 0 || length.Value > 10000)) throw new ArgumentException("Label lengths must be finite, nonnegative and at most 10000 points.");
            if (Minimum.HasValue && !IsFinite(Minimum.Value) || Maximum.HasValue && !IsFinite(Maximum.Value)) throw new ArgumentException("Axis bounds must be finite.");
            if (Minimum.HasValue && Maximum.HasValue && Minimum.Value >= Maximum.Value) throw new ArgumentException("Axis minimum must be below maximum.");
            if (DisplayOrder != null && DisplayOrder.Count > 0)
            {
                var seen = new HashSet<int>();
                foreach (int index in DisplayOrder) if (index < 0 || index >= series || !seen.Add(index)) throw new ArgumentException("DisplayOrder must be a permutation of the series indices.");
                if (seen.Count != series) throw new ArgumentException("DisplayOrder must include each series exactly once.");
            }
            if (SeriesColors != null && SeriesColors.Count > 0)
            {
                if (SeriesColors.Count != series) throw new ArgumentException("Appearance requires one color per series.");
                foreach (string color in SeriesColors) ParseColor(color);
            }
            if (ValueLabelOffsets != null)
            {
                var seenOffsets = new HashSet<int>();
                foreach (ChartLabelOffset offset in ValueLabelOffsets)
                {
                    if (offset == null || offset.SeriesIndex < 0 || offset.SeriesIndex >= series || offset.CategoryIndex < 0 || offset.CategoryIndex >= categories ||
                        !IsFinite(offset.X) || !IsFinite(offset.Y) || Math.Abs(offset.X) > 100000 || Math.Abs(offset.Y) > 100000 ||
                        !seenOffsets.Add(offset.CategoryIndex * series + offset.SeriesIndex))
                        throw new ArgumentException("Value label offsets must identify unique cells with finite positions.");
                }
            }
            if (PlotRect != null)
            {
                Rectangle(PlotRect);
                if (PlotRect.X < 0 || PlotRect.Y < 0 || PlotRect.Width <= 0 || PlotRect.Height <= 0 || PlotRect.X + PlotRect.Width > 1.000001 || PlotRect.Y + PlotRect.Height > 1.000001)
                    throw new ArgumentException("PlotRect must fit inside the chart frame.");
            }
            if (SourceFrame != null)
            {
                Rectangle(SourceFrame);
                if (SourceFrame.Width <= 0 || SourceFrame.Height <= 0) throw new ArgumentException("SourceFrame must have positive dimensions.");
            }
        }
        private static void Rectangle(ChartRectangle r)
        {
            if (!IsFinite(r.X) || !IsFinite(r.Y) || !IsFinite(r.Width) || !IsFinite(r.Height)) throw new ArgumentException("Appearance rectangles must be finite.");
        }
        internal static int ParseColor(string color)
        {
            int rgb;
            if (color == null || color.Length != 7 || color[0] != '#' || !Int32.TryParse(color.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out rgb))
                throw new ArgumentException("Series colors must use #RRGGBB.");
            return ((rgb >> 16) & 255) | (rgb & 65280) | ((rgb & 255) << 16);
        }
        internal static string Format(double value, string format, int digits)
        {
            // Only the formats independently supported by this renderer. Unknown
            // Excel formats fail explicitly instead of silently changing meaning.
            if (format == "General" || String.IsNullOrEmpty(format))
                return digits >= 0 ? value.ToString("F" + digits, CultureInfo.InvariantCulture) : value.ToString("G5", CultureInfo.InvariantCulture);
            if (format == "#,##0;\"-\"#,##0" || format == "#,##0;-#,##0" || format == "#,##0")
                return Math.Round(value, 0, MidpointRounding.AwayFromZero).ToString("#,##0;-#,##0;0", CultureInfo.InvariantCulture);
            if (format == "0") return Math.Round(value, 0, MidpointRounding.AwayFromZero).ToString("0", CultureInfo.InvariantCulture);
            throw new NotSupportedException("Unsupported observed number format: " + format);
        }
    }
}
