using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OpenSlideCharts
{
    public sealed class PieOptions
    {
        public double StartAngle = -90;
        public double? HoleRatio;
        public string Order = "SheetClockwise";
        public string LabelContent = "Percent";
        public bool ShowCategory = true;
        public int PercentDigits = 0;
        public string CenterTitle;
        public List<PieSliceStyle> Slices = new List<PieSliceStyle>();
    }

    public sealed class PieSliceStyle
    {
        public int Index;
        public double Explode;
        public string Color;
    }

    // Coordinates are chart-local points, including cubic control points.
    // Each shape is one explicitly closed contour. A doughnut traverses the
    // outer arc clockwise and the inner arc counterclockwise, joined radially.
    public sealed class ShapePath
    {
        public double StartX, StartY;
        public List<PathSegment> Segments = new List<PathSegment>();
    }

    public sealed class PathSegment
    {
        public bool Curve;
        public double X, Y, C1X, C1Y, C2X, C2Y;
    }

    public static partial class ChartEngine
    {
        private sealed class PiePart
        {
            internal int Index, Color;
            internal double Share, Start, Sweep, Middle, Explode, CenterX, CenterY, AnchorX, AnchorY;
            internal string Label, Category, Numeric;
        }

        private static double EffectiveHole(ChartSpec spec, PieOptions options)
        {
            return options.HoleRatio ?? (String.Equals(spec.Kind, "Doughnut", StringComparison.OrdinalIgnoreCase) ? .5 : 0);
        }

        // Nonnegative input and defaults below are original policies, not claims
        // about the reference product's behavior for negative data or defaults.
        internal static void ValidatePie(ChartSpec spec)
        {
            if (spec == null) throw new ArgumentNullException("spec");
            if (!String.Equals(spec.Kind, "Pie", StringComparison.OrdinalIgnoreCase) && !String.Equals(spec.Kind, "Doughnut", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Pie validation requires Kind=Pie or Doughnut.");
            if (spec.Appearance != null || spec.Waterfall != null || spec.Series != null && spec.Series.Count != 0)
                throw new NotSupportedException("Pie charts require scalar rows without a series matrix, waterfall or imported appearance.");
            if (spec.LayoutMode != null && spec.LayoutMode != "Clustered")
                throw new NotSupportedException("Pie charts do not use stacked or percent LayoutMode.");
            if (spec.Rows == null || spec.Rows.Count < 1 || spec.Rows.Count > 80)
                throw new ArgumentException("Pie charts require between 1 and 80 rows.");
            bool positive = false;
            foreach (ChartRow row in spec.Rows)
            {
                if (row == null || !Finite(row.Value) || row.Value < 0 || !Finite(row.Secondary))
                    throw new ArgumentException("Pie values must be finite and nonnegative.");
                if (row.Secondary != 0 || row.Total || row.Values != null && row.Values.Count != 0)
                    throw new NotSupportedException("Pie rows use Value only; secondary, total and series values are ambiguous.");
                positive |= row.Value > 0;
            }
            if (!positive) throw new ArgumentException("Pie charts require at least one positive value.");
            PieOptions p = spec.Pie ?? new PieOptions();
            if (!Finite(p.StartAngle)) throw new ArgumentException("Pie StartAngle must be finite.");
            double hole = EffectiveHole(spec, p);
            if (!Finite(hole) || hole < 0 || hole >= 1) throw new ArgumentException("Pie HoleRatio must be in [0,1).");
            if (p.Order != "SheetClockwise" && p.Order != "SheetCounterclockwise" && p.Order != "Ascending" && p.Order != "Descending")
                throw new ArgumentException("Unknown pie slice order.");
            if (p.LabelContent != "Percent" && p.LabelContent != "Value" && p.LabelContent != "Both" && p.LabelContent != "None")
                throw new ArgumentException("Unknown pie label content.");
            if (p.PercentDigits < 0 || p.PercentDigits > 6) throw new ArgumentException("Pie PercentDigits must be between 0 and 6.");
            if (!String.IsNullOrEmpty(p.CenterTitle) && hole == 0)
                throw new ArgumentException("A center title requires a positive hole ratio.");
            HashSet<int> indexes = new HashSet<int>();
            if (p.Slices != null) foreach (PieSliceStyle style in p.Slices)
            {
                if (style == null || style.Index < 0 || style.Index >= spec.Rows.Count || !indexes.Add(style.Index))
                    throw new ArgumentException("Pie slice styles require unique valid data indexes.");
                if (!Finite(style.Explode) || style.Explode < 0 || style.Explode > .5)
                    throw new ArgumentException("Pie slice Explode must be between 0 and 0.5.");
                if (style.Color != null) PieColor(style.Color);
            }
        }

        private static int PieColor(string value)
        {
            if (value.Length != 7 || value[0] != '#') throw new ArgumentException("Pie colors must use #RRGGBB.");
            for (int i = 1; i < value.Length; i++)
                if (!(value[i] >= '0' && value[i] <= '9' || value[i] >= 'a' && value[i] <= 'f' || value[i] >= 'A' && value[i] <= 'F'))
                    throw new ArgumentException("Pie colors must use #RRGGBB.");
            return ChartAppearance.ParseColor(value);
        }

        private static void BuildPie(Layout layout, ChartSpec spec, double width, double height)
        {
            PieOptions p = spec.Pie ?? new PieOptions();
            double hole = EffectiveHole(spec, p), maximum = 0;
            foreach (ChartRow row in spec.Rows) maximum = Math.Max(maximum, row.Value);
            Dictionary<int, PieSliceStyle> styles = new Dictionary<int, PieSliceStyle>();
            double maxExplode = 0;
            if (p.Slices != null) foreach (PieSliceStyle s in p.Slices)
            {
                styles.Add(s.Index, s);
                if (spec.Rows[s.Index].Value > 0) maxExplode = Math.Max(maxExplode, s.Explode);
            }
            List<int> order = new List<int>();
            for (int i = 0; i < spec.Rows.Count; i++) if (spec.Rows[i].Value > 0) order.Add(i);
            if (p.Order == "Ascending" || p.Order == "Descending") order.Sort(delegate(int a, int b)
            {
                int comparison = spec.Rows[a].Value.CompareTo(spec.Rows[b].Value);
                if (p.Order == "Descending") comparison = -comparison;
                return comparison == 0 ? a.CompareTo(b) : comparison;
            });
            // Normalize before addition: even eighty Double.MaxValue values cannot overflow.
            double total = 0, correction = 0;
            foreach (int i in order) AddCompensated(ref total, ref correction, spec.Rows[i].Value / maximum);
            bool labels = p.ShowCategory || p.LabelContent != "None";
            double lane = labels ? Math.Min(180, width * .25) : 0;
            double top = 48, bottom = height - 10, cx = width / 2, cy = (top + bottom) / 2;
            double radius = Math.Min(width / 2 - lane - 20, (bottom - top) / 2 - 12) / (1 + maxExplode);
            double rotation = (p.StartAngle % 360) * Math.PI / 180;
            double direction = p.Order == "SheetCounterclockwise" ? -1 : 1;
            List<PiePart> parts = new List<PiePart>();
            double cumulative = 0, compensation = 0, previous = 0;
            for (int n = 0; n < order.Count; n++)
            {
                int index = order[n];
                double normalized = spec.Rows[index].Value / maximum;
                AddCompensated(ref cumulative, ref compensation, normalized);
                double end = n == order.Count - 1 ? Math.PI * 2 : Math.Min(Math.PI * 2, cumulative / total * (Math.PI * 2));
                double sweep = Math.Max(0, end - previous);
                double start = direction > 0 ? rotation + previous : rotation - end;
                previous = end;
                // Contributions below floating-point angular resolution have no drawable area.
                // They are never inflated into a visible slice or divided by zero.
                if (normalized == 0 || sweep == 0 || start + sweep == start) continue;
                PieSliceStyle style;
                styles.TryGetValue(index, out style);
                PiePart part = new PiePart { Index = index, Share = normalized / total, Start = start, Sweep = sweep,
                    Middle = start + sweep / 2, Explode = style == null ? 0 : style.Explode,
                    Color = style != null && style.Color != null ? PieColor(style.Color) : SeriesColors[index % SeriesColors.Length] };
                part.CenterX = cx + Math.Cos(part.Middle) * radius * part.Explode;
                part.CenterY = cy + Math.Sin(part.Middle) * radius * part.Explode;
                part.AnchorX = part.CenterX + Math.Cos(part.Middle) * radius;
                part.AnchorY = part.CenterY + Math.Sin(part.Middle) * radius;
                part.Numeric = PieNumeric(spec.Rows[index], part.Share, p);
                part.Category = p.ShowCategory ? CleanPieText(spec.Rows[index].Label) : "";
                part.Label = part.Category + part.Numeric;
                parts.Add(part);
            }
            // Label fit is checked before emitting paths so failure cannot yield a partial layout.
            List<PiePart> left = new List<PiePart>(), right = new List<PiePart>();
            foreach (PiePart part in parts) if (part.Label.Length != 0)
            {
                if (Math.Cos(part.Middle) < 0) left.Add(part); else right.Add(part);
            }
            double leftFont = PieLabelFont(left, bottom - top, lane), rightFont = PieLabelFont(right, bottom - top, lane);
            foreach (PiePart part in parts)
            {
                ShapePath path = PiePath(part.CenterX, part.CenterY, radius, hole, part.Start, part.Sweep);
                Primitive shape = new Primitive { Kind = "path", Path = path, Color = part.Color,
                    Role = "PIE_SLICE", EntityId = part.Index.ToString(Invariant), LineWidth = 0 };
                PathBounds(shape);
                layout.Items.Add(shape);
            }
            PieLabels(layout, left, true, lane, width, top, bottom, cx, radius * (1 + maxExplode), leftFont);
            PieLabels(layout, right, false, lane, width, top, bottom, cx, radius * (1 + maxExplode), rightFont);
            if (!String.IsNullOrEmpty(p.CenterTitle))
            {
                // An inscribed square stays inside the central hole. Exploded slices may have
                // displaced hole centers, so reduce to the circle shared by every inner hole.
                double safeHole = radius * Math.Max(0, hole - maxExplode);
                double square = safeHole * Math.Sqrt(2) * .90;
                if (square < 10) throw new ArgumentException("Pie center title has insufficient space; increase hole or chart size, or reduce explosion.");
                double font = Math.Min(14, square / 1.4);
                string title = FitPieText(CleanPieText(p.CenterTitle), square, font);
                layout.Items.Add(new Primitive { Kind = "text", Role = "PIE_CENTER_TITLE", Text = title,
                    X = cx - square / 2, Y = cy - font * .7, W = square, H = font * 1.4,
                    FontSize = font, Color = Ink, Alignment = "center", VerticalAlignment = "center", NoWrap = true,
                    MarginLeft = 0, MarginRight = 0 });
            }
        }

        private static void AddCompensated(ref double sum, ref double correction, double value)
        {
            double adjusted = value - correction, next = sum + adjusted;
            correction = (next - sum) - adjusted;
            sum = next;
        }

        private static string PieNumeric(ChartRow row, double share, PieOptions p)
        {
            string numeric = "";
            if (p.LabelContent == "Value" || p.LabelContent == "Both") numeric = Number(row.Value);
            if (p.LabelContent == "Percent" || p.LabelContent == "Both")
            {
                string percent = (share * 100).ToString("F" + p.PercentDigits, Invariant) + "%";
                numeric += numeric.Length == 0 ? percent : " (" + percent + ")";
            }
            return numeric;
        }

        private static double PieLabelFont(List<PiePart> parts, double available, double lane)
        {
            int count = parts.Count;
            if (count == 0) return 11;
            double font = Math.Min(11, (available - (count - 1) * 2) / (count * 1.4));
            foreach (PiePart part in parts)
            {
                // Reserve every numeric character. Category text may be abbreviated,
                // but neither an exponent nor a percent sign is silently dropped.
                int minimum = part.Numeric.Length + (part.Category.Length == 0 ? 0 : part.Numeric.Length == 0 ? 1 : 2);
                if (minimum > 0) font = Math.Min(font, lane / (minimum * 1.2));
            }
            if (font < 6) throw new ArgumentException("Pie labels do not fit this chart frame at 6 points. Increase chart width/height or hide labels.");
            return font;
        }

        private static string FitPieText(string text, double width, double font)
        {
            // Conservative width budget without a dependency on installed fonts or Office.
            int characters = Math.Max(1, (int)Math.Floor(width / (font * 1.2) + 1e-10));
            if (text.Length <= characters) return text;
            int prefix = characters - 1;
            if (prefix > 0 && Char.IsHighSurrogate(text[prefix - 1])) prefix--;
            return text.Substring(0, prefix) + "\u2026";
        }

        private static string CleanPieText(string text)
        {
            if (String.IsNullOrEmpty(text)) return "";
            StringBuilder clean = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (Char.IsControl(c) || c == '\u2028' || c == '\u2029') clean.Append(' ');
                else if (Char.IsHighSurrogate(c))
                {
                    if (i + 1 < text.Length && Char.IsLowSurrogate(text[i + 1])) { clean.Append(c); clean.Append(text[++i]); }
                    else clean.Append('\ufffd');
                }
                else clean.Append(Char.IsLowSurrogate(c) ? '\ufffd' : c);
            }
            return clean.ToString();
        }

        private static void PieLabels(Layout layout, List<PiePart> parts, bool left, double lane, double width,
            double top, double bottom, double cx, double outer, double font)
        {
            if (parts.Count == 0) return;
            parts.Sort(delegate(PiePart a, PiePart b)
            {
                int comparison = a.AnchorY.CompareTo(b.AnchorY);
                return comparison == 0 ? a.Index.CompareTo(b.Index) : comparison;
            });
            double h = font * 1.4, spacing = h + 2;
            double[] centers = new double[parts.Count];
            for (int i = 0; i < parts.Count; i++)
                centers[i] = Math.Max(top + h / 2, Math.Max(parts[i].AnchorY, i == 0 ? top + h / 2 : centers[i - 1] + spacing));
            centers[centers.Length - 1] = Math.Min(centers[centers.Length - 1], bottom - h / 2);
            for (int i = centers.Length - 2; i >= 0; i--) centers[i] = Math.Min(centers[i], centers[i + 1] - spacing);
            for (int i = 0; i < parts.Count; i++)
            {
                PiePart part = parts[i];
                string id = part.Index.ToString(Invariant);
                double x = left ? 8 : width - 8 - lane, edge = left ? x + lane + 2 : x - 2;
                double elbow = cx + (left ? -1 : 1) * (outer + 6);
                layout.Items.Add(new Primitive { Kind = "line", Role = "PIE_LEADER", EntityId = id,
                    X = part.AnchorX, Y = part.AnchorY, W = elbow - part.AnchorX, H = centers[i] - part.AnchorY, Color = Gray, LineWidth = .75 });
                layout.Items.Add(new Primitive { Kind = "line", Role = "PIE_LEADER", EntityId = id,
                    X = elbow, Y = centers[i], W = edge - elbow, H = 0, Color = Gray, LineWidth = .75 });
                string label = part.Numeric;
                if (part.Category.Length != 0)
                {
                    double categoryWidth = lane - (part.Numeric.Length == 0 ? 0 : (part.Numeric.Length + 1) * font * 1.2);
                    string category = FitPieText(part.Category, categoryWidth, font);
                    label = category + (part.Numeric.Length == 0 ? "" : " " + part.Numeric);
                }
                layout.Items.Add(new Primitive { Kind = "text", Role = "PIE_LABEL", EntityId = id,
                    X = x, Y = centers[i] - h / 2, W = lane, H = h, Text = label,
                    FontSize = font, Color = Ink, Alignment = left ? "right" : "left", VerticalAlignment = "center",
                    NoWrap = true, MarginLeft = 0, MarginRight = 0 });
            }
        }

        private static ShapePath PiePath(double cx, double cy, double radius, double hole, double start, double sweep)
        {
            ShapePath path = new ShapePath { StartX = cx + radius * Math.Cos(start), StartY = cy + radius * Math.Sin(start) };
            AppendPieArc(path, cx, cy, radius, start, sweep);
            if (hole > 0)
            {
                path.Segments.Add(new PathSegment { X = cx + radius * hole * Math.Cos(start + sweep), Y = cy + radius * hole * Math.Sin(start + sweep) });
                AppendPieArc(path, cx, cy, radius * hole, start + sweep, -sweep);
            }
            else path.Segments.Add(new PathSegment { X = cx, Y = cy });
            path.Segments.Add(new PathSegment { X = path.StartX, Y = path.StartY });
            return path;
        }

        private static void AppendPieArc(ShapePath path, double cx, double cy, double radius, double start, double sweep)
        {
            double direction = sweep < 0 ? -1 : 1, quarter = Math.PI / 2;
            double end = start + sweep, a = start;
            // Split at cardinal angles as well as the final angle. Control points
            // then remain inside the circle's bounding square for any rotation.
            for (int i = 0; i < 12 && direction * (end - a) > 0; i++)
            {
                double boundary = direction > 0 ? (Math.Floor(a / quarter) + 1) * quarter : (Math.Ceiling(a / quarter) - 1) * quarter;
                if (Math.Abs(boundary - a) < 1e-12) boundary += direction * quarter;
                double b = direction > 0 ? Math.Min(end, boundary) : Math.Max(end, boundary);
                double tangent = 4.0 / 3 * Math.Tan((b - a) / 4);
                path.Segments.Add(new PathSegment { Curve = true,
                    C1X = cx + radius * (Math.Cos(a) - tangent * Math.Sin(a)),
                    C1Y = cy + radius * (Math.Sin(a) + tangent * Math.Cos(a)),
                    C2X = cx + radius * (Math.Cos(b) + tangent * Math.Sin(b)),
                    C2Y = cy + radius * (Math.Sin(b) - tangent * Math.Cos(b)),
                    X = cx + radius * Math.Cos(b), Y = cy + radius * Math.Sin(b) });
                a = b;
            }
        }

        private static void PathBounds(Primitive primitive)
        {
            ShapePath path = primitive.Path;
            double x1 = path.StartX, y1 = path.StartY, x2 = x1, y2 = y1;
            foreach (PathSegment s in path.Segments)
            {
                x1 = Math.Min(x1, s.X); y1 = Math.Min(y1, s.Y); x2 = Math.Max(x2, s.X); y2 = Math.Max(y2, s.Y);
                if (s.Curve)
                {
                    x1 = Math.Min(x1, Math.Min(s.C1X, s.C2X)); y1 = Math.Min(y1, Math.Min(s.C1Y, s.C2Y));
                    x2 = Math.Max(x2, Math.Max(s.C1X, s.C2X)); y2 = Math.Max(y2, Math.Max(s.C1Y, s.C2Y));
                }
            }
            primitive.X = x1; primitive.Y = y1; primitive.W = x2 - x1; primitive.H = y2 - y1;
        }

        private static void ValidatePathFrame(ShapePath path, double width, double height)
        {
            ValidatePathPoint(path.StartX, path.StartY, width, height);
            if (path.Segments == null || path.Segments.Count == 0) throw new ArgumentException("Shape paths must have segments.");
            foreach (PathSegment s in path.Segments)
            {
                if (s == null) throw new ArgumentException("Shape paths must have valid segments.");
                ValidatePathPoint(s.X, s.Y, width, height);
                if (s.Curve) { ValidatePathPoint(s.C1X, s.C1Y, width, height); ValidatePathPoint(s.C2X, s.C2Y, width, height); }
            }
            PathSegment last = path.Segments[path.Segments.Count - 1];
            if (last.X != path.StartX || last.Y != path.StartY) throw new ArgumentException("Shape paths must be explicitly closed.");
        }

        private static void ValidatePathPoint(double x, double y, double width, double height)
        {
            if (!Finite(x) || !Finite(y) || x < -1e-5 || y < -1e-5 || x > width + 1e-5 || y > height + 1e-5)
                throw new ArgumentException("Shape path exceeds chart frame.");
        }
    }
}
