using System;
using System.Collections.Generic;
using System.Globalization;
using System.Numerics;

namespace OpenSlideCharts
{
    public static class WaterfallSolver
    {
        public static WaterfallSolution Solve(WaterfallSpec spec)
        {
            return new Solver(spec).Run();
        }

        private sealed class Solver
        {
            private readonly WaterfallSpec spec;
            private readonly Dictionary<string, int> columns = new Dictionary<string, int>(StringComparer.Ordinal);
            private readonly List<string> variables = new List<string>();
            private readonly List<int> bases = new List<int>();
            private readonly List<int> equals = new List<int>();
            private readonly List<Rational[]> values = new List<Rational[]>();

            internal Solver(WaterfallSpec input)
            {
                spec = input;
                Validate();
            }

            private void Validate()
            {
                if (spec == null || spec.Columns == null || spec.Columns.Count < 1 || spec.Columns.Count > 80)
                    throw new ArgumentException("A waterfall requires 1 to 80 columns.");
                if (spec.Connectors == null || spec.Pins == null || spec.Connectors.Count > 1024 || spec.Pins.Count > 1024)
                    throw new ArgumentException("Connector and pin collections are required, with at most 1024 entries each.");
                if (spec.SegmentOrder != "SheetTopToBottom" && spec.SegmentOrder != "SheetBottomToTop")
                    throw new ArgumentException("SegmentOrder must be SheetTopToBottom or SheetBottomToTop.");
                HashSet<string> segmentIds = new HashSet<string>(StringComparer.Ordinal);
                for (int i = 0; i < spec.Columns.Count; i++)
                {
                    WaterfallColumn column = spec.Columns[i];
                    if (column == null) throw new ArgumentException("Waterfall columns cannot be null.");
                    Id(column.Id, "Column");
                    if (columns.ContainsKey(column.Id)) throw new ArgumentException("Duplicate column ID: " + column.Id);
                    columns.Add(column.Id, i);
                    if (column.Segments == null || column.Segments.Count < 1 || column.Segments.Count > 8)
                        throw new ArgumentException("Each waterfall column requires 1 to 8 segments.");
                    bases.Add(variables.Count);
                    variables.Add(column.Id + ":base");
                    int calculated = -1;
                    HashSet<int> sheetRows = new HashSet<int>();
                    Rational[] segmentValues = new Rational[column.Segments.Count];
                    for (int j = 0; j < column.Segments.Count; j++)
                    {
                        WaterfallSegment segment = column.Segments[j];
                        if (segment == null) throw new ArgumentException("Waterfall segments cannot be null.");
                        Id(segment.Id, "Segment");
                        if (!segmentIds.Add(segment.Id)) throw new ArgumentException("Duplicate segment ID: " + segment.Id);
                        if (segment.SheetRow < -1 || segment.SheetRow > 7)
                            throw new ArgumentException("Segment SheetRow must be -1 or a row index from 0 to 7.");
                        int effectiveRow = segment.SheetRow >= 0 ? segment.SheetRow : j;
                        if (!sheetRows.Add(effectiveRow))
                            throw new ArgumentException("Effective segment sheet rows must be unique within each column, including legacy list positions.");
                        if (segment.Calculated)
                        {
                            if (calculated >= 0) throw new ArgumentException("At most one calculated e is allowed per column.");
                            if (segment.Value.HasValue) throw new ArgumentException("A calculated e must not also contain a numeric value.");
                            calculated = variables.Count;
                            variables.Add(column.Id + ":e");
                            segmentValues[j] = Rational.Zero;
                        }
                        else
                        {
                            if (!segment.Value.HasValue) throw new ArgumentException("A numeric segment requires a value.");
                            segmentValues[j] = Rational.FromDouble(segment.Value.Value);
                        }
                    }
                    equals.Add(calculated);
                    values.Add(segmentValues);
                }
                HashSet<string> connectorIds = new HashSet<string>(StringComparer.Ordinal);
                foreach (WaterfallConnector connector in spec.Connectors)
                {
                    if (connector == null) throw new ArgumentException("Connectors cannot be null.");
                    Id(connector.Id, "Connector");
                    if (!connectorIds.Add(connector.Id)) throw new ArgumentException("Duplicate connector ID: " + connector.Id);
                    ValidateAnchor(connector.Left);
                    ValidateAnchor(connector.Right);
                }
                HashSet<string> pinIds = new HashSet<string>(StringComparer.Ordinal);
                foreach (WaterfallPin pin in spec.Pins)
                {
                    if (pin == null) throw new ArgumentException("Pins cannot be null.");
                    Id(pin.Id, "Pin");
                    if (!pinIds.Add(pin.Id)) throw new ArgumentException("Duplicate pin ID: " + pin.Id);
                    ValidateAnchor(pin.Anchor);
                    Rational.FromDouble(pin.Level);
                }
            }

            private static void Id(string value, string kind)
            {
                if (String.IsNullOrWhiteSpace(value) || value.Length > 256)
                    throw new ArgumentException(kind + " ID must contain 1 to 256 characters.");
            }

            private void ValidateAnchor(WaterfallAnchor anchor)
            {
                if (anchor == null || anchor.ColumnId == null || !columns.ContainsKey(anchor.ColumnId))
                    throw new ArgumentException("Every anchor must reference an existing column ID.");
                if (anchor.Weights == null || anchor.Weights.Count != spec.Columns[columns[anchor.ColumnId]].Segments.Count)
                    throw new ArgumentException("Anchor weight count must match its column's segment count.");
                foreach (double weight in anchor.Weights) Rational.FromDouble(weight);
            }

            private Rational[] Expression(WaterfallAnchor anchor)
            {
                Rational[] expression = Zeros(variables.Count + 1);
                int columnIndex = columns[anchor.ColumnId];
                WaterfallColumn column = spec.Columns[columnIndex];
                expression[bases[columnIndex]] = Rational.One;
                for (int j = 0; j < column.Segments.Count; j++)
                {
                    Rational weight = Rational.FromDouble(anchor.Weights[j]);
                    if (column.Segments[j].Calculated) expression[equals[columnIndex]] += weight;
                    else expression[variables.Count] += weight * values[columnIndex][j];
                }
                return expression;
            }

            internal WaterfallSolution Run()
            {
                List<Rational[]> matrix = new List<Rational[]>();
                foreach (WaterfallPin pin in spec.Pins)
                {
                    Rational[] row = Expression(pin.Anchor);
                    row[variables.Count] = Rational.FromDouble(pin.Level) - row[variables.Count];
                    matrix.Add(row);
                }
                foreach (WaterfallConnector connector in spec.Connectors)
                {
                    if (!connector.Constrain) continue;
                    Rational[] left = Expression(connector.Left), right = Expression(connector.Right);
                    for (int i = 0; i < variables.Count; i++) left[i] -= right[i];
                    left[variables.Count] = right[variables.Count] - left[variables.Count];
                    matrix.Add(left);
                }
                Rational[] answer = Eliminate(matrix);
                WaterfallSolution result = new WaterfallSolution();
                for (int i = 0; i < spec.Columns.Count; i++) result.Columns.Add(RenderColumn(i, answer));
                foreach (WaterfallConnector connector in spec.Connectors)
                {
                    Rational left = Coordinate(connector.Left, answer), right = Coordinate(connector.Right, answer);
                    Rational residual = right - left;
                    result.Connectors.Add(new WaterfallConnectorSolution
                    {
                        Id = connector.Id, LeftColumnId = connector.Left.ColumnId, RightColumnId = connector.Right.ColumnId,
                        LeftLevel = left.ToDouble(), RightLevel = right.ToDouble(), Residual = residual.ToDouble(),
                        Skew = !residual.IsZero, Constrained = connector.Constrain
                    });
                }
                return result;
            }

            private Rational[] Eliminate(List<Rational[]> matrix)
            {
                int count = variables.Count, rowIndex = 0;
                List<int> pivots = new List<int>();
                for (int column = 0; column < count; column++)
                {
                    int pivot = rowIndex;
                    while (pivot < matrix.Count && matrix[pivot][column].IsZero) pivot++;
                    if (pivot == matrix.Count) continue;
                    Rational[] swap = matrix[rowIndex]; matrix[rowIndex] = matrix[pivot]; matrix[pivot] = swap;
                    Rational divisor = matrix[rowIndex][column];
                    for (int j = column; j <= count; j++) matrix[rowIndex][j] /= divisor;
                    for (int i = 0; i < matrix.Count; i++)
                    {
                        if (i == rowIndex || matrix[i][column].IsZero) continue;
                        Rational factor = matrix[i][column];
                        for (int j = column; j <= count; j++) matrix[i][j] -= factor * matrix[rowIndex][j];
                    }
                    pivots.Add(column);
                    rowIndex++;
                }
                foreach (Rational[] row in matrix)
                {
                    bool empty = true;
                    for (int j = 0; j < count; j++) if (!row[j].IsZero) { empty = false; break; }
                    if (empty && !row[count].IsZero)
                        throw new WaterfallSolveException(WaterfallSolveFailure.Inconsistent,
                            "Conflicting waterfall anchors: no all-horizontal solution exists.", null);
                }
                if (pivots.Count != count)
                {
                    List<string> free = new List<string>();
                    for (int j = 0; j < count; j++) if (!pivots.Contains(j)) free.Add(variables[j]);
                    throw new WaterfallSolveException(WaterfallSolveFailure.Underdetermined,
                        "Waterfall requires additional anchors: " + String.Join(", ", free), free);
                }
                Rational[] result = Zeros(count);
                for (int i = 0; i < pivots.Count; i++) result[pivots[i]] = matrix[i][count];
                return result;
            }

            private Rational Coordinate(WaterfallAnchor anchor, Rational[] answer)
            {
                Rational[] expression = Expression(anchor);
                Rational coordinate = expression[variables.Count];
                for (int i = 0; i < variables.Count; i++) coordinate += expression[i] * answer[i];
                return coordinate;
            }

            private WaterfallColumnSolution RenderColumn(int i, Rational[] answer)
            {
                WaterfallColumn column = spec.Columns[i];
                Rational start = answer[bases[i]], sum = start;
                Rational[] actual = new Rational[column.Segments.Count];
                WaterfallSegmentSolution[] segments = new WaterfallSegmentSolution[column.Segments.Count];
                List<int> sheetOrder = new List<int>();
                for (int j = 0; j < actual.Length; j++)
                {
                    sheetOrder.Add(j);
                    actual[j] = column.Segments[j].Calculated ? answer[equals[i]] : values[i][j];
                    sum += actual[j];
                }
                sheetOrder.Sort(delegate(int a, int b)
                {
                    int rowA = column.Segments[a].SheetRow >= 0 ? column.Segments[a].SheetRow : a;
                    int rowB = column.Segments[b].SheetRow >= 0 ? column.Segments[b].SheetRow : b;
                    // Validation rejects collisions between explicit rows and
                    // legacy list positions before any layout or UI projection.
                    return rowA.CompareTo(rowB);
                });
                // Divergent signs share the same base. Explicit sheet order
                // controls visual order within each sign group.
                for (int group = 0; group < 2; group++)
                {
                    bool positive = group == 0;
                    bool reverse = positive == (spec.SegmentOrder == "SheetTopToBottom");
                    Rational cursor = start;
                    for (int index = 0; index < actual.Length; index++)
                    {
                        int j = sheetOrder[reverse ? actual.Length - 1 - index : index];
                        Rational value = actual[j];
                        if ((value.Sign >= 0) != positive) continue;
                        Rational end = cursor + value;
                        segments[j] = new WaterfallSegmentSolution
                        {
                            Id = column.Segments[j].Id, Calculated = column.Segments[j].Calculated,
                            SheetRow = column.Segments[j].SheetRow,
                            Value = value.ToDouble(), Start = cursor.ToDouble(), End = end.ToDouble(),
                            Bottom = (value.Sign >= 0 ? cursor : end).ToDouble(),
                            Top = (value.Sign >= 0 ? end : cursor).ToDouble(),
                            Extent = (value.Sign >= 0 ? value : -value).ToDouble()
                        };
                        cursor = end;
                    }
                }
                return new WaterfallColumnSolution
                {
                    ColumnId = column.Id, Label = column.Label, Base = start.ToDouble(), Output = sum.ToDouble(),
                    Segments = new List<WaterfallSegmentSolution>(segments)
                };
            }

            private static Rational[] Zeros(int count)
            {
                Rational[] result = new Rational[count];
                for (int i = 0; i < count; i++) result[i] = Rational.Zero;
                return result;
            }
        }

        // Rational arithmetic prevents small residual tolerances from hiding
        // contradictory connectors. Inputs are interpreted as their invariant
        // round-trip decimal representation, rather than raw IEEE binary bits.
        private struct Rational
        {
            private readonly BigInteger numerator, denominator;
            internal static readonly Rational Zero = new Rational(BigInteger.Zero, BigInteger.One);
            internal static readonly Rational One = new Rational(BigInteger.One, BigInteger.One);
            internal bool IsZero { get { return numerator.IsZero; } }
            internal int Sign { get { return numerator.Sign; } }

            private Rational(BigInteger n, BigInteger d)
            {
                if (d.IsZero) throw new DivideByZeroException();
                if (d.Sign < 0) { n = -n; d = -d; }
                BigInteger gcd = BigInteger.GreatestCommonDivisor(BigInteger.Abs(n), d);
                numerator = n / gcd; denominator = d / gcd;
            }

            internal static Rational FromDouble(double value)
            {
                if (Double.IsNaN(value) || Double.IsInfinity(value)) throw new ArgumentException("Waterfall values, weights and levels must be finite.");
                string text = value.ToString("R", CultureInfo.InvariantCulture);
                int exponent = 0, at = text.IndexOf('E');
                if (at < 0) at = text.IndexOf('e');
                if (at >= 0)
                {
                    exponent = Int32.Parse(text.Substring(at + 1), CultureInfo.InvariantCulture);
                    text = text.Substring(0, at);
                }
                at = text.IndexOf('.');
                if (at >= 0) { exponent -= text.Length - at - 1; text = text.Remove(at, 1); }
                BigInteger n = BigInteger.Parse(text, CultureInfo.InvariantCulture);
                return exponent >= 0 ? new Rational(n * BigInteger.Pow(10, exponent), BigInteger.One)
                    : new Rational(n, BigInteger.Pow(10, -exponent));
            }

            internal double ToDouble()
            {
                if (IsZero) return 0;
                double n = (double)numerator, d = (double)denominator;
                double result;
                if (!Double.IsInfinity(n) && !Double.IsInfinity(d)) result = n / d;
                else
                {
                    string ns = BigInteger.Abs(numerator).ToString(CultureInfo.InvariantCulture);
                    string ds = denominator.ToString(CultureInfo.InvariantCulture);
                    int nl = Math.Min(17, ns.Length), dl = Math.Min(17, ds.Length);
                    double nm = Double.Parse(ns.Substring(0, nl), CultureInfo.InvariantCulture) / Math.Pow(10, nl - 1);
                    double dm = Double.Parse(ds.Substring(0, dl), CultureInfo.InvariantCulture) / Math.Pow(10, dl - 1);
                    int exponent = ns.Length - ds.Length;
                    // Parsing scientific notation handles subnormals without
                    // prematurely underflowing Math.Pow(10, exponent).
                    try
                    {
                        result = Double.Parse((nm / dm).ToString("R", CultureInfo.InvariantCulture)
                            + "E" + exponent.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                    }
                    catch (OverflowException ex)
                    {
                        throw new ArgumentException("Solved waterfall coordinates exceed the supported double range.", ex);
                    }
                    if (numerator.Sign < 0) result = -result;
                }
                if (Double.IsNaN(result) || Double.IsInfinity(result) || result == 0)
                    throw new ArgumentException("Solved waterfall coordinates exceed the supported nonzero double range.");
                return result;
            }

            public static Rational operator +(Rational a, Rational b)
            {
                return new Rational(a.numerator * b.denominator + b.numerator * a.denominator, a.denominator * b.denominator);
            }
            public static Rational operator -(Rational a, Rational b) { return a + (-b); }
            public static Rational operator -(Rational a) { return new Rational(-a.numerator, a.denominator); }
            public static Rational operator *(Rational a, Rational b) { return new Rational(a.numerator * b.numerator, a.denominator * b.denominator); }
            public static Rational operator /(Rational a, Rational b) { return new Rational(a.numerator * b.denominator, a.denominator * b.numerator); }
        }
    }
}
