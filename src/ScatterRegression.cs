using System;
using System.Collections.Generic;
using System.Numerics;

namespace OpenSlideCharts
{
    public sealed class ScatterFit
    {
        // Means and slope use x/XUnit and y/YUnit, not raw coordinates.
        public double XUnit, YUnit, MeanX, MeanY, Slope, RSquared;
        public int Count;
        internal double RawSlope, RawIntercept;
        internal Func<double[], double[]> ExactClip;

        // Axis transforms are raw = Origin + Unit * (Low + fraction * (High-Low)).
        // Returned fractions are not reversed; the renderer applies axis reversal.
        internal double[] ClipToAxes(double xOrigin, double xUnit, double xLow, double xHigh,
            double yOrigin, double yUnit, double yLow, double yHigh)
        {
            double[] axes = { xOrigin, xUnit, xLow, xHigh, yOrigin, yUnit, yLow, yHigh };
            for (int i = 0; i < axes.Length; i++)
                if (Double.IsNaN(axes[i]) || Double.IsInfinity(axes[i])) throw new ArgumentException("Axis transforms must be finite.");
            if (xUnit <= 0 || yUnit <= 0 || xLow >= xHigh || yLow >= yHigh)
                throw new ArgumentException("Axis units must be positive and axis lower bounds less than upper bounds.");
            if (ExactClip == null) throw new InvalidOperationException("Exact clipping requires a fitted result from ScatterRegression.Fit.");
            return ExactClip(axes);
        }

        /// <summary>
        /// Predicts normalized y for finite normalized x. Extrapolations outside
        /// the double range may return infinity. NaN/infinite requests throw.
        /// Public fit fields are intended as an unmodified result snapshot.
        /// </summary>
        public double PredictNormalized(double xNormalized)
        {
            if (Double.IsNaN(xNormalized) || Double.IsInfinity(xNormalized))
                throw new ArgumentOutOfRangeException("xNormalized", "A finite normalized x coordinate is required.");
            return MeanY + Slope * (xNormalized - MeanX);
        }
    }

    /// <summary>
    /// Original least-squares policy. Weighted fits use positive raw Size as
    /// weight; no claim that this reproduces a vendor's bubble-size convention.
    /// Null/missing/nonfinite coordinates are ignored. Weighted fits also ignore
    /// missing/nonfinite/nonpositive Size; ordinary fits never inspect Size.
    /// group=null selects all points; other strings match ordinally, with a
    /// null point.Group treated as the empty group. Inputs are never changed.
    /// </summary>
    public static class ScatterRegression
    {
        /// <summary>
        /// Returns null for fewer than two eligible points, constant x, or a
        /// normalized slope that cannot be represented by a finite double.
        /// Empty/null input returns null. Data omissions do not throw.
        /// Exact rational accumulations preserve the binary double inputs even
        /// when normalizing a coordinate or weight would underflow in double.
        /// Final exposed coefficients are rounded to nearest, ties to even.
        /// </summary>
        public static ScatterFit Fit(List<ScatterPoint> points, bool weighted, string group = null)
        {
            if (points == null) return null;
            List<ScatterPoint> eligible = new List<ScatterPoint>();
            double xUnit = 0, yUnit = 0, weightUnit = 0;
            foreach (ScatterPoint point in points)
            {
                if (point == null || !Finite(point.X) || !Finite(point.Y)) continue;
                if (group != null && !String.Equals(point.Group ?? "", group, StringComparison.Ordinal)) continue;
                if (weighted && (!Finite(point.Size) || point.Size.Value <= 0)) continue;
                eligible.Add(point);
                xUnit = Math.Max(xUnit, Math.Abs(point.X.Value));
                yUnit = Math.Max(yUnit, Math.Abs(point.Y.Value));
                if (weighted) weightUnit = Math.Max(weightUnit, point.Size.Value);
            }
            if (eligible.Count < 2) return null;
            if (xUnit == 0) xUnit = 1;
            if (yUnit == 0) yUnit = 1;
            if (!weighted) weightUnit = 1;
            Rational ux = Rational.FromDouble(xUnit), uy = Rational.FromDouble(yUnit), uw = Rational.FromDouble(weightUnit);
            Rational sumW = Rational.Zero, sumX = Rational.Zero, sumY = Rational.Zero;
            Rational sumXX = Rational.Zero, sumXY = Rational.Zero, sumYY = Rational.Zero;
            foreach (ScatterPoint point in eligible)
            {
                Rational x = Rational.FromDouble(point.X.Value) / ux;
                Rational y = Rational.FromDouble(point.Y.Value) / uy;
                Rational w = weighted ? Rational.FromDouble(point.Size.Value) / uw : Rational.One;
                sumW += w; sumX += w * x; sumY += w * y;
                sumXX += w * x * x; sumXY += w * x * y; sumYY += w * y * y;
            }
            // These centered moments would suffer catastrophic cancellation in
            // floating arithmetic. Here each multiplication/subtraction is exact.
            Rational xx = sumW * sumXX - sumX * sumX;
            if (xx.Sign == 0) return null;
            Rational xy = sumW * sumXY - sumX * sumY;
            Rational yy = sumW * sumYY - sumY * sumY;
            Rational exactSlope = xy / xx, exactMeanX = sumX / sumW, exactMeanY = sumY / sumW;
            double slope = exactSlope.ToDouble();
            if (Double.IsNaN(slope) || Double.IsInfinity(slope)) return null;
            double squared = yy.Sign == 0 ? 1 : (xy * xy / (xx * yy)).ToDouble();
            Rational rawMeanX = exactMeanX * ux, rawMeanY = exactMeanY * uy;
            Rational rawSlope = exactSlope * uy / ux;
            return new ScatterFit {
                XUnit = xUnit, YUnit = yUnit, MeanX = exactMeanX.ToDouble(), MeanY = exactMeanY.ToDouble(),
                Slope = slope, RSquared = Math.Max(0, Math.Min(1, squared)), Count = eligible.Count,
                RawSlope = rawSlope.ToDouble(), RawIntercept = (rawMeanY - rawSlope * rawMeanX).ToDouble(),
                ExactClip = delegate(double[] axes) { return ClipExact(rawMeanX, rawMeanY, rawSlope, axes); }
            };
        }

        private static double[] ClipExact(Rational meanX, Rational meanY, Rational slope, double[] axes)
        {
            Rational ox = Rational.FromDouble(axes[0]), ux = Rational.FromDouble(axes[1]);
            Rational lx = Rational.FromDouble(axes[2]), hx = Rational.FromDouble(axes[3]);
            Rational oy = Rational.FromDouble(axes[4]), uy = Rational.FromDouble(axes[5]);
            Rational ly = Rational.FromDouble(axes[6]), hy = Rational.FromDouble(axes[7]);
            // Virtual endpoints can lie outside double's range, or between two
            // adjacent doubles. Never materialize raw endpoints in floating point.
            Rational xBase = ox + ux * lx, xSpan = ux * (hx - lx);
            Rational yBase = oy + uy * ly, ySpan = uy * (hy - ly);
            Rational intercept = (meanY + slope * (xBase - meanX) - yBase) / ySpan;
            Rational gradient = slope * xSpan / ySpan;
            List<Rational[]> hits = new List<Rational[]>();
            AddHit(hits, Rational.Zero, intercept);
            AddHit(hits, Rational.One, intercept + gradient);
            if (gradient.Sign != 0)
            {
                AddHit(hits, (Rational.Zero - intercept) / gradient, Rational.Zero);
                AddHit(hits, (Rational.One - intercept) / gradient, Rational.One);
            }
            if (hits.Count < 2) return null;
            return new double[] { hits[0][0].ToDouble(), hits[0][1].ToDouble(), hits[1][0].ToDouble(), hits[1][1].ToDouble() };
        }

        private static void AddHit(List<Rational[]> hits, Rational x, Rational y)
        {
            if (x.Sign < 0 || (x - Rational.One).Sign > 0 || y.Sign < 0 || (y - Rational.One).Sign > 0) return;
            foreach (Rational[] previous in hits)
                if ((previous[0] - x).Sign == 0 && (previous[1] - y).Sign == 0) return;
            hits.Add(new Rational[] { x, y });
        }

        private static bool Finite(double? value)
        { return value.HasValue && !Double.IsNaN(value.Value) && !Double.IsInfinity(value.Value); }

        // Fractions are private, immutable and reduced. They are not a chart
        // serialization format. Exact binary extraction avoids decimal parsing.
        private struct Rational
        {
            private readonly BigInteger numerator, denominator;
            internal static readonly Rational Zero = new Rational(BigInteger.Zero, BigInteger.One);
            internal static readonly Rational One = new Rational(BigInteger.One, BigInteger.One);
            internal int Sign { get { return numerator.Sign; } }
            private Rational(BigInteger n, BigInteger d)
            {
                if (d.Sign == 0) throw new DivideByZeroException();
                if (n.Sign == 0) { numerator = BigInteger.Zero; denominator = BigInteger.One; return; }
                if (d.Sign < 0) { n = -n; d = -d; }
                BigInteger gcd = BigInteger.GreatestCommonDivisor(BigInteger.Abs(n), d);
                numerator = n / gcd; denominator = d / gcd;
            }
            internal static Rational FromDouble(double value)
            {
                if (value == 0) return Zero;
                long bits = BitConverter.DoubleToInt64Bits(value);
                int exponent = (int)((bits >> 52) & 0x7FF);
                long mantissa = bits & 0xFFFFFFFFFFFFFL;
                if (exponent == 0) exponent = -1074;
                else { mantissa |= 1L << 52; exponent -= 1075; }
                BigInteger n = new BigInteger(mantissa);
                if (bits < 0) n = -n;
                return exponent >= 0 ? new Rational(n << exponent, BigInteger.One) : new Rational(n, BigInteger.One << -exponent);
            }
            public static Rational operator +(Rational a, Rational b)
            { return new Rational(a.numerator * b.denominator + b.numerator * a.denominator, a.denominator * b.denominator); }
            public static Rational operator -(Rational a, Rational b)
            { return new Rational(a.numerator * b.denominator - b.numerator * a.denominator, a.denominator * b.denominator); }
            public static Rational operator *(Rational a, Rational b)
            { return new Rational(a.numerator * b.numerator, a.denominator * b.denominator); }
            public static Rational operator /(Rational a, Rational b)
            { return new Rational(a.numerator * b.denominator, a.denominator * b.numerator); }

            private static int BitLength(BigInteger value)
            {
                byte[] bytes = value.ToByteArray(); int last = bytes.Length - 1;
                while (last > 0 && bytes[last] == 0) last--;
                int length = last * 8; byte top = bytes[last];
                while (top != 0) { length++; top >>= 1; }
                return length;
            }
            internal double ToDouble()
            {
                if (numerator.Sign == 0) return 0;
                BigInteger n = BigInteger.Abs(numerator), d = denominator;
                int exponent = BitLength(n) - BitLength(d);
                if (exponent >= 0 ? n < (d << exponent) : (n << -exponent) < d) exponent--;
                if (exponent > 1023) return numerator.Sign < 0 ? Double.NegativeInfinity : Double.PositiveInfinity;
                if (exponent < -1075) return numerator.Sign < 0 ? -0.0 : 0.0;
                // Round once at the binary64 mantissa or subnormal boundary.
                int shift = exponent >= -1022 ? 52 - exponent : 1074;
                if (shift >= 0) n <<= shift; else d <<= -shift;
                BigInteger remainder;
                BigInteger rounded = BigInteger.DivRem(n, d, out remainder);
                int midpoint = (remainder << 1).CompareTo(d);
                if (midpoint > 0 || midpoint == 0 && !rounded.IsEven) rounded++;
                double result = (double)rounded * (exponent >= -1022 ? Math.Pow(2, exponent - 52) : Double.Epsilon);
                return numerator.Sign < 0 ? -result : result;
            }
        }
    }
}
