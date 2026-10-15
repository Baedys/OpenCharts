using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace OpenSlideCharts
{
    // Affine mapping for an unrotated chart. No Office objects or document mutations.
    [ComVisible(false)]
    public sealed class ConnectorScreenProjection
    {
        private readonly Point origin, corner;
        private readonly double width, height;
        public ConnectorScreenProjection(Point origin, Point corner, double width, double height)
        {
            if (!Finite(width) || !Finite(height) || width <= 0 || height <= 0 || corner.X <= origin.X || corner.Y <= origin.Y)
                throw new ArgumentException("Projection du graphique invalide.");
            this.origin = origin; this.corner = corner; this.width = width; this.height = height;
        }
        private static bool Finite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }
        public Point ToScreen(double x, double y)
        {
            return new Point(checked((int)Math.Round(origin.X + x / width * ((double)corner.X - origin.X))),
                checked((int)Math.Round(origin.Y + y / height * ((double)corner.Y - origin.Y))));
        }
        public PointF ToChart(Point point)
        {
            return new PointF((float)(((double)point.X - origin.X) * width / ((double)corner.X - origin.X)),
                (float)(((double)point.Y - origin.Y) * height / ((double)corner.Y - origin.Y)));
        }
        public Rectangle Bounds
        {
            get { return Rectangle.FromLTRB(checked(origin.X - 16), checked(origin.Y - 16), checked(corner.X + 16), checked(corner.Y + 16)); }
        }
        public bool SameGeometry(ConnectorScreenProjection other)
        { return other != null && origin == other.origin && corner == other.corner && width == other.width && height == other.height; }
    }
}
