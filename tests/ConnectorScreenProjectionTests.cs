using System;
using System.Drawing;
using OpenSlideCharts;

internal static class ConnectorScreenProjectionTests
{
    private static int checks;
    private static void Check(bool ok) { checks++; if (!ok) throw new Exception("Projection check " + checks); }
    public static int Main()
    {
        foreach (double scale in new double[] { .5, 1, 2, 3 })
        {
            var p = new ConnectorScreenProjection(new Point(-500, 200), new Point(-500 + (int)(900 * scale), 200 + (int)(470 * scale)), 900, 470);
            foreach (PointF point in new PointF[] { new PointF(0, 0), new PointF(900, 470), new PointF(333.5f, 44.5f), new PointF(-10, 480) })
            {
                PointF back = p.ToChart(p.ToScreen(point.X, point.Y));
                Check(Math.Abs(back.X - point.X) <= .501 / scale && Math.Abs(back.Y - point.Y) <= .501 / scale);
            }
            Check(p.Bounds.Left == -516 && p.Bounds.Top == 184);
            Check(p.SameGeometry(p));
            Check(!p.SameGeometry(new ConnectorScreenProjection(new Point(-499, 200), new Point(-499 + (int)(900 * scale), 200 + (int)(470 * scale)), 900, 470)));
        }
        foreach (double bad in new double[] { 0, -1, Double.NaN, Double.PositiveInfinity })
        {
            bool rejected = false;
            try { new ConnectorScreenProjection(Point.Empty, new Point(1, 1), bad, 1); } catch (ArgumentException) { rejected = true; }
            Check(rejected);
        }
        Console.WriteLine("Connector screen projection: " + checks + " checks passed."); return 0;
    }
}
