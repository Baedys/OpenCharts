using System;
using System.Globalization;

namespace OpenSlideCharts
{
    public sealed class OfficeChartSelection
    {
        public object Chart, Shape;
        public ChartSpec Spec;
        public Primitive Primitive;

        public static OfficeChartSelection Resolve(object selectionObject)
        {
            if (selectionObject == null) return null;
            dynamic selection = selectionObject;
            if ((int)selection.Type != 2) return null;
            dynamic range = (bool)selection.HasChildShapeRange ? selection.ChildShapeRange : selection.ShapeRange;
            if ((int)range.Count != 1) return null;
            dynamic selected = range.Item(1), root = selected;
            for (int depth = 0; depth < 16; depth++)
            {
                if ((string)root.Tags.Item("OSC_VERSION") == "1")
                {
                    // Replacing nested chart groups would change their parent/z-order.
                    if ((int)root.Child != 0) return null;
                    ChartSpec spec = ChartStorage.Read((object)root);
                    var result = new OfficeChartSelection { Chart = root, Shape = selected, Spec = spec };
                    int index;
                    if (Int32.TryParse((string)selected.Tags.Item("OSC_ITEM"), NumberStyles.None, CultureInfo.InvariantCulture, out index))
                    {
                        Layout layout = ChartEngine.Build(spec, (double)root.Width, (double)root.Height);
                        if (index >= 0 && index < layout.Items.Count && String.Equals(layout.Items[index].Role, (string)selected.Tags.Item("OSC_ROLE"), StringComparison.OrdinalIgnoreCase)) result.Primitive = layout.Items[index];
                    }
                    return result;
                }
                if ((int)root.Child == 0) return null;
                root = root.ParentGroup;
            }
            return null;
        }
    }

    public sealed class ConnectorEndpoints
    {
        public double X1, Y1, X2, Y2;
        public double Difference(ConnectorEndpoints other)
        { return Math.Max(Math.Max(Math.Abs(X1 - other.X1), Math.Abs(Y1 - other.Y1)), Math.Max(Math.Abs(X2 - other.X2), Math.Abs(Y2 - other.Y2))); }
        public static ConnectorEndpoints Read(object shapeObject)
        {
            dynamic shape = shapeObject;
            if (Math.Abs((double)shape.Rotation) > .01) throw new NotSupportedException("Déplacez une extrémité du connecteur sans utiliser la poignée de rotation.");
            double x = shape.Left, y = shape.Top, w = shape.Width, h = shape.Height;
            bool flipX = (int)shape.HorizontalFlip == -1, flipY = (int)shape.VerticalFlip == -1;
            return new ConnectorEndpoints { X1 = x + (flipX ? w : 0), X2 = x + (flipX ? 0 : w), Y1 = y + (flipY ? h : 0), Y2 = y + (flipY ? 0 : h) };
        }
    }
}
