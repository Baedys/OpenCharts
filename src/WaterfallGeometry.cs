using System;
using System.Collections.Generic;

namespace OpenSlideCharts
{
    public sealed class WaterfallSceneColumn
    {
        public string ColumnId;
        public double X, Width, BaseY, OutputY;
    }

    public sealed class WaterfallSceneConnector
    {
        public string Id;
        public double X1, Y1, X2, Y2;
    }

    public sealed class WaterfallTarget
    {
        public double X, Y;
        public WaterfallAnchor Anchor;
        public string ColumnId, SegmentId, Edge, Side;
    }

    public sealed class WaterfallHit
    {
        // Equivalent anchor aliases are collapsed; distinct equal-distance
        // anchors remain candidates for an explicit UI choice.
        public List<WaterfallTarget> Candidates = new List<WaterfallTarget>();
        public double Distance = Double.PositiveInfinity;
        public bool IsAmbiguous { get { return Candidates.Count > 1; } }
        public WaterfallTarget Target { get { return Candidates.Count == 1 ? Candidates[0] : null; } }
    }

    public sealed class WaterfallScene
    {
        public WaterfallSolution Solution;
        public double Minimum, Maximum, PlotLeft, PlotTop, PlotWidth, PlotHeight, Step, BarWidth;
        public List<WaterfallSceneColumn> Columns = new List<WaterfallSceneColumn>();
        public List<WaterfallSceneConnector> Connectors = new List<WaterfallSceneConnector>();
        public List<WaterfallTarget> Targets = new List<WaterfallTarget>();
        private double unit, scaledMinimum, scaledMaximum;

        public static WaterfallScene Build(WaterfallSpec spec, double width, double height)
        {
            if (!Finite(width) || !Finite(height) || width < 320 || height < 240 || width > 100000 || height > 100000)
                throw new ArgumentException("Waterfall scene dimensions must be finite, from 320 x 240 to 100000 x 100000 points.");
            WaterfallScene scene = new WaterfallScene { Solution = WaterfallSolver.Solve(spec) };
            foreach (WaterfallColumnSolution column in scene.Solution.Columns)
            {
                scene.Include(column.Base); scene.Include(column.Output);
                foreach (WaterfallSegmentSolution segment in column.Segments) { scene.Include(segment.Bottom); scene.Include(segment.Top); }
            }
            foreach (WaterfallConnectorSolution edge in scene.Solution.Connectors) { scene.Include(edge.LeftLevel); scene.Include(edge.RightLevel); }
            scene.unit = Math.Max(Math.Abs(scene.Minimum), Math.Abs(scene.Maximum));
            if (scene.unit == 0) scene.unit = 1;
            scene.scaledMinimum = scene.Minimum / scene.unit; scene.scaledMaximum = scene.Maximum / scene.unit;
            if (scene.scaledMinimum == scene.scaledMaximum) { scene.scaledMinimum = 0; scene.scaledMaximum = 1; }
            scene.PlotLeft = 58; scene.PlotTop = 64; scene.PlotWidth = width - 82; scene.PlotHeight = height - 142;
            scene.Step = scene.PlotWidth / scene.Solution.Columns.Count; scene.BarWidth = .62 * scene.Step;
            Dictionary<string, int> indices = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < scene.Solution.Columns.Count; i++)
            {
                WaterfallColumnSolution column = scene.Solution.Columns[i];
                indices.Add(column.ColumnId, i);
                WaterfallSceneColumn geometry = new WaterfallSceneColumn
                {
                    ColumnId = column.ColumnId, X = scene.PlotLeft + scene.Step * (i + .5) - scene.BarWidth / 2,
                    Width = scene.BarWidth, BaseY = scene.MapY(column.Base), OutputY = scene.MapY(column.Output)
                };
                scene.Columns.Add(geometry);
                scene.AddTargets(spec.Columns[i], column, geometry, spec.SegmentOrder);
            }
            foreach (WaterfallConnectorSolution connector in scene.Solution.Connectors)
            {
                int a = indices[connector.LeftColumnId], b = indices[connector.RightColumnId];
                double direction = a < b ? 1 : -1;
                scene.Connectors.Add(new WaterfallSceneConnector
                {
                    Id = connector.Id,
                    X1 = scene.PlotLeft + scene.Step * (a + .5) + direction * scene.BarWidth / 2,
                    X2 = scene.PlotLeft + scene.Step * (b + .5) - direction * scene.BarWidth / 2,
                    Y1 = scene.MapY(connector.LeftLevel), Y2 = scene.MapY(connector.RightLevel)
                });
            }
            return scene;
        }

        public double MapY(double value)
        {
            if (!Finite(value)) throw new ArgumentException("Scene coordinates must be finite.");
            double result = PlotTop + PlotHeight * (1 - (value / unit - scaledMinimum) / (scaledMaximum - scaledMinimum));
            if (!Finite(result)) throw new ArgumentException("Value cannot be represented in this scene's scale.");
            return result;
        }

        public WaterfallHit FindNearest(double x, double y, double tolerance)
        {
            if (!Finite(x) || !Finite(y) || !Finite(tolerance) || tolerance < 0)
                throw new ArgumentException("Hit coordinates and nonnegative tolerance must be finite.");
            WaterfallHit hit = new WaterfallHit();
            foreach (WaterfallTarget target in Targets)
            {
                double distance = Distance(x - target.X, y - target.Y);
                if (distance > tolerance) continue;
                if (distance < hit.Distance - 1e-7)
                {
                    hit.Candidates.Clear(); hit.Distance = distance;
                }
                if (Math.Abs(distance - hit.Distance) > 1e-7) continue;
                bool alias = false;
                foreach (WaterfallTarget existing in hit.Candidates)
                    if (WaterfallInteraction.SameAnchor(existing.Anchor, target.Anchor)) { alias = true; break; }
                if (!alias) hit.Candidates.Add(target);
            }
            return hit;
        }

        private void Include(double value) { Minimum = Math.Min(Minimum, value); Maximum = Math.Max(Maximum, value); }
        private static bool Finite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }
        private static double Distance(double x, double y)
        {
            x = Math.Abs(x); y = Math.Abs(y); double maximum = Math.Max(x, y);
            if (maximum == 0) return 0;
            if (Double.IsInfinity(maximum)) return Double.PositiveInfinity;
            return maximum * Math.Sqrt((x / maximum) * (x / maximum) + (y / maximum) * (y / maximum));
        }

        private void AddTargets(WaterfallColumn source, WaterfallColumnSolution solved, WaterfallSceneColumn geometry, string order)
        {
            AddPair(geometry, WaterfallAnchor.AtBase(source), null, "Base", solved.Base);
            AddPair(geometry, WaterfallAnchor.AtOutput(source), null, "Output", solved.Output);
            List<int> indices = new List<int>();
            for (int i = 0; i < source.Segments.Count; i++) indices.Add(i);
            indices.Sort(delegate(int a, int b)
            {
                int rowA = source.Segments[a].SheetRow >= 0 ? source.Segments[a].SheetRow : a;
                int rowB = source.Segments[b].SheetRow >= 0 ? source.Segments[b].SheetRow : b;
                return rowA.CompareTo(rowB);
            });
            for (int group = 0; group < 2; group++)
            {
                bool positive = group == 0;
                bool reverse = positive == (order == "SheetTopToBottom");
                double[] cursor = new double[source.Segments.Count];
                for (int at = 0; at < indices.Count; at++)
                {
                    int i = indices[reverse ? indices.Count - 1 - at : at];
                    WaterfallSegmentSolution segment = solved.Segments[i];
                    if ((segment.Value >= 0) != positive) continue;
                    double[] end = (double[])cursor.Clone(); end[i] = 1;
                    AddPair(geometry, WaterfallAnchor.ForWeights(source.Id, positive ? cursor : end), segment.Id, "Bottom", segment.Bottom);
                    AddPair(geometry, WaterfallAnchor.ForWeights(source.Id, positive ? end : cursor), segment.Id, "Top", segment.Top);
                    cursor = end;
                }
            }
        }

        private void AddPair(WaterfallSceneColumn column, WaterfallAnchor anchor, string segmentId, string edge, double level)
        {
            for (int side = 0; side < 2; side++)
                Targets.Add(new WaterfallTarget
                {
                    X = column.X + (side == 0 ? 0 : column.Width), Y = MapY(level),
                    Anchor = WaterfallInteraction.CopyAnchor(anchor), ColumnId = column.ColumnId,
                    SegmentId = segmentId, Edge = edge, Side = side == 0 ? "Left" : "Right"
                });
        }
    }

    public sealed class WaterfallEditResult
    {
        public WaterfallSolution Solution;
        public List<string> AddedPinIds = new List<string>();
        public List<string> RemovedPinIds = new List<string>();
    }

    public static class WaterfallInteraction
    {
        public static WaterfallEditResult Reattach(WaterfallSpec model, string connectorId, bool left, WaterfallAnchor targetAnchor)
        {
            WaterfallSolver.Solve(model);
            WaterfallConnector connector = FindConnector(model, connectorId);
            WaterfallAnchor previous = left ? connector.Left : connector.Right;
            WaterfallAnchor next = CopyAnchor(targetAnchor);
            List<WaterfallPin> originalPins = model.Pins;
            model.Pins = new List<WaterfallPin>(originalPins);
            if (left) connector.Left = next; else connector.Right = next;
            try
            {
                WaterfallEditResult result = new WaterfallEditResult();
                if (connector.Constrain && !SameAnchor(previous, next)) RepairDetachedBase(model, previous, result);
                result.Solution = WaterfallSolver.Solve(model);
                return result;
            }
            catch
            {
                if (left) connector.Left = previous; else connector.Right = previous;
                model.Pins = originalPins;
                throw;
            }
        }

        public static WaterfallEditResult DeleteConnector(WaterfallSpec model, string connectorId)
        {
            WaterfallSolver.Solve(model);
            WaterfallConnector connector = FindConnector(model, connectorId);
            List<WaterfallConnector> originalConnectors = model.Connectors;
            List<WaterfallPin> originalPins = model.Pins;
            model.Connectors = new List<WaterfallConnector>(originalConnectors);
            model.Pins = new List<WaterfallPin>(originalPins);
            model.Connectors.Remove(connector);
            try
            {
                WaterfallEditResult result = new WaterfallEditResult();
                if (connector.Constrain)
                {
                    RepairDetachedBase(model, connector.Left, result);
                    RepairDetachedBase(model, connector.Right, result);
                }
                result.Solution = WaterfallSolver.Solve(model);
                return result;
            }
            catch
            {
                model.Connectors = originalConnectors; model.Pins = originalPins;
                throw;
            }
        }

        // Bounded original repair rule, not a claim about think-cell internals:
        // detach a numeric base -> start a new sum there; release only zero-level
        // generated output pins reachable along numeric output-to-base links.
        // An e base, a manual pin, and every other topology are never rewritten.
        // Remaining conflicts/degrees of freedom are reported by the solver.
        private static void RepairDetachedBase(WaterfallSpec model, WaterfallAnchor previous, WaterfallEditResult result)
        {
            WaterfallColumn root = FindColumn(model, previous.ColumnId);
            if (HasEquals(root) || !SameAnchor(previous, WaterfallAnchor.AtBase(root))) return;
            foreach (WaterfallConnector edge in model.Connectors)
                if (edge.Constrain && (SameAnchor(edge.Left, previous) || SameAnchor(edge.Right, previous))) return;
            HashSet<string> downstream = new HashSet<string>(StringComparer.Ordinal);
            Queue<WaterfallColumn> pending = new Queue<WaterfallColumn>();
            downstream.Add(root.Id); pending.Enqueue(root);
            while (pending.Count > 0)
            {
                WaterfallColumn column = pending.Dequeue();
                WaterfallAnchor output = WaterfallAnchor.AtOutput(column);
                foreach (WaterfallConnector edge in model.Connectors)
                {
                    if (!edge.Constrain) continue;
                    WaterfallAnchor other = SameAnchor(edge.Left, output) ? edge.Right : (SameAnchor(edge.Right, output) ? edge.Left : null);
                    if (other == null) continue;
                    WaterfallColumn next = FindColumn(model, other.ColumnId);
                    if (HasEquals(next) || !SameAnchor(other, WaterfallAnchor.AtBase(next))) continue;
                    if (downstream.Add(next.Id)) pending.Enqueue(next);
                }
            }
            for (int i = model.Pins.Count - 1; i >= 0; i--)
            {
                WaterfallPin pin = model.Pins[i];
                if (!pin.AutoGenerated || pin.Level != 0 || !downstream.Contains(pin.Anchor.ColumnId)) continue;
                WaterfallColumn column = FindColumn(model, pin.Anchor.ColumnId);
                if (!SameAnchor(pin.Anchor, WaterfallAnchor.AtOutput(column))) continue;
                result.RemovedPinIds.Insert(0, pin.Id); model.Pins.RemoveAt(i);
            }
            foreach (WaterfallPin pin in model.Pins) if (SameAnchor(pin.Anchor, previous)) return;
            string prefix = "auto_base_" + model.Columns.IndexOf(root);
            string id = prefix;
            int suffix = 0;
            while (model.Pins.Exists(delegate(WaterfallPin pin) { return pin.Id == id; })) id = prefix + "_" + (++suffix);
            model.Pins.Add(new WaterfallPin { Id = id, Anchor = CopyAnchor(previous), Level = 0, AutoGenerated = true });
            result.AddedPinIds.Add(id);
        }

        private static bool HasEquals(WaterfallColumn column)
        { foreach (WaterfallSegment segment in column.Segments) if (segment.Calculated) return true; return false; }
        private static WaterfallColumn FindColumn(WaterfallSpec model, string id)
        {
            foreach (WaterfallColumn column in model.Columns) if (column.Id == id) return column;
            throw new ArgumentException("Unknown waterfall column: " + id);
        }
        private static WaterfallConnector FindConnector(WaterfallSpec model, string id)
        {
            foreach (WaterfallConnector edge in model.Connectors) if (edge.Id == id) return edge;
            throw new ArgumentException("Unknown waterfall connector: " + id);
        }
        internal static WaterfallAnchor CopyAnchor(WaterfallAnchor anchor)
        {
            if (anchor == null || anchor.Weights == null) throw new ArgumentException("Connector target anchor is required.");
            return new WaterfallAnchor { ColumnId = anchor.ColumnId, Weights = new List<double>(anchor.Weights) };
        }
        public static bool SameAnchor(WaterfallAnchor a, WaterfallAnchor b)
        {
            if (a == null || b == null || a.ColumnId != b.ColumnId || a.Weights == null || b.Weights == null || a.Weights.Count != b.Weights.Count) return false;
            for (int i = 0; i < a.Weights.Count; i++) if (a.Weights[i] != b.Weights[i]) return false;
            return true;
        }
    }
}
