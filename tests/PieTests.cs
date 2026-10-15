using System;
using System.Collections.Generic;
using System.Globalization;
using OpenSlideCharts;

internal static class PieTests
{
    private struct Point { internal double X, Y; internal Point(double x, double y) { X = x; Y = y; } }
    private static int checks;
    private static void Check(bool ok, string message) { if (!ok) throw new Exception(message); checks++; }
    private static void Near(double actual, double expected, double tolerance, string message)
    { Check(Math.Abs(actual - expected) <= tolerance, message + ": " + actual + " vs " + expected); }
    private static ChartSpec Spec(params double[] values)
    {
        ChartSpec spec = new ChartSpec { Kind = "Pie", Pie = new PieOptions { LabelContent = "None", ShowCategory = false } };
        for (int i = 0; i < values.Length; i++) spec.Rows.Add(new ChartRow { Label = "Category " + i, Value = values[i] });
        return spec;
    }
    private static List<Primitive> Role(Layout l, string role)
    { return l.Items.FindAll(delegate(Primitive p) { return p.Role == role; }); }
    private static Layout Build(ChartSpec spec) { return ChartEngine.Build(spec, 320, 240); }
    private static void Reject(ChartSpec spec, string message)
    {
        bool rejected = false;
        try { Build(spec); } catch (ArgumentException) { rejected = true; } catch (NotSupportedException) { rejected = true; }
        Check(rejected, message);
    }
    private static List<Point> Flatten(ShapePath path)
    {
        List<Point> result = new List<Point>(); Point previous = new Point(path.StartX, path.StartY); result.Add(previous);
        foreach (PathSegment s in path.Segments)
        {
            if (s.Curve) for (int i = 1; i <= 96; i++)
            {
                double t = i / 96.0, u = 1 - t;
                result.Add(new Point(u*u*u*previous.X + 3*u*u*t*s.C1X + 3*u*t*t*s.C2X + t*t*t*s.X,
                    u*u*u*previous.Y + 3*u*u*t*s.C1Y + 3*u*t*t*s.C2Y + t*t*t*s.Y));
            }
            else result.Add(new Point(s.X, s.Y));
            previous = new Point(s.X, s.Y);
        }
        return result;
    }
    private static double Area(List<Point> points)
    {
        double twice = 0;
        for (int i = 0; i < points.Count; i++)
        { Point a = points[i], b = points[(i + 1) % points.Count]; twice += a.X * b.Y - b.X * a.Y; }
        return twice / 2;
    }
    private static int Winding(List<Point> polygon, Point p)
    {
        int winding = 0;
        for (int i = 0; i < polygon.Count; i++)
        {
            Point a = polygon[i], b = polygon[(i + 1) % polygon.Count];
            double cross = (b.X-a.X)*(p.Y-a.Y)-(p.X-a.X)*(b.Y-a.Y);
            if (a.Y <= p.Y && b.Y > p.Y && cross > 0) winding++;
            if (a.Y > p.Y && b.Y <= p.Y && cross < 0) winding--;
        }
        return winding;
    }
    private static void Coordinate(double x, double y, double width, double height)
    {
        if (Double.IsNaN(x) || Double.IsInfinity(x) || Double.IsNaN(y) || Double.IsInfinity(y) ||
            x < -1e-7 || y < -1e-7 || x > width + 1e-7 || y > height + 1e-7)
            throw new Exception("Nonfinite or out-of-frame path coordinate.");
    }
    private static void Frame(Layout l, double width, double height)
    {
        foreach (Primitive p in l.Items)
        {
            Coordinate(p.X, p.Y, width, height); Coordinate(p.X+p.W, p.Y+p.H, width, height);
            if (p.Path == null) continue;
            Coordinate(p.Path.StartX, p.Path.StartY, width, height);
            foreach (PathSegment s in p.Path.Segments)
            {
                Coordinate(s.X, s.Y, width, height);
                if (s.Curve) { Coordinate(s.C1X,s.C1Y,width,height); Coordinate(s.C2X,s.C2Y,width,height); }
            }
            PathSegment last = p.Path.Segments[p.Path.Segments.Count-1];
            Check(last.X == p.Path.StartX && last.Y == p.Path.StartY && !last.Curve, "Explicit radial closure");
        }
        checks++;
    }
    private static void Labels(Layout l)
    {
        List<Primitive> labels = Role(l,"PIE_LABEL");
        for (int i = 0; i < labels.Count; i++) for (int j = i+1; j < labels.Count; j++)
        {
            Primitive a=labels[i],b=labels[j];
            bool intersect = a.X < b.X+b.W && b.X < a.X+a.W && a.Y < b.Y+b.H-1e-7 && b.Y < a.Y+a.H-1e-7;
            if (intersect) throw new Exception("Labels overlap: "+a.EntityId+", "+b.EntityId);
        }
            foreach (Primitive p in labels)
            {
                Check(p.FontSize >= 6 && p.NoWrap && p.VerticalAlignment == "center", "Readable single-line label contract");
                Check(p.Text.Length * p.FontSize * 1.2 <= p.W + 1e-7, "Conservative UTF16 width budget");
                Check(CleanUtf16(p.Text), "Sanitized, unbroken UTF16 labels");
            }
        checks++;
    }

        private static bool CleanUtf16(string text)
        {
            for (int i=0;i<text.Length;i++)
            {
                if (Char.IsControl(text[i]) || text[i]=='\u2028' || text[i]=='\u2029') return false;
                if (Char.IsHighSurrogate(text[i]))
                { if (i+1>=text.Length || !Char.IsLowSurrogate(text[++i])) return false; }
                else if (Char.IsLowSurrogate(text[i])) return false;
            }
            return true;
        }

    public static int Main()
    {
        try
        {
            ChartSpec single=Spec(1);
            Primitive disc=Role(Build(single),"PIE_SLICE")[0];
            double radius=79, discArea=Math.PI*radius*radius;
            Near(Area(Flatten(disc.Path)),discArea,discArea*.0005,"Full disc area");
            Check(Winding(Flatten(disc.Path),new Point(160,139)) == 1,"Pie fills center");
            single.Kind="Doughnut";
            Primitive ring=Role(Build(single),"PIE_SLICE")[0];
            Near(Area(Flatten(ring.Path)),discArea*.75,discArea*.0005,"Annulus subtracts inner area");
            Check(Winding(Flatten(ring.Path),new Point(160,139)) == 0,"Full ring hole genuinely unfilled");
            Check(Winding(Flatten(ring.Path),new Point(160+radius*.75,139)) == 1,"Ring band filled");
            int firstLine=ring.Path.Segments.FindIndex(delegate(PathSegment s) { return !s.Curve; });
            Check(firstLine==4 && ring.Path.Segments.Count==10,"Full ring outer arc, inward seam, reversed inner arc, closing seam");
            ShapePath outer=new ShapePath {StartX=ring.Path.StartX,StartY=ring.Path.StartY};
            outer.Segments.AddRange(ring.Path.Segments.GetRange(0,firstLine));
            PathSegment inward=ring.Path.Segments[firstLine];
            ShapePath inner=new ShapePath {StartX=inward.X,StartY=inward.Y};
            inner.Segments.AddRange(ring.Path.Segments.GetRange(firstLine+1,4));
            Check(Area(Flatten(outer))>0 && Area(Flatten(inner))<0,"Outer clockwise, inner reversed");
            Frame(Build(single),320,240);

            ChartSpec divided=Spec(1,2,3,4);
            List<Primitive> slices=Role(Build(divided),"PIE_SLICE"); double area=0;
            for(int i=0;i<slices.Count;i++)
            {
                double actual=Area(Flatten(slices[i].Path)); area+=actual;
                Near(actual,discArea*(i+1)/10,discArea*.0005,"Per-slice proportions");
            }
            Near(area,discArea,discArea*.0005,"Slices cover circle without excess");
            divided.Kind="Doughnut"; divided.Pie.HoleRatio=.8;
            area=0; foreach(Primitive p in Role(Build(divided),"PIE_SLICE")) area+=Area(Flatten(p.Path));
            Near(area,discArea*.36,discArea*.0005,"Partial annuli cover ring");

            ChartSpec order=Spec(4,1,1,2); ChartRow original=order.Rows[0];
            List<Primitive> originalSlices=Role(Build(order),"PIE_SLICE");
            order.Pie.Order="Ascending";
            List<Primitive> sorted=Role(Build(order),"PIE_SLICE");
            Check(sorted[0].EntityId=="1" && sorted[1].EntityId=="2" && sorted[2].EntityId=="3" && sorted[3].EntityId=="0","Stable ascending ties");
            foreach(Primitive p in sorted) Check(p.Color==originalSlices[Int32.Parse(p.EntityId)].Color,"Color follows data index");
            order.Pie.Order="Descending"; sorted=Role(Build(order),"PIE_SLICE");
            Check(sorted[0].EntityId=="0" && sorted[1].EntityId=="3" && sorted[2].EntityId=="1" && sorted[3].EntityId=="2","Stable descending ties");
            Check(Object.ReferenceEquals(original,order.Rows[0]) && order.Rows[0].Value==4 && order.Pie.Slices.Count==0,"Input rows and styles unchanged");
            ChartSpec counter=Spec(1,3); counter.Pie.Order="SheetCounterclockwise";
            Primitive first=Role(Build(counter),"PIE_SLICE")[0];
            Near(first.Path.StartX,160-radius,1e-7,"Counterclockwise first slice starts to left");
            Near(first.Path.StartY,139,1e-7,"Counterclockwise first slice position");
            Check(Area(Flatten(first.Path))>0,"Counterclockwise placement keeps clockwise contour");
            counter.Pie.Order="SheetClockwise"; counter.Pie.StartAngle=0;
            first=Role(Build(counter),"PIE_SLICE")[0];
            Near(first.Path.StartX,160+radius,1e-7,"Rotation starts at right");
            Near(first.Path.StartY,139,1e-7,"Screen angle zero");

            ChartSpec extreme=Spec(Double.MaxValue,Double.MaxValue,0,Double.Epsilon);
            Layout extremeLayout=Build(extreme); Frame(extremeLayout,320,240);
            Check(Role(extremeLayout,"PIE_SLICE").Count==2,"Zero and unrepresentable tiny values emit no degenerate shapes");
            area=0; foreach(Primitive p in Role(extremeLayout,"PIE_SLICE")) area+=Area(Flatten(p.Path));
            Near(area,discArea,discArea*.0005,"Overflow-safe full coverage");
            Frame(Build(Spec(Double.Epsilon,Double.Epsilon)),320,240);
            ChartSpec rotated=Spec(1); rotated.Kind="Doughnut"; rotated.Pie.StartAngle=Double.MaxValue;
            Frame(Build(rotated),320,240);
            rotated.Pie.StartAngle=43; Frame(ChartEngine.Build(rotated,100000,100000),100000,100000);
            rotated.Pie.HoleRatio=0; Check(Winding(Flatten(Role(Build(rotated),"PIE_SLICE")[0].Path),new Point(160,139))==1,"Doughnut zero hole becomes disc");
            rotated.Pie.HoleRatio=.999999; Frame(Build(rotated),320,240);
            ChartSpec exploded=Spec(1,1); exploded.Pie.StartAngle=0;
            exploded.Pie.Slices.Add(new PieSliceStyle {Index=0,Explode=.5,Color="#12ABEF"});
            Primitive explodedSlice=Role(Build(exploded),"PIE_SLICE")[0];
            Check(explodedSlice.Color==(0x12|(0xab<<8)|(0xef<<16)),"Office packed custom RGB");
            PathSegment center=explodedSlice.Path.Segments[explodedSlice.Path.Segments.Count-2];
            Near(center.X,160,1e-7,"Explosion midpoint direction X");
            Near(center.Y,139+(79/1.5)*.5,1e-7,"Explosion fraction of radius"); Frame(Build(exploded),320,240);

            ChartSpec labels=ChartSamples.Create("Pie");
            Layout labelLayout=Build(labels); Labels(labelLayout); Frame(labelLayout,320,240);
            Check(labels.Pie==null,"Defaults never mutate options");
            labels.Pie=new PieOptions {LabelContent="None",ShowCategory=true};
            labelLayout=Build(labels); Check(Role(labelLayout,"PIE_LABEL").Count==3,"None still permits category labels");
            labels.Pie.ShowCategory=false; Check(Role(Build(labels),"PIE_LABEL").Count==0,"No label content and no category");
            labels.Pie.LabelContent="Both"; labels.Pie.PercentDigits=2;
            labelLayout=ChartEngine.Build(labels,640,480); Check(Role(labelLayout,"PIE_LABEL")[0].Text.Contains("%") && Role(labelLayout,"PIE_LABEL")[0].Text.Contains("("),"Both preserves value and percentage");
            labels.Pie.ShowCategory=true; labels.Rows[0].Label=new String('W',200);
            labelLayout=ChartEngine.Build(labels,640,480);
            Primitive longLabel=Role(labelLayout,"PIE_LABEL").Find(delegate(Primitive p) {return p.EntityId=="0";});
            Check(longLabel.Text.Contains("\u2026") && longLabel.Text.EndsWith("40 (40.00%)"),"Only category abbreviated, numeric content intact");
            Labels(labelLayout);
            labels.Kind="Doughnut"; labels.Pie.CenterTitle="Total";
            Check(Role(ChartEngine.Build(labels,640,480),"PIE_CENTER_TITLE").Count==1,"Doughnut center title");
            ChartSpec wide=Spec(1); wide.Pie.LabelContent="None"; wide.Pie.ShowCategory=true;
            wide.Rows[0].Label=new String('W',100);
            Layout wideLayout=ChartEngine.Build(wide,720,400); Labels(wideLayout);
            Primitive wideLabel=Role(wideLayout,"PIE_LABEL")[0];
            Check(wideLabel.Text.Length==13 && wideLabel.Text.EndsWith("\u2026"),"Wide W regression limits 180pt lane to conservative 13 UTF16 units");
            wide.Rows[0].Label=new String('\u754c',100);
            Labels(ChartEngine.Build(wide,720,400));
            wide.Rows[0].Label="A\tB\r\nC\0D\u0085E\u2028F";
            wideLayout=ChartEngine.Build(wide,720,400); Labels(wideLayout);
            Check(Role(wideLayout,"PIE_LABEL")[0].Text.StartsWith("A B  C D E F"),"Tabs, newlines and controls become spaces");
            wide.Rows[0].Label=new String('A',11)+"\ud83d\ude00"+new String('W',20);
            wideLayout=ChartEngine.Build(wide,720,400); Labels(wideLayout);
            Check(Role(wideLayout,"PIE_LABEL")[0].Text==new String('A',11)+"\u2026","Truncation never leaves half surrogate pair");
            wide.Kind="Doughnut"; wide.Pie.CenterTitle="A\tB\nC\0D";
            Primitive cleanTitle=Role(ChartEngine.Build(wide,720,400),"PIE_CENTER_TITLE")[0];
            Check(CleanUtf16(cleanTitle.Text) && cleanTitle.Text.Length*cleanTitle.FontSize*1.2 <= cleanTitle.W+1e-7,"Center title sanitation and width budget");
            ChartSpec dense=Spec(); for(int i=0;i<80;i++) dense.Rows.Add(new ChartRow {Label="R"+i,Value=1});
            dense.Pie=new PieOptions();
            bool frameError=false; try {Build(dense);} catch(ArgumentException e){frameError=e.Message.Contains("frame");}
            Check(frameError,"Dense minimum frame fails explicitly rather than overlapping");
            Layout denseLayout=ChartEngine.Build(dense,960,1000); Labels(denseLayout); Frame(denseLayout,960,1000);
            Check(Role(denseLayout,"PIE_LABEL").Count==80,"80 labels supported with adequate frame");
            dense.Rows.RemoveRange(32,48); Layout medium=ChartEngine.Build(dense,640,480); Labels(medium); Frame(medium,640,480);
            Check(Role(medium,"PIE_LABEL").Count==32,"32 labels distributed without overlap");
            ChartSpec tall=Spec(); tall.Pie=new PieOptions();
            for(int i=0;i<40;i++) tall.Rows.Add(new ChartRow {Label=new String('W',100),Value=i==0?100000:1});
            typeof(ChartEngine).GetMethod("ValidatePie",System.Reflection.BindingFlags.Static|System.Reflection.BindingFlags.NonPublic).Invoke(null,new object[]{tall});
            Check(true,"Data validation does not impose a provisional chart frame");
            bool smallFrame=false;try {ChartEngine.Build(tall,720,400);}catch(ArgumentException e){smallFrame=e.Message.Contains("frame");}
            Check(smallFrame,"Actual rendering still validates dense-label frame");
            Layout tallLayout=ChartEngine.Build(tall,720,1000); Labels(tallLayout); Frame(tallLayout,720,1000);

            Reject(Spec(-1,2),"Negative values rejected"); Reject(Spec(0,0),"All-zero rejected");
            Reject(Spec(Double.NaN),"NaN rejected"); Reject(Spec(Double.PositiveInfinity),"Infinity rejected");
            ChartSpec bad=Spec(1); bad.Pie.StartAngle=Double.NaN; Reject(bad,"Angle finite");
            bad=Spec(1); bad.Pie.HoleRatio=1; Reject(bad,"Hole ratio excludes empty annulus");
            bad=Spec(1); bad.Pie.HoleRatio=-.1; Reject(bad,"Negative hole rejected");
            bad=Spec(1); bad.Pie.Order="Random"; Reject(bad,"Unknown order rejected");
            bad=Spec(1); bad.Pie.LabelContent="Unknown"; Reject(bad,"Unknown labels rejected");
            bad=Spec(1); bad.Pie.PercentDigits=7; Reject(bad,"Precision bound");
            bad=Spec(1); bad.Series.Add("Series"); Reject(bad,"Matrix header rejected");
            bad=Spec(1); bad.Rows[0].Values.Add(1); Reject(bad,"Hidden matrix row rejected");
            bad=Spec(1); bad.Rows[0].Secondary=2; Reject(bad,"Ambiguous secondary rejected");
            bad=Spec(1); bad.Rows[0].Total=true; Reject(bad,"Ambiguous total rejected");
            bad=Spec(1); bad.Appearance=new ChartAppearance(); Reject(bad,"Imported appearance rejected");
            bad=Spec(1); bad.Waterfall=new WaterfallSpec(); Reject(bad,"Waterfall model rejected");
            bad=Spec(1); bad.Pie.Slices.Add(new PieSliceStyle {Index=1}); Reject(bad,"Invalid style index");
            bad=Spec(1); bad.Pie.Slices.Add(new PieSliceStyle()); bad.Pie.Slices.Add(new PieSliceStyle()); Reject(bad,"Duplicate style index");
            bad=Spec(1); bad.Pie.Slices.Add(new PieSliceStyle {Explode=.5001}); Reject(bad,"Explosion bound");
            bad=Spec(1); bad.Pie.Slices.Add(new PieSliceStyle {Explode=Double.NaN}); Reject(bad,"Explosion finite");
            bad=Spec(1); bad.Pie.Slices.Add(new PieSliceStyle {Color="#  1234"}); Reject(bad,"Strict RGB syntax");
            bad=Spec(1); bad.Pie.CenterTitle="Center"; Reject(bad,"Center title requires hole");
            bad=Spec(1); bad.Kind="Column"; Reject(bad,"Pie options on other kind rejected");
            Console.WriteLine("PieTests: "+checks+" checks passed (original geometry/policies; no Office rendering claim).");
            return 0;
        }
        catch(Exception e) {Console.Error.WriteLine(e);return 1;}
    }
}
