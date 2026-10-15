using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using OpenSlideCharts;

internal static class ScatterTests
{
    private static int checks;
    private static void Check(bool ok,string name){if(!ok)throw new Exception(name);checks++;}
    private static void Near(double a,double b,double tolerance,string name){Check(Math.Abs(a-b)<=tolerance,name+": "+a+" vs "+b);}
    private static ChartSpec Spec(bool bubble=false)
    {return new ChartSpec {Kind=bubble?"Bubble":"Scatter",Scatter=new ScatterSpec {ShowLegend=false,ShowLabels=false}};}
    private static ScatterPoint Add(ChartSpec s,double? x,double? y,double? size=null,string group=null,string label="Point")
    {ScatterPoint p=new ScatterPoint {X=x,Y=y,Size=size,Group=group,Label=label};s.Scatter.Points.Add(p);return p;}
    private static List<Primitive> Role(Layout l,string role){return l.Items.FindAll(delegate(Primitive p){return p.Role==role;});}
    private static Layout Build(ChartSpec s,double width=720,double height=400){return ChartEngine.Build(s,width,height);}
    private static Primitive Point(Layout l,string id){return Role(l,"SCATTER_POINT").Find(delegate(Primitive p){return p.EntityId==id;});}
    private static double CX(Primitive p){return p.X+p.W/2;}private static double CY(Primitive p){return p.Y+p.H/2;}
    private static void Reject(ChartSpec s,string reason)
    {bool rejected=false;try{Build(s);}catch(ArgumentException){rejected=true;}catch(NotSupportedException){rejected=true;}Check(rejected,reason);}
    private static bool Finite(double x){return !Double.IsNaN(x)&&!Double.IsInfinity(x);}
    private static void Coordinate(double x,double y,double width,double height)
    {Check(Finite(x)&&Finite(y)&&x>=-1e-7&&y>=-1e-7&&x<=width+1e-7&&y<=height+1e-7,"Finite coordinates in frame");}
    private static bool Clean(string text)
    {
        foreach(char c in text)if(Char.IsControl(c)||c=='\u2028'||c=='\u2029')return false;
        for(int i=0;i<text.Length;i++)if(Char.IsHighSurrogate(text[i])){if(i+1>=text.Length||!Char.IsLowSurrogate(text[++i]))return false;}else if(Char.IsLowSurrogate(text[i]))return false;
        return true;
    }
    private static void Frame(Layout l,double width=720,double height=400)
    {
        foreach(Primitive p in l.Items)
        {
            Coordinate(p.X,p.Y,width,height);Coordinate(p.X+p.W,p.Y+p.H,width,height);
            if(p.Path!=null)
            {
                Coordinate(p.Path.StartX,p.Path.StartY,width,height);
                foreach(PathSegment segment in p.Path.Segments){Coordinate(segment.X,segment.Y,width,height);if(segment.Curve){Coordinate(segment.C1X,segment.C1Y,width,height);Coordinate(segment.C2X,segment.C2Y,width,height);}}
                PathSegment last=p.Path.Segments[p.Path.Segments.Count-1];Check(last.X==p.Path.StartX&&last.Y==p.Path.StartY,"Marker path explicitly closed");
            }
            if(p.Role=="SCATTER_LABEL"||p.Role=="SCATTER_SECOND_LABEL")
            {Check(Clean(p.Text),"Sanitized marker label");Check(p.Text.Length*p.FontSize*1.2<=p.W+1e-7,"Conservative label width");}
        }
    }
    private static ChartSpec Basic(bool bubble=false)
    {ChartSpec s=Spec(bubble);Add(s,1,2,bubble?(double?)1:null,"A");Add(s,4,7,bubble?(double?)4:null,"B");return s;}
    private static void OnTrend(Layout layout,double x,double y,string reason)
    {
        List<Primitive> lines=Role(layout,"SCATTER_TREND");Check(lines.Count==1,reason+" trend exists");
        Primitive line=lines[0];double length=Math.Sqrt(line.W*line.W+line.H*line.H);
        Near(((x-line.X)*line.H-(y-line.Y)*line.W)/length,0,1e-7,reason+" passes expected point");
    }
    private static void AdjacentTrend(double xOrigin,double xStep,double yOrigin,double yStep,int count,bool reverse)
    {
        ChartSpec s=Spec();
        for(int i=0;i<count;i++)Add(s,xOrigin+i*xStep,yOrigin+i*yStep);
        s.Scatter.XAxis.Reversed=reverse;s.Scatter.YAxis.Reversed=reverse;
        s.Scatter.Trendlines.Add(new ScatterTrendline {ShowEquation=true});
        Layout result=Build(s);Frame(result);
        for(int i=0;i<count;i++)
        {
            Primitive p=Point(result,i.ToString(CultureInfo.InvariantCulture));
            Check(p!=null,"Adjacent coordinate retained");OnTrend(result,CX(p),CY(p),"Adjacent input "+i);
            if(i>1)
            {
                Primitive previous=Point(result,(i-1).ToString(CultureInfo.InvariantCulture));
                Primitive first=Point(result,"0"),second=Point(result,"1");
                Near(CX(p)-CX(previous),CX(second)-CX(first),1e-7,"Adjacent X affine spacing");
                Near(CY(p)-CY(previous),CY(second)-CY(first),1e-7,"Adjacent Y affine spacing");
            }
        }
    }

    public static int Main()
    {
        try
        {
            ChartSpec scatter=Basic();Layout normal=Build(scatter);Frame(normal);
            Check(Role(normal,"SCATTER_POINT").Count==2,"Two scalar XY points");
            Check(CX(Point(normal,"0"))<CX(Point(normal,"1"))&&CY(Point(normal,"0"))>CY(Point(normal,"1")),"Independent axis directions");
            Near(Point(normal,"0").W,8,1e-12,"Default marker size");
            scatter.Scatter.XAxis.Reversed=true;scatter.Scatter.YAxis.Reversed=true;
            Layout reversed=Build(scatter);Check(CX(Point(reversed,"0"))>CX(Point(reversed,"1"))&&CY(Point(reversed,"0"))<CY(Point(reversed,"1")),"Both axes reverse independently");
            scatter.Scatter.XAxis=new ScatterAxisSpec {Minimum=0,Maximum=10};scatter.Scatter.YAxis=new ScatterAxisSpec {Minimum=0,Maximum=10};
            Layout fixedAxes=Build(scatter);Primitive xa=Role(fixedAxes,"SCATTER_X_AXIS")[0],ya=Role(fixedAxes,"SCATTER_Y_AXIS")[0];
            Near(CX(Point(fixedAxes,"0")),xa.X+xa.W*.1,1e-9,"Fixed X coordinate");
            Near(CY(Point(fixedAxes,"0")),ya.Y+ya.H*.8,1e-9,"Fixed Y coordinate");
            Add(scatter,100,100);Check(Role(Build(scatter),"SCATTER_POINT").Count==2,"Outside fixed domain points clipped by center");
            Check(scatter.Scatter.Points.Count==3,"Clipping does not remove input");

            ChartSpec incomplete=Basic();Layout before=Build(incomplete);
            Add(incomplete,1e100,null);Add(incomplete,null,-1e100);Add(incomplete,null,null);
            Layout after=Build(incomplete);Check(Role(after,"SCATTER_POINT").Count==2,"Incomplete XY rows retained but not drawn");
            Near(CX(Point(after,"0")),CX(Point(before,"0")),1e-9,"Incomplete rows do not change drawable X domain");
            Near(CY(Point(after,"0")),CY(Point(before,"0")),1e-9,"Incomplete rows do not change drawable Y domain");
            Check(incomplete.Scatter.Points.Count==5,"Missing rows preserved");
            ChartSpec colors=Spec();Add(colors,null,null,null,"first");Add(colors,1,1,null,"second");Add(colors,2,2,null,"first");Add(colors,3,3,null,"second");
            Layout palette=Build(colors);Check(Point(palette,"1").Color==Point(palette,"3").Color&&Point(palette,"1").Color!=Point(palette,"2").Color,"Group colors stable first occurrence including incomplete rows");
            string[] markers={"Circle","Square","Triangle","Diamond","Plus","Cross","None"};
            ChartSpec shapes=Spec();for(int i=0;i<markers.Length;i++)Add(shapes,i,i,null,"G"+i).Marker=markers[i];
            Layout shapeLayout=Build(shapes);Frame(shapeLayout);Check(Role(shapeLayout,"SCATTER_POINT").Count==6,"All native marker shapes, explicit None hidden");
            Check(Point(shapeLayout,"0").Kind=="ellipse"&&Point(shapeLayout,"1").Kind=="rect"&&Point(shapeLayout,"2").Kind=="path","Marker primitive types");
            shapes.Scatter.MarkerScheme="Shapes";foreach(ScatterPoint p in shapes.Scatter.Points)p.Marker=null;
            shapeLayout=Build(shapes);Check(Point(shapeLayout,"0").Kind==Point(shapeLayout,"6").Kind,"Group marker cycle repeats stably");
            shapes.Scatter.MarkerScheme="None";shapes.Scatter.ShowLabels=true;
            Layout noMarkers=Build(shapes);Check(Role(noMarkers,"SCATTER_POINT").Count==0&&Role(noMarkers,"SCATTER_LABEL").Count==7,"No markers retains labels");
            shapes.Scatter.Points[0].Marker="Diamond";shapes.Scatter.Points[0].Color="#ABCDEF";
            Primitive overridden=Point(Build(shapes),"0");Check(overridden.Kind=="path"&&overridden.Color==(0xab|(0xcd<<8)|(0xef<<16)),"Point overrides scheme and packed RGB");

            ChartSpec bubble=Basic(true);Layout areas=Build(bubble);Frame(areas);
            Near(Point(areas,"1").W/Point(areas,"0").W,2,1e-10,"Area proportional bubble diameters");
            Check(Role(areas,"SCATTER_POINT")[0].EntityId=="1"&&Role(areas,"SCATTER_POINT")[1].EntityId=="0","Larger bubbles behind smaller");
            Check(bubble.Scatter.Points[0].Size==1&&bubble.Scatter.Points[1].Size==4,"Bubble sorting never mutates rows");
            bubble.Scatter.BubbleSizing="Diameter";Layout diameters=Build(bubble);
            Near(Point(diameters,"1").W/Point(diameters,"0").W,4,1e-10,"Diameter proportional sizing");
            ChartSpec edgeBubbles=Spec(true);Add(edgeBubbles,0,0,1);Add(edgeBubbles,1,1,1);
            edgeBubbles.Scatter.MaxBubbleDiameter=200;
            edgeBubbles.Scatter.XAxis=new ScatterAxisSpec {Minimum=0,Maximum=1};
            edgeBubbles.Scatter.YAxis=new ScatterAxisSpec {Minimum=0,Maximum=1};
            Layout edges=Build(edgeBubbles);Frame(edges);
            Near(Point(edges,"0").W,200,1e-10,"Maximum bubble diameter retained at fixed lower bound");
            Near(Point(edges,"1").W,200,1e-10,"Maximum bubble diameter retained at fixed upper bound");
            edgeBubbles.Scatter.XAxis.Reversed=true;edgeBubbles.Scatter.YAxis.Reversed=true;Frame(Build(edgeBubbles));
            bubble.Scatter.BubbleSizing="Area";Add(bubble,null,null,100);Add(bubble,2,3,0);Add(bubble,3,4,null);
            Layout sizeOnly=Build(bubble);Check(Role(sizeOnly,"SCATTER_POINT").Count==2,"Size-only, zero and missing sizes produce no markers");
            Near(Point(sizeOnly,"0").W,5,1e-10,"Size-only row sets common size scale");
            Near(Point(sizeOnly,"1").W,10,1e-10,"Size-only global normalization");
            ChartSpec tiny=Spec(true);Add(tiny,0,0,1e-12);Add(tiny,1,1,1);
            Near(Point(Build(tiny),"0").W,5e-5,1e-15,"Microbubble has no invented minimum");
            tiny.Scatter.Points[0].Size=Double.Epsilon;tiny.Scatter.Points[1].Size=Double.MaxValue;
            Layout underflow=Build(tiny);Frame(underflow);Check(Role(underflow,"SCATTER_POINT").Count==1,"Below coordinate resolution bubble omitted");
            bubble.Scatter.MarkerScheme="None";bubble.Scatter.ShowLabels=true;
            Check(Role(Build(bubble),"SCATTER_POINT").Count==0&&Role(Build(bubble),"SCATTER_LABEL").Count==2,"Bubble None keeps eligible labels");

            double[][] extremes={new double[]{-Double.MaxValue,Double.MaxValue},new double[]{Double.Epsilon,Double.Epsilon*2},new double[]{0,0},new double[]{Double.MaxValue,Double.MaxValue},new double[]{Double.Epsilon,Double.Epsilon}};
            foreach(double[] values in extremes)
            {
                ChartSpec e=Spec();Add(e,values[0],values[1]);Add(e,values[1],values[0]);Layout result=Build(e,320,240);Frame(result,320,240);
                foreach(Primitive tick in result.Items.FindAll(delegate(Primitive p){return p.Role=="SCATTER_X_TICK"||p.Role=="SCATTER_Y_TICK";}))
                {double parsed;Check(Double.TryParse(tick.Text,NumberStyles.Float,CultureInfo.InvariantCulture,out parsed)&&Finite(parsed),"Finite tick text for extremes: "+tick.Text);}
            }
            ChartSpec log=Spec();Add(log,1,1);Add(log,10,100);Add(log,100,10000);
            log.Scatter.XAxis.Logarithmic=true;log.Scatter.YAxis.Logarithmic=true;
            Layout logarithmic=Build(log);Frame(logarithmic);
            Near(CX(Point(logarithmic,"1")),(CX(Point(logarithmic,"0"))+CX(Point(logarithmic,"2")))/2,1e-8,"Log X equal decade spacing");
            Near(CY(Point(logarithmic,"1")),(CY(Point(logarithmic,"0"))+CY(Point(logarithmic,"2")))/2,1e-8,"Log Y independent decades");
            log.Scatter.Points[0].X=-1;log.Scatter.Points[0].Y=0;
            Layout warning=Build(log,320,240);Frame(warning,320,240);
            Check(Role(warning,"SCATTER_WARNING").Count==1,"Nonpositive log coordinates warn");
            Near(CX(Point(warning,"0")),Role(warning,"SCATTER_X_AXIS")[0].X,1e-8,"Nonpositive log X maps lower bound");
            Near(CY(Point(warning,"0")),Role(warning,"SCATTER_X_AXIS")[0].Y,1e-8,"Nonpositive log Y maps lower bound");
            Check(log.Scatter.Points[0].X==-1&&log.ImportWarnings.Count==0,"Log fallback does not mutate data or warning collection");
            log.Scatter.XAxis.Reversed=true;warning=Build(log);
            xa=Role(warning,"SCATTER_X_AXIS")[0];Near(CX(Point(warning,"0")),xa.X+xa.W,1e-8,"Log baseline respects reversed axis");
            ChartSpec logExtreme=Spec();Add(logExtreme,Double.Epsilon,Double.MaxValue);Add(logExtreme,Double.MaxValue,Double.Epsilon);
            logExtreme.Scatter.XAxis.Logarithmic=true;logExtreme.Scatter.YAxis.Logarithmic=true;Frame(Build(logExtreme));

            ChartSpec lines=Basic();lines.Scatter.XAxis=new ScatterAxisSpec {Minimum=0,Maximum=10};lines.Scatter.YAxis=new ScatterAxisSpec {Minimum=0,Maximum=10};
            lines.Scatter.Partitions.Add(new ScatterPartition {X1=4,Y1=2,X2=4,Y2=4,Label="Vertical"});
            lines.Scatter.Partitions.Add(new ScatterPartition {X1=2,Y1=5,X2=3,Y2=5});
            lines.Scatter.Partitions.Add(new ScatterPartition {X1=-1e100,Y1=-1e100,X2=1e100,Y2=1e100});
            Layout partitioned=Build(lines);Frame(partitioned);List<Primitive> partitions=Role(partitioned,"SCATTER_PARTITION");
            Check(partitions.Count==3,"Three infinite partitions intersect plot");
            ya=Role(partitioned,"SCATTER_Y_AXIS")[0];xa=Role(partitioned,"SCATTER_X_AXIS")[0];
            Near(partitions[0].W,0,1e-8,"Vertical partition");Near(Math.Abs(partitions[0].H),ya.H,1e-8,"Vertical partition extends beyond input segment");
            Near(partitions[1].H,0,1e-8,"Horizontal partition");Near(Math.Abs(partitions[1].W),xa.W,1e-8,"Horizontal partition clipped full width");
            Near(Math.Abs(partitions[2].W),xa.W,1e-8,"Extreme diagonal partition full plot");
            lines.Scatter.Partitions.Add(new ScatterPartition {X1=20,Y1=1,X2=20,Y2=2});Check(Role(Build(lines),"SCATTER_PARTITION").Count==3,"Outside infinite partition omitted");
            ChartSpec trends=Spec();for(int i=0;i<4;i++)Add(trends,i,2*i+1,null,"A");
            trends.Scatter.Trendlines.Add(new ScatterTrendline {Group="A",ShowEquation=true});trends.Scatter.Trendlines.Add(new ScatterTrendline {Group="Missing"});
            Layout fitted=Build(trends);Frame(fitted);Check(Role(fitted,"SCATTER_TREND").Count==1,"Grouped trendline; absent group silently not drawn");
            Check(Role(fitted,"SCATTER_TREND_LABEL")[0].Text.Contains("R\u00b2 = 1")&&Role(fitted,"SCATTER_TREND_LABEL")[0].Text.Contains("y = 2 x + 1"),"Equation and R squared");
            ChartSpec hugeSlope=Spec();Add(hugeSlope,Double.Epsilon,0);Add(hugeSlope,Double.Epsilon*2,Double.MaxValue/2);Add(hugeSlope,Double.Epsilon*3,Double.MaxValue);
            hugeSlope.Scatter.Trendlines.Add(new ScatterTrendline {ShowEquation=true});Layout normalized=Build(hugeSlope);Frame(normalized);
            Check(Role(normalized,"SCATTER_TREND_LABEL")[0].Text.Contains("y / ")&&!Role(normalized,"SCATTER_TREND_LABEL")[0].Text.Contains("Infinity"),"Overflow coefficients use normalized equation");
            ChartSpec remote=Spec();Add(remote,0,.5,null,"local");Add(remote,1e307,.5,null,"far");Add(remote,1e308,.5,null,"far");
            remote.Scatter.XAxis=new ScatterAxisSpec {Minimum=0,Maximum=1};remote.Scatter.YAxis=new ScatterAxisSpec {Minimum=0,Maximum=1};remote.Scatter.Trendlines.Add(new ScatterTrendline {Group="far",ShowRSquared=false});
            Layout remoteFit=Build(remote);Check(Role(remoteFit,"SCATTER_TREND").Count==1,"Far-anchor horizontal trend still clips through fixed plot");Frame(remoteFit);
            ChartSpec weighted=Basic(true);Add(weighted,7,8,3,"A");weighted.Scatter.Trendlines.Add(new ScatterTrendline());Frame(Build(weighted));
            AdjacentTrend(1e16,2,0,1,2,false);
            AdjacentTrend(1e16,2,1e16,2,3,false);
            AdjacentTrend(-1e16,2,1e16,2,4,true);
            AdjacentTrend(0,Double.Epsilon,0,Double.Epsilon,4,false);
            AdjacentTrend(-Double.MaxValue,Double.MaxValue,Double.MaxValue,-Double.MaxValue,2,false);
            ChartSpec oppositeExtremes=Spec();Add(oppositeExtremes,-Double.MaxValue,Double.MaxValue);Add(oppositeExtremes,0,0);Add(oppositeExtremes,Double.MaxValue,-Double.MaxValue);
            oppositeExtremes.Scatter.Trendlines.Add(new ScatterTrendline {ShowEquation=true});
            Layout oppositeFit=Build(oppositeExtremes);Frame(oppositeFit);
            foreach(Primitive p in Role(oppositeFit,"SCATTER_POINT"))OnTrend(oppositeFit,CX(p),CY(p),"Opposite MaxValue regression");
            Check(Role(oppositeFit,"SCATTER_TREND_LABEL")[0].Text.Contains("y = -1 x + 0"),"Exact raw coefficient display at extremes");
            ChartSpec adjacentFixed=Spec();Add(adjacentFixed,1e16,1e16);Add(adjacentFixed,1e16+2,1e16+2);
            adjacentFixed.Scatter.XAxis=new ScatterAxisSpec {Minimum=1e16,Maximum=1e16+2};
            adjacentFixed.Scatter.YAxis=new ScatterAxisSpec {Minimum=1e16,Maximum=1e16+2};
            adjacentFixed.Scatter.Trendlines.Add(new ScatterTrendline());
            Layout adjacentBounds=Build(adjacentFixed);Frame(adjacentBounds);
            foreach(Primitive p in Role(adjacentBounds,"SCATTER_POINT"))OnTrend(adjacentBounds,CX(p),CY(p),"Adjacent explicit bounds regression");
            Near(CX(Point(adjacentBounds,"0")),Role(adjacentBounds,"SCATTER_X_AXIS")[0].X,1e-9,"Adjacent explicit minimum exact position");
            Near(CX(Point(adjacentBounds,"1")),Role(adjacentBounds,"SCATTER_X_AXIS")[0].X+Role(adjacentBounds,"SCATTER_X_AXIS")[0].W,1e-9,"Adjacent explicit maximum exact position");
            ChartSpec adjacentWeighted=Spec(true);Add(adjacentWeighted,1e16,0,1);Add(adjacentWeighted,1e16+2,1,2);Add(adjacentWeighted,1e16+4,0,3);
            adjacentWeighted.Scatter.Trendlines.Add(new ScatterTrendline());
            Layout weightedFit=Build(adjacentWeighted);Frame(weightedFit);
            Primitive weightedFirst=Point(weightedFit,"0"),weightedMiddle=Point(weightedFit,"1"),weightedLast=Point(weightedFit,"2");
            double weightedYSpan=CY(weightedFirst)-CY(weightedMiddle);
            OnTrend(weightedFit,CX(weightedFirst),CY(weightedFirst)-weightedYSpan*.6,"Weighted adjacent regression at first X");
            OnTrend(weightedFit,CX(weightedMiddle),CY(weightedFirst)-weightedYSpan*.4,"Weighted adjacent regression at middle X");
            OnTrend(weightedFit,CX(weightedLast),CY(weightedFirst)-weightedYSpan*.2,"Weighted adjacent regression at last X");

            ChartSpec labels=Spec();Add(labels,1,1,null,null,new String('W',100)+"\t\u754c\ud83d\ude00");Add(labels,2,2,null,null,"B\nC");
            labels.Scatter.ShowLabels=true;labels.Scatter.SecondLabelContent="XY";Layout labeled=Build(labels);Frame(labeled);
            Check(Role(labeled,"SCATTER_LABEL").Count==2&&Role(labeled,"SCATTER_SECOND_LABEL").Count==2,"Two labels per point");
            Check(Role(labeled,"SCATTER_LABEL")[0].Text.EndsWith("\u2026"),"Wide category abbreviated");
            Check(Role(labeled,"SCATTER_SECOND_LABEL")[0].Text=="(1, 1)","Numeric label complete");
            List<Primitive> labelBoxes=Role(labeled,"SCATTER_LABEL");labelBoxes.AddRange(Role(labeled,"SCATTER_SECOND_LABEL"));
            for(int i=0;i<labelBoxes.Count;i++)for(int j=i+1;j<labelBoxes.Count;j++)
            {Primitive a=labelBoxes[i],b=labelBoxes[j];Check(!(a.X<b.X+b.W&&b.X<a.X+a.W&&a.Y<b.Y+b.H&&b.Y<a.Y+a.H),"Sparse label placements do not overlap");}
            ChartSpec many=Spec();many.Scatter.ShowLabels=null;for(int i=0;i<301;i++)Add(many,i,i);
            Check(Role(Build(many),"SCATTER_LABEL").Count==0,"More than 300 rows disables labels by default");
            many.Scatter.Points.RemoveAt(300);Check(Role(Build(many),"SCATTER_LABEL").Count==300,"300 rows keeps default labels");
            many.Scatter.ShowLabels=false;for(int i=300;i<1000;i++)Add(many,i,i);
            Layout thousand=Build(many);Check(Role(thousand,"SCATTER_POINT").Count==1000,"1000 points supported");
            Check(many.Scatter.Points[0].X==0&&many.Scatter.Points[999].X==999,"No data order mutation");
            Add(many,1001,1001);Reject(many,"1001 points rejected");

            ChartSpec invalid=Basic();invalid.Scatter.Points[0].Size=0;Reject(invalid,"Scatter size even zero rejected");
            invalid=Basic(true);invalid.Scatter.Points[0].Size=-1;Reject(invalid,"Negative size policy");
            invalid=Spec();Add(invalid,1,null);Reject(invalid,"No drawable coordinates rejected");
            invalid=Spec(true);Add(invalid,1,2,0);Reject(invalid,"No positive bubble size rejected");
            invalid=Basic();invalid.Scatter.Points[0].X=Double.NaN;Reject(invalid,"NaN coordinate rejected");
            invalid=Basic(true);invalid.Scatter.Points[0].Size=Double.PositiveInfinity;Reject(invalid,"Infinite size rejected");
            invalid=Basic();invalid.Scatter.XAxis.Minimum=2;invalid.Scatter.XAxis.Maximum=1;Reject(invalid,"Inverted fixed bounds rejected");
            invalid=Basic();invalid.Scatter.XAxis.Logarithmic=true;invalid.Scatter.XAxis.Minimum=0;Reject(invalid,"Nonpositive fixed log bound rejected");
            invalid=Spec();Add(invalid,0,1);invalid.Scatter.XAxis.Logarithmic=true;Reject(invalid,"Log axis without positive data rejected");
            invalid=Basic();invalid.Scatter.XAxis.Logarithmic=true;invalid.Scatter.Trendlines.Add(new ScatterTrendline());Reject(invalid,"Log trend rejected");
            invalid=Basic();invalid.Scatter.YAxis.Logarithmic=true;invalid.Scatter.Partitions.Add(new ScatterPartition {X2=1});Reject(invalid,"Log partition rejected");
            invalid=Basic();invalid.Scatter.Partitions.Add(new ScatterPartition());Reject(invalid,"Coincident partition points rejected");
            invalid=Basic();invalid.Scatter.Points[0].Color="#  1234";Reject(invalid,"Strict point color");
            invalid=Basic();invalid.Scatter.Points[0].Marker="Star";Reject(invalid,"Unknown marker");
            invalid=Basic(true);invalid.Scatter.MarkerScheme="Shapes";Reject(invalid,"Bubble Shapes rejected");
            invalid=Basic(true);invalid.Scatter.Points[0].Marker="Square";Reject(invalid,"Bubble noncircular override rejected");
            invalid=Basic();invalid.Scatter.MarkerSize=0;Reject(invalid,"MarkerSize positive");
            invalid=Basic();invalid.Scatter.LabelFontSize=5;Reject(invalid,"Label minimum size");
            invalid=Basic();invalid.Scatter.SecondLabelContent="Group";Reject(invalid,"Unknown second label mode");
            invalid=Basic();invalid.Rows.Add(new ChartRow());Reject(invalid,"Rows ambiguity rejected");
            invalid=Basic();invalid.Series.Add("A");Reject(invalid,"Series ambiguity rejected");
            invalid=Basic();invalid.Appearance=new ChartAppearance();Reject(invalid,"Appearance ambiguity rejected");
            invalid=Basic();invalid.Pie=new PieOptions();Reject(invalid,"Pie ambiguity rejected");
            invalid=Basic();invalid.Waterfall=new WaterfallSpec();Reject(invalid,"Waterfall ambiguity rejected");
            invalid=Basic();invalid.Kind="Pie";Reject(invalid,"Scatter options on different kind rejected");
            ChartSpec dataOnly=Basic();dataOnly.Scatter.ShowLegend=true;for(int i=0;i<100;i++)Add(dataOnly,i,i,null,"Group "+i);
            typeof(ChartEngine).GetMethod("ValidateScatter",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{dataOnly});
            Check(true,"Data validation independent of frame");Reject(dataOnly,"Legend too tall fails actual frame explicitly");
            Frame(Build(ChartSamples.Create("Scatter"),640,480),640,480);Frame(Build(ChartSamples.Create("Bubble"),640,480),640,480);
            Console.WriteLine("ScatterTests: "+checks+" checks passed; original geometry policies, no Office claim.");return 0;
        }
        catch(Exception e){Console.Error.WriteLine(e);return 1;}
    }
}
