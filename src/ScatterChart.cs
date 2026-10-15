using System;
using System.Collections.Generic;

namespace OpenSlideCharts
{
    public sealed class ScatterSpec
    {
        public List<ScatterPoint> Points = new List<ScatterPoint>();
        public ScatterAxisSpec XAxis = new ScatterAxisSpec(), YAxis = new ScatterAxisSpec();
        public string XTitle, YTitle;
        public bool ShowGrid = true, ShowLegend = true;
        public double MarkerSize = 8, MaxBubbleDiameter = 50;
        public string BubbleSizing = "Area", MarkerScheme = "Colors";
        public bool? ShowLabels;
        public string LabelContent = "Label", SecondLabelContent = "None";
        public double LabelFontSize = 10;
        public List<ScatterTrendline> Trendlines = new List<ScatterTrendline>();
        public List<ScatterPartition> Partitions = new List<ScatterPartition>();
    }
    public sealed class ScatterPoint
    {
        public string Label, Group, Color, Marker;
        public double? X, Y, Size;
    }
    public sealed class ScatterAxisSpec
    {
        public double? Minimum, Maximum;
        public bool Logarithmic, Reversed;
    }
    public sealed class ScatterTrendline
    {
        public string Group, Color;
        public bool ShowRSquared = true, ShowEquation = false;
    }
    public sealed class ScatterPartition
    {
        public double X1, Y1, X2, Y2;
        public string Label, Color;
    }

    public static partial class ChartEngine
    {
        private static readonly string[] ScatterMarkers = { "Circle", "Square", "Triangle", "Diamond", "Plus", "Cross" };
        private sealed class ScatterBox
        {
            internal double X, Y, W, H;
            internal ScatterBox(double x, double y, double w, double h) { X=x; Y=y; W=w; H=h; }
        }
        private sealed class PlottedPoint
        {
            internal ScatterPoint Data;
            internal int Index, Color;
            internal string Marker;
            internal double X, Y, Diameter;
        }
        private sealed class FittedTrend
        {
            internal ScatterFit Fit;
            internal string Text, Id;
            internal int Color;
            internal double TextHeight;
        }

        // Data validation deliberately has no slide-size dependency (Excel refresh
        // validates data here; Build validates the actual chart frame separately).
        internal static void ValidateScatter(ChartSpec spec)
        {
            if (spec == null) throw new ArgumentNullException("spec");
            bool bubble = String.Equals(spec.Kind,"Bubble",StringComparison.OrdinalIgnoreCase);
            if (!bubble && !String.Equals(spec.Kind,"Scatter",StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Scatter validation requires Kind=Scatter or Bubble.");
            if (spec.Pie != null || spec.Waterfall != null || spec.Appearance != null ||
                spec.Rows != null && spec.Rows.Count != 0 || spec.Series != null && spec.Series.Count != 0 ||
                spec.LayoutMode != null && spec.LayoutMode != "Clustered")
                throw new NotSupportedException("Scatter/Bubble require their own point data without rows, series, appearance, pie or waterfall options.");
            ScatterSpec s=spec.Scatter;
            if (s==null || s.Points==null || s.Points.Count<1 || s.Points.Count>1000)
                throw new ArgumentException("Scatter/Bubble require between 1 and 1000 points.");
            if (!Finite(s.MarkerSize) || s.MarkerSize<=0 || s.MarkerSize>100 || !Finite(s.MaxBubbleDiameter) || s.MaxBubbleDiameter<=0 || s.MaxBubbleDiameter>200)
                throw new ArgumentException("MarkerSize must be in (0,100] and MaxBubbleDiameter in (0,200].");
            if (!Finite(s.LabelFontSize) || s.LabelFontSize<6 || s.LabelFontSize>24)
                throw new ArgumentException("Scatter label font size must be between 6 and 24 points.");
            if (s.BubbleSizing!="Area" && s.BubbleSizing!="Diameter") throw new ArgumentException("BubbleSizing must be Area or Diameter.");
            if (s.MarkerScheme!="Colors" && s.MarkerScheme!="Shapes" && s.MarkerScheme!="None") throw new ArgumentException("Unknown marker scheme.");
            if (bubble && s.MarkerScheme=="Shapes") throw new NotSupportedException("Bubble charts use circular markers; Shapes is a scatter-only scheme.");
            if (!ScatterContent(s.LabelContent) || !ScatterContent(s.SecondLabelContent)) throw new ArgumentException("Unknown scatter label content.");
            ScatterAxisSpec x=s.XAxis??new ScatterAxisSpec(),y=s.YAxis??new ScatterAxisSpec();
            ValidateScatterAxis(x); ValidateScatterAxis(y);
            bool drawable=false,positiveX=false,positiveY=false;
            foreach(ScatterPoint p in s.Points)
            {
                if(p==null || p.X.HasValue&&!Finite(p.X.Value) || p.Y.HasValue&&!Finite(p.Y.Value) || p.Size.HasValue&&!Finite(p.Size.Value))
                    throw new ArgumentException("Scatter coordinates and sizes must be finite when present.");
                if(!bubble && p.Size.HasValue) throw new NotSupportedException("A Size value requires Kind=Bubble.");
                if(bubble && p.Size.HasValue && p.Size.Value<0) throw new ArgumentException("Bubble sizes must be nonnegative.");
                if(p.Color!=null) PieColor(p.Color);
                if(p.Marker!=null && p.Marker!="None" && Array.IndexOf(ScatterMarkers,p.Marker)<0) throw new ArgumentException("Unknown point marker.");
                if(bubble && p.Marker!=null && p.Marker!="None" && p.Marker!="Circle") throw new NotSupportedException("Bubble point markers must be Circle or None.");
                if(ScatterEligible(p,bubble)) { drawable=true; positiveX|=p.X.Value>0; positiveY|=p.Y.Value>0; }
            }
            if(!drawable) throw new ArgumentException("Provide at least one point with X and Y (and positive Size for Bubble).");
            if(x.Logarithmic&&!positiveX || y.Logarithmic&&!positiveY) throw new ArgumentException("Each logarithmic axis requires at least one positive drawable coordinate.");
            if((x.Logarithmic||y.Logarithmic) && (s.Trendlines!=null&&s.Trendlines.Count>0 || s.Partitions!=null&&s.Partitions.Count>0))
                throw new NotSupportedException("Trendlines and partitions require two linear axes.");
            if(s.Trendlines!=null) foreach(ScatterTrendline t in s.Trendlines)
            { if(t==null) throw new ArgumentException("Trendline definitions cannot be null."); if(t.Color!=null) PieColor(t.Color); }
            if(s.Partitions!=null) foreach(ScatterPartition p in s.Partitions)
            {
                if(p==null || !Finite(p.X1)||!Finite(p.Y1)||!Finite(p.X2)||!Finite(p.Y2) || p.X1==p.X2&&p.Y1==p.Y2)
                    throw new ArgumentException("Partitions require two distinct finite data points.");
                if(p.Color!=null) PieColor(p.Color);
            }
        }
        private static bool ScatterEligible(ScatterPoint p,bool bubble)
        { return p.X.HasValue&&p.Y.HasValue&&(!bubble || p.Size.HasValue&&p.Size.Value>0); }
        private static bool ScatterContent(string s)
        { return s=="None"||s=="Label"||s=="X"||s=="Y"||s=="XY"||s=="Size"; }
        private static void ValidateScatterAxis(ScatterAxisSpec a)
        {
            if(a.Minimum.HasValue&&!Finite(a.Minimum.Value) || a.Maximum.HasValue&&!Finite(a.Maximum.Value)) throw new ArgumentException("Axis bounds must be finite.");
            if(a.Minimum.HasValue&&a.Maximum.HasValue&&a.Minimum.Value>=a.Maximum.Value) throw new ArgumentException("Axis minimum must be less than maximum.");
            if(a.Logarithmic&&(a.Minimum.HasValue&&a.Minimum.Value<=0 || a.Maximum.HasValue&&a.Maximum.Value<=0)) throw new ArgumentException("Logarithmic bounds must be positive.");
        }

        private sealed class ScatterScale
        {
            internal double Origin,Unit,Low,High;
            internal bool Log,Reverse;
            private static readonly double LogMin=Math.Log10(Double.Epsilon), LogMax=Math.Log10(Double.MaxValue);
            internal ScatterScale(List<double> values,ScatterAxisSpec settings)
            {
                settings=settings??new ScatterAxisSpec(); Log=settings.Logarithmic; Reverse=settings.Reversed;
                Origin=Log?0:settings.Minimum??values[0];
                Unit=0;
                foreach(double v in values) Unit=Math.Max(Unit,Math.Abs(v));
                if(settings.Minimum.HasValue) Unit=Math.Max(Unit,Math.Abs(settings.Minimum.Value));
                if(settings.Maximum.HasValue) Unit=Math.Max(Unit,Math.Abs(settings.Maximum.Value));
                if(Unit==0) Unit=1;
                Low=Double.PositiveInfinity; High=Double.NegativeInfinity;
                foreach(double v in values)
                {
                    if(Log&&v<=0) continue;
                    double n=Transform(v);
                    Low=Math.Min(Low,n); High=Math.Max(High,n);
                }
                if(!Finite(Low)||!Finite(High)) throw new ArgumentException("Logarithmic axis has no positive data.");
                if(settings.Minimum.HasValue) Low=Transform(settings.Minimum.Value);
                if(settings.Maximum.HasValue) High=Transform(settings.Maximum.Value);
                double limit=Log?LogMax:Transform(Double.MaxValue), floor=Log?LogMin:Transform(-Double.MaxValue);
                if(High<=Low)
                {
                    double middle=settings.Minimum.HasValue?Low:settings.Maximum.HasValue?High:Low;
                    double pad=Log?.5:Math.Max(Math.Max(Math.Abs(middle)*.1,.1),Double.Epsilon/Unit);
                    if(!settings.Minimum.HasValue) Low=Math.Max(floor,middle-pad);
                    if(!settings.Maximum.HasValue) High=Math.Min(limit,middle+pad);
                }
                else
                {
                    double pad=(High-Low)*.05;
                    if(!settings.Minimum.HasValue) Low=Math.Max(floor,Low-pad);
                    if(!settings.Maximum.HasValue) High=Math.Min(limit,High+pad);
                }
                if(!Finite(Low)||!Finite(High)||High<=Low) throw new ArgumentException("Axis range cannot be represented; adjust the explicit bounds.");
            }
            internal double Transform(double v)
            {
                if(Log)return Math.Log10(v);
                double delta=v-Origin;
                return Finite(delta)?delta/Unit:v/Unit-Origin/Unit;
            }
            internal double Fraction(double v)
            {
                double n=Log&&v<=0?Low:Transform(v);
                double f=(n-Low)/(High-Low);
                return Reverse?1-f:f;
            }
            internal double Tick(double fraction)
            {
                if(Reverse) fraction=1-fraction;
                double n=(1-fraction)*Low+fraction*High;
                if(!Log)
                {
                    double value=Origin+n*Unit;
                    if(!Finite(value))value=(Origin/Unit+n)*Unit;
                    return Finite(value)?value:Math.Sign(Origin/Unit+n)*Double.MaxValue;
                }
                if(n<=LogMin)return Double.Epsilon;if(n>=LogMax)return Double.MaxValue;
                double result=Math.Pow(10,n);return result==0?Double.Epsilon:Double.IsInfinity(result)?Double.MaxValue:result;
            }
        }

        private static void BuildScatter(Layout layout,ChartSpec spec,double width,double height)
        {
            ScatterSpec s=spec.Scatter; bool bubble=String.Equals(spec.Kind,"Bubble",StringComparison.OrdinalIgnoreCase);
            List<double> xs=new List<double>(),ys=new List<double>();
            List<string> groups=new List<string>();Dictionary<string,int> groupIndexes=new Dictionary<string,int>(StringComparer.Ordinal);
            double maximumSize=0;
            bool logWarning=false;
            ScatterAxisSpec xOptions=s.XAxis??new ScatterAxisSpec(),yOptions=s.YAxis??new ScatterAxisSpec();
            foreach(ScatterPoint p in s.Points)
            {
                string group=p.Group??"";
                if(!groupIndexes.ContainsKey(group)){groupIndexes.Add(group,groups.Count);groups.Add(group);}
                if(p.Size.HasValue)maximumSize=Math.Max(maximumSize,p.Size.Value);
                if(!ScatterEligible(p,bubble))continue;
                xs.Add(p.X.Value);ys.Add(p.Y.Value);
                logWarning|=xOptions.Logarithmic&&p.X.Value<=0 || yOptions.Logarithmic&&p.Y.Value<=0;
            }
            ScatterScale xAxis=new ScatterScale(xs,xOptions),yAxis=new ScatterScale(ys,yOptions);
            double header=logWarning?58:42;
            if(logWarning) ScatterText(layout,"SCATTER_WARNING",null,"Log axis: values <= 0 shown at minimum.",8,38,width-16,14,6,Red,"left",true);
            int columns=Math.Max(1,(int)((width-20)/120)),legendRows=s.ShowLegend?(groups.Count+columns-1)/columns:0;
            if(s.ShowLegend) for(int i=0;i<groups.Count;i++)
            {
                double cell=(width-20)/columns, x=10+(i%columns)*cell,y=header+(i/columns)*18;
                string marker=s.MarkerScheme=="None"?"None":s.MarkerScheme=="Shapes"?ScatterMarkers[i%ScatterMarkers.Length]:"Circle";
                AddScatterMarker(layout,marker,x+5,y+7,8,SeriesColors[i%SeriesColors.Length],"SCATTER_LEGEND_MARKER",i.ToString(Invariant));
                ScatterText(layout,"SCATTER_LEGEND",i.ToString(Invariant),FitPieText(CleanPieText(groups[i].Length==0?"(ungrouped)":groups[i]),cell-20,8),x+14,y,cell-20,14,8,Ink,"left",true);
            }
            header+=legendRows*18;
            List<FittedTrend> trends=new List<FittedTrend>();
            if(s.Trendlines!=null) for(int i=0;i<s.Trendlines.Count;i++)
            {
                ScatterTrendline definition=s.Trendlines[i];ScatterFit fit=ScatterRegression.Fit(s.Points,bubble,definition.Group);
                if(fit==null)continue;
                string text=ScatterTrendText(fit,definition);
                FittedTrend trend=new FittedTrend {Fit=fit,Id=i.ToString(Invariant),Text=text,Color=definition.Color==null?Ink:PieColor(definition.Color)};
                if(text.Length!=0)
                {
                    trend.TextHeight=Math.Ceiling(text.Length*7*1.2/(width-24))*11;
                    ScatterText(layout,"SCATTER_TREND_LABEL",trend.Id,text,12,header,width-24,trend.TextHeight,7,trend.Color,"left",false);
                    header+=trend.TextHeight;
                }
                trends.Add(trend);
            }
            double radius=(bubble?s.MaxBubbleDiameter:s.MarkerSize)/2;
            double plotX=Math.Max(100,radius+12),plotY=header+radius+22,plotW=width-plotX-Math.Max(50,radius+12),plotH=height-58-radius-plotY;
            if(plotW<60 || plotH<32)throw new ArgumentException("Scatter chart frame is too small for its legend, trend annotations and marker size. Increase the chart frame or simplify these options.");
            ScatterBox plot=new ScatterBox(plotX,plotY,plotW,plotH);
            ScatterAxes(layout,s,xAxis,yAxis,plot,width,height);
            if(s.Partitions!=null) for(int i=0;i<s.Partitions.Count;i++)
            {
                ScatterPartition p=s.Partitions[i];double[] clipped;
                if(ClipScatterLine(xAxis.Fraction(p.X1),yAxis.Fraction(p.Y1),xAxis.Fraction(p.X2),yAxis.Fraction(p.Y2),out clipped))
                {
                    int color=p.Color==null?Gray:PieColor(p.Color);string id=i.ToString(Invariant);
                    ScatterClippedLine(layout,plot,clipped,color,"SCATTER_PARTITION",id);
                    if(!String.IsNullOrEmpty(p.Label))
                    {
                        double boxW=Math.Min(140,plot.W),x=plot.X+(clipped[0]+clipped[2])/2*plot.W-boxW/2;
                        double y=plot.Y+(1-(clipped[1]+clipped[3])/2)*plot.H-8;
                        ScatterText(layout,"SCATTER_PARTITION_LABEL",id,FitPieText(CleanPieText(p.Label),boxW,8),Math.Max(plot.X,Math.Min(plot.X+plot.W-boxW,x)),Math.Max(plot.Y,Math.Min(plot.Y+plot.H-14,y)),boxW,14,8,color,"center",true);
                    }
                }
            }
            foreach(FittedTrend trend in trends)
            {
                ScatterFit fit=trend.Fit;
                double[] clipped=fit.ClipToAxes(xAxis.Origin,xAxis.Unit,xAxis.Low,xAxis.High,yAxis.Origin,yAxis.Unit,yAxis.Low,yAxis.High);
                if(clipped!=null)
                {
                    if(xAxis.Reverse){clipped[0]=1-clipped[0];clipped[2]=1-clipped[2];}
                    if(yAxis.Reverse){clipped[1]=1-clipped[1];clipped[3]=1-clipped[3];}
                    ScatterClippedLine(layout,plot,clipped,trend.Color,"SCATTER_TREND",trend.Id);
                }
            }
            List<PlottedPoint> points=new List<PlottedPoint>();
            for(int i=0;i<s.Points.Count;i++)
            {
                ScatterPoint p=s.Points[i];if(!ScatterEligible(p,bubble))continue;
                double fx=xAxis.Fraction(p.X.Value),fy=yAxis.Fraction(p.Y.Value);
                if(!Finite(fx)||!Finite(fy)||fx<0||fx>1||fy<0||fy>1)continue;
                int group=groupIndexes[p.Group??""];
                string marker=p.Marker??(s.MarkerScheme=="None"?"None":s.MarkerScheme=="Shapes"?ScatterMarkers[group%ScatterMarkers.Length]:"Circle");
                double diameter=s.MarkerSize;
                if(bubble)
                {
                    double ratio=p.Size.Value/maximumSize;
                    diameter=s.MaxBubbleDiameter*(s.BubbleSizing=="Area"?(ratio>0?Math.Sqrt(ratio):Math.Sqrt(p.Size.Value)/Math.Sqrt(maximumSize)):ratio);
                }
                double cx=plot.X+fx*plot.W,cy=plot.Y+(1-fy)*plot.H;
                if(diameter==0 || cx-diameter/2==cx+diameter/2 || cy-diameter/2==cy+diameter/2) marker="None";
                points.Add(new PlottedPoint {Data=p,Index=i,X=cx,Y=cy,Diameter=marker=="None"?0:diameter,Marker=marker,Color=p.Color==null?SeriesColors[group%SeriesColors.Length]:PieColor(p.Color)});
            }
            if(bubble)points.Sort(delegate(PlottedPoint a,PlottedPoint b){int cmp=b.Data.Size.Value.CompareTo(a.Data.Size.Value);return cmp==0?a.Index.CompareTo(b.Index):cmp;});
            foreach(PlottedPoint p in points) AddScatterMarker(layout,p.Marker,p.X,p.Y,p.Diameter,p.Color,"SCATTER_POINT",p.Index.ToString(Invariant));
            if(s.ShowLabels??(s.Points.Count<=300)) ScatterPointLabels(layout,points,s,plot);
        }

        private static void ScatterText(Layout layout,string role,string id,string text,double x,double y,double w,double h,double font,int color,string alignment,bool nowrap)
        {layout.Items.Add(new Primitive {Kind="text",Role=role,EntityId=id,Text=text,X=x,Y=y,W=w,H=h,FontSize=font,Color=color,Alignment=alignment,VerticalAlignment="center",NoWrap=nowrap,MarginLeft=0,MarginRight=0});}
        private static void ScatterLine(Layout l,double x,double y,double w,double h,int color,string role,string id)
        {l.Items.Add(new Primitive {Kind="line",X=x,Y=y,W=w,H=h,Color=color,LineWidth=.75,Role=role,EntityId=id});}
        private static void ScatterAxes(Layout l,ScatterSpec s,ScatterScale x,ScatterScale y,ScatterBox p,double width,double height)
        {
            int xTicks=p.W>=400?5:3;HashSet<string> seenX=new HashSet<string>(),seenY=new HashSet<string>();
            for(int i=0;i<xTicks;i++)
            {
                double f=i/(double)(xTicks-1),at=p.X+f*p.W;string text=ScatterNumber(x.Tick(f),4);
                if(s.ShowGrid)ScatterLine(l,at,p.Y,0,p.H,Gray,"SCATTER_GRID_X",i.ToString(Invariant));
                if(seenX.Add(text))ScatterText(l,"SCATTER_X_TICK",i.ToString(Invariant),text,at-46,p.Y+p.H+8,92,16,Math.Min(8,92/(text.Length*1.2)),Ink,"center",true);
            }
            for(int i=0;i<5;i++)
            {
                double f=i/4.0,at=p.Y+(1-f)*p.H;string text=ScatterNumber(y.Tick(f),4);
                if(s.ShowGrid)ScatterLine(l,p.X,at,p.W,0,Gray,"SCATTER_GRID_Y",i.ToString(Invariant));
                if(seenY.Add(text))ScatterText(l,"SCATTER_Y_TICK",i.ToString(Invariant),text,4,at-8,88,16,Math.Min(8,88/(text.Length*1.2)),Ink,"right",true);
            }
            ScatterLine(l,p.X,p.Y+p.H,p.W,0,Ink,"SCATTER_X_AXIS",null);ScatterLine(l,p.X,p.Y,0,p.H,Ink,"SCATTER_Y_AXIS",null);
            if(!String.IsNullOrEmpty(s.XTitle))ScatterText(l,"SCATTER_X_TITLE",null,FitPieText(CleanPieText(s.XTitle),p.W,10),p.X,height-26,p.W,18,10,Ink,"center",true);
            if(!String.IsNullOrEmpty(s.YTitle))ScatterText(l,"SCATTER_Y_TITLE",null,FitPieText(CleanPieText(s.YTitle),width-24,10),12,p.Y-21,width-24,18,10,Ink,"left",true);
        }

        private static void AddScatterMarker(Layout l,string marker,double cx,double cy,double diameter,int color,string role,string id)
        {
            if(marker=="None"||diameter<=0)return;
            Primitive p=new Primitive {Kind=marker=="Circle"?"ellipse":marker=="Square"?"rect":"path",X=cx-diameter/2,Y=cy-diameter/2,W=diameter,H=diameter,Color=color,Role=role,EntityId=id,LineWidth=0};
            if(p.Kind=="path")
            {
                double[,] vertices;
                if(marker=="Triangle") vertices=new double[,]{{0,-.5},{.5,.5},{-.5,.5}};
                else if(marker=="Diamond")vertices=new double[,]{{0,-.5},{.5,0},{0,.5},{-.5,0}};
                else vertices=new double[,]{{-.16,-.5},{.16,-.5},{.16,-.16},{.5,-.16},{.5,.16},{.16,.16},{.16,.5},{-.16,.5},{-.16,.16},{-.5,.16},{-.5,-.16},{-.16,-.16}};
                ShapePath path=new ShapePath();
                for(int i=0;i<vertices.GetLength(0);i++)
                {
                    double vx=vertices[i,0],vy=vertices[i,1];
                    if(marker=="Cross"){double temp=vx;vx=(vx-vy)/Math.Sqrt(2);vy=(temp+vy)/Math.Sqrt(2);}
                    double px=cx+vx*diameter,py=cy+vy*diameter;
                    if(i==0){path.StartX=px;path.StartY=py;}else path.Segments.Add(new PathSegment {X=px,Y=py});
                }
                path.Segments.Add(new PathSegment {X=path.StartX,Y=path.StartY});p.Path=path;
            }
            l.Items.Add(p);
        }

        // Clip an infinite line, not merely the segment between the supplied points.
        // Scaling all coordinates before subtracting avoids overflow for opposite extremes.
        private static bool ClipScatterLine(double x1,double y1,double x2,double y2,out double[] result)
        {
            result=null;
            if(!Finite(x1)||!Finite(y1)||!Finite(x2)||!Finite(y2))throw new ArgumentException("Line coordinates exceed numeric resolution for these axis bounds.");
            double scale=Math.Max(1,Math.Max(Math.Max(Math.Abs(x1),Math.Abs(x2)),Math.Max(Math.Abs(y1),Math.Abs(y2))));
            x1/=scale;y1/=scale;x2/=scale;y2/=scale;
            return ClipScatterScaled(x1,y1,x2-x1,y2-y1,scale,out result);
        }
        private static bool ClipScatterScaled(double x1,double y1,double dx,double dy,double scale,out double[] result)
        {
            result=null;
            double limit=1/scale,tolerance=1e-9/scale;
            if(dx==0&&dy==0)return false;
            List<double[]> hits=new List<double[]>();
            if(dx!=0)for(int i=0;i<2;i++){double bx=i*limit,t=(bx-x1)/dx,by=y1+t*dy;if(Finite(by)&&by>=-tolerance&&by<=limit+tolerance)hits.Add(new double[]{i,Math.Max(0,Math.Min(1,by*scale))});}
            if(dy!=0)for(int i=0;i<2;i++){double by=i*limit,t=(by-y1)/dy,bx=x1+t*dx;if(Finite(bx)&&bx>=-tolerance&&bx<=limit+tolerance)hits.Add(new double[]{Math.Max(0,Math.Min(1,bx*scale)),i});}
            if(hits.Count<2)return false;
            double longest=0;double[] a=null,b=null;
            for(int i=0;i<hits.Count;i++)for(int j=i+1;j<hits.Count;j++){double xx=hits[i][0]-hits[j][0],yy=hits[i][1]-hits[j][1],distance=xx*xx+yy*yy;if(distance>longest){longest=distance;a=hits[i];b=hits[j];}}
            if(longest==0)return false;result=new double[]{a[0],a[1],b[0],b[1]};return true;
        }
        private static void ScatterClippedLine(Layout l,ScatterBox p,double[] c,int color,string role,string id)
        {ScatterLine(l,p.X+c[0]*p.W,p.Y+(1-c[1])*p.H,(c[2]-c[0])*p.W,(c[1]-c[3])*p.H,color,role,id);}
        private static string ScatterTrendText(ScatterFit fit,ScatterTrendline definition)
        {
            string text="";
            if(definition.ShowEquation)
            {
                double slope=fit.RawSlope;
                double intercept=fit.RawIntercept;
                if(Finite(slope)&&Finite(intercept)&&(fit.Slope==0||slope!=0)&&((fit.MeanY-fit.Slope*fit.MeanX)==0||intercept!=0))text="y = "+ScatterNumber(slope,5)+" x + "+ScatterNumber(intercept,5);
                else text="y / "+ScatterNumber(fit.YUnit,5)+" = "+ScatterNumber(fit.Slope,5)+" (x / "+ScatterNumber(fit.XUnit,5)+" - "+ScatterNumber(fit.MeanX,5)+") + "+ScatterNumber(fit.MeanY,5);
            }
            if(definition.ShowRSquared)text+=(text.Length==0?"":"; ")+"R\u00b2 = "+fit.RSquared.ToString("G4",Invariant);
            if(text.Length>0&&definition.Group!=null)text=FitPieText(CleanPieText(definition.Group),120,7)+": "+text;
            return text;
        }

        private static string ScatterLabelText(ScatterPoint p,string content)
        {
            if(content=="Label")return CleanPieText(p.Label);
            if(content=="X")return ScatterNumber(p.X.Value,5);
            if(content=="Y")return ScatterNumber(p.Y.Value,5);
            if(content=="XY")return "("+ScatterNumber(p.X.Value,5)+", "+ScatterNumber(p.Y.Value,5)+")";
            if(content=="Size"&&p.Size.HasValue)return ScatterNumber(p.Size.Value,5);
            return "";
        }
        private static string ScatterNumber(double value,int digits)
        {
            // Rounded MaxValue at G4/G5 exceeds the representable range. Increase
            // precision until the displayed finite value can itself be parsed.
            for(int precision=digits;precision<=17;precision++)
            {
                string text=value.ToString("G"+precision,Invariant);double parsed;
                if(Double.TryParse(text,System.Globalization.NumberStyles.Float,Invariant,out parsed)&&Finite(parsed)&&(value==0||parsed!=0))return text;
            }
            return value.ToString("R",Invariant);
        }
        private static double Overlap(ScatterBox a,ScatterBox b)
        {return Math.Max(0,Math.Min(a.X+a.W,b.X+b.W)-Math.Max(a.X,b.X))*Math.Max(0,Math.Min(a.Y+a.H,b.Y+b.H)-Math.Max(a.Y,b.Y));}
        private static void ScatterPointLabels(Layout l,List<PlottedPoint> points,ScatterSpec s,ScatterBox plot)
        {
            List<ScatterBox> occupied=new List<ScatterBox>();
            // Order by input index so label priority does not change with bubble z-order.
            List<PlottedPoint> order=new List<PlottedPoint>(points);order.Sort(delegate(PlottedPoint a,PlottedPoint b){return a.Index.CompareTo(b.Index);});
            foreach(PlottedPoint point in order)for(int labelIndex=0;labelIndex<2;labelIndex++)
            {
                string content=labelIndex==0?s.LabelContent:s.SecondLabelContent,text=ScatterLabelText(point.Data,content);
                if(text.Length==0)continue;
                double font=s.LabelFontSize,maxWidth=Math.Min(180,plot.W),h=font*1.4;
                if(content=="Label")text=FitPieText(text,maxWidth,font);
                else{font=Math.Min(font,maxWidth/(text.Length*1.2));if(font<6)throw new ArgumentException("Numeric scatter labels do not fit the chart frame at 6 points. Increase its width or change label content.");h=font*1.4;}
                if(h>plot.H)throw new ArgumentException("Scatter label height exceeds the plot frame. Increase chart height or reduce label font size.");
                double w=Math.Min(maxWidth,text.Length*font*1.2),r=point.Diameter/2+3;
                double[,] candidates={{point.X+r,point.Y-h/2},{point.X-w-r,point.Y-h/2},{point.X-w/2,point.Y-h-r},{point.X-w/2,point.Y+r},{point.X+r,point.Y-h-r},{point.X-w-r,point.Y-h-r},{point.X+r,point.Y+r},{point.X-w-r,point.Y+r}};
                if(point.Marker=="None"){candidates[0,0]=point.X-w/2;candidates[0,1]=point.Y-h/2;}
                ScatterBox best=null;double bestScore=Double.PositiveInfinity;int candidateIndex=0;
                for(int c=0;c<8;c++)
                {
                    ScatterBox box=new ScatterBox(Math.Max(plot.X,Math.Min(plot.X+plot.W-w,candidates[c,0])),Math.Max(plot.Y,Math.Min(plot.Y+plot.H-h,candidates[c,1])),w,h);
                    double score=0;
                    foreach(ScatterBox used in occupied)score+=Overlap(box,used)*10;
                    foreach(PlottedPoint p in points)if(p.Diameter>0)score+=Overlap(box,new ScatterBox(p.X-p.Diameter/2,p.Y-p.Diameter/2,p.Diameter,p.Diameter));
                    if(score<bestScore){bestScore=score;best=box;candidateIndex=c;}
                }
                occupied.Add(best);string id=point.Index.ToString(Invariant);
                double nearX=Math.Max(best.X,Math.Min(best.X+best.W,point.X)),nearY=Math.Max(best.Y,Math.Min(best.Y+best.H,point.Y));
                double distance=Math.Sqrt((nearX-point.X)*(nearX-point.X)+(nearY-point.Y)*(nearY-point.Y));
                if(distance>point.Diameter/2+5 || candidateIndex!=0&&distance>1)ScatterLine(l,point.X,point.Y,nearX-point.X,nearY-point.Y,Gray,"SCATTER_LABEL_LEADER",id);
                ScatterText(l,labelIndex==0?"SCATTER_LABEL":"SCATTER_SECOND_LABEL",id,text,best.X,best.Y,best.W,best.H,font,Ink,"center",true);
            }
        }
    }
}
