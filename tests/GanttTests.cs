using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using OpenSlideCharts;

internal static class GanttTests
{
    private static int checks;
    private static void Check(bool value,string reason){if(!value)throw new Exception(reason);checks++;}
    private static void Near(double a,double b,string reason){Check(Math.Abs(a-b)<1e-7,reason+": "+a+" / "+b);}
    private static bool Finite(double x){return !Double.IsNaN(x)&&!Double.IsInfinity(x);}
    private static ChartSpec Spec(string start="2024-02-26",string end="2024-03-03")
    {
        GanttSpec g=new GanttSpec {StartDate=start,EndDate=end,WeekendDisplay="Plain",ScaleMode="Manual"};
        g.Scales.Add(new GanttScale {Unit="Days"});g.Activities.Add(new GanttActivity {Id="A",Label="Activity"});
        return new ChartSpec {Kind="Gantt",Title="Dated Gantt",Gantt=g};
    }
    private static GanttItem Add(ChartSpec s,string id,string start,string end,string kind="Bar",string row="A")
    {GanttItem item=new GanttItem {Id=id,RowId=row,StartDate=start,EndDate=end,Kind=kind};s.Gantt.Items.Add(item);return item;}
    private static Layout Build(ChartSpec s,double w=720,double h=400){return ChartEngine.Build(s,w,h);}
    private static List<Primitive> Role(Layout l,string role){return l.Items.FindAll(delegate(Primitive p){return p.Role==role;});}
    private static List<Primitive> Item(Layout l,string role,string id){return l.Items.FindAll(delegate(Primitive p){return p.Role==role&&p.EntityId==id;});}
    private static Primitive One(Layout l,string role,string id){List<Primitive> p=Item(l,role,id);Check(p.Count==1,"One "+role+" "+id);return p[0];}
    private static void Coordinate(double x,double y,double w,double h)
    {Check(Finite(x)&&Finite(y)&&x>=-1e-7&&y>=-1e-7&&x<=w+1e-7&&y<=h+1e-7,"Finite geometry inside frame");}
    private static bool Clean(string text)
    {
        for(int i=0;i<text.Length;i++)
        {
            char c=text[i];if(Char.IsControl(c)||c=='\u2028'||c=='\u2029')return false;
            if(Char.IsHighSurrogate(c)){if(i+1>=text.Length||!Char.IsLowSurrogate(text[++i]))return false;}
            else if(Char.IsLowSurrogate(c))return false;
        }
        return true;
    }
    private static void Frame(Layout l,double w=720,double h=400)
    {
        foreach(Primitive p in l.Items)
        {
            Coordinate(p.X,p.Y,w,h);Coordinate(p.X+p.W,p.Y+p.H,w,h);
            if(p.Path!=null)
            {
                Coordinate(p.Path.StartX,p.Path.StartY,w,h);
                foreach(PathSegment q in p.Path.Segments){Coordinate(q.X,q.Y,w,h);if(q.Curve){Coordinate(q.C1X,q.C1Y,w,h);Coordinate(q.C2X,q.C2Y,w,h);}}
                PathSegment last=p.Path.Segments[p.Path.Segments.Count-1];Near(last.X,p.Path.StartX,"Closed path X");Near(last.Y,p.Path.StartY,"Closed path Y");
            }
            if(p.Role!=null&&p.Kind=="text")
            {
                if(p.Role=="GANTT_ITEM_LABEL"&&!p.NoWrap)
                {
                    string[] lines=p.Text.Split('\n');
                    foreach(string line in lines){Check(Clean(line),"Sanitized internal process label");Check(line.Length*p.FontSize*1.2<=p.W+1e-7,"Complete process line fits conservative width");}
                    Check(lines.Length*p.FontSize*1.4<=p.H+1e-7,"Process text fits allocated height");
                }
                else Check(Clean(p.Text),"Sanitized Gantt text");
                if(p.NoWrap)Check(p.Text.Length*p.FontSize*1.2<=p.W+1e-7,"Conservative Gantt text budget");
            }
        }
    }
    private static void Reject(ChartSpec s,string reason)
    {
        bool rejected=false;try{Build(s);}catch(ArgumentException){rejected=true;}catch(FormatException){rejected=true;}
        Check(rejected,reason);
    }
    private static void Validate(ChartSpec s)
    {typeof(ChartEngine).GetMethod("ValidateGantt",BindingFlags.NonPublic|BindingFlags.Static).Invoke(null,new object[]{s});}
    public static int Main()
    {
        try
        {
            ChartSpec s=Spec();Add(s,"leap","2024-02-29","2024-02-29");Layout l=Build(s);Frame(l);
            Primitive edge=Role(l,"GANTT_SCALE_EDGE")[0],bar=One(l,"GANTT_ITEM","leap");
            Near(bar.W,edge.W/7,"Leap day inclusive one-day width");Near(bar.X,edge.X+edge.W*3/7,"Leap day correct position");
            Check(bar.ColumnId=="A"&&bar.Kind=="rect","Source row and native bar primitive");
            Check(One(l,"GANTT_ACTIVITY","__milestones").Y>One(l,"GANTT_ACTIVITY","A").Y,"Permanent milestone row at bottom");
            Check(Role(l,"GANTT_SCALE_SEPARATOR").Count==6,"No left/right outer calendar borders");
            Check(Role(l,"GANTT_SCALE_LABEL")[0].Alignment=="left","Calendar labels aligned left");
            s.Gantt.Items.Clear();s.Gantt.Activities.Clear();l=Build(s,320,240);Frame(l,320,240);
            Check(Role(l,"GANTT_ACTIVITY").Count==1,"Empty ordinary activity list preserves milestone row");

            s=Spec();Add(s,"first","2024-02-26","2024-02-26");Add(s,"last","2024-03-03","2024-03-03");
            l=Build(s);Near(One(l,"GANTT_ITEM","first").Y,One(l,"GANTT_ITEM","last").Y,"Disjoint items share one activity lane");
            Add(s,"overlap","2024-02-26","2024-02-27");l=Build(s);Frame(l);
            Check(One(l,"GANTT_ITEM","first").Y!=One(l,"GANTT_ITEM","overlap").Y,"Overlapping items get independent lanes");
            Check(s.Gantt.Items[0].Id=="first"&&s.Gantt.Items[1].Id=="last","Lane sort never mutates item order");

            s=Spec();Add(s,"milestone","2024-02-29",null,"Milestone");l=Build(s);edge=Role(l,"GANTT_SCALE_EDGE")[0];
            Primitive milestone=One(l,"GANTT_ITEM","milestone");Near(milestone.X+milestone.W/2,edge.X+edge.W*3.5/7,"Day scale milestone at day center");
            s.Gantt.Scales[0].Unit="Months";l=Build(s);edge=Role(l,"GANTT_SCALE_EDGE")[0];milestone=One(l,"GANTT_ITEM","milestone");
            Near(milestone.X+milestone.W/2,edge.X+edge.W*3/7,"Coarser scale milestone at day boundary");
            Add(s,"bracket","2024-02-27","2024-03-01","Bracket","__milestones");l=Build(s);Frame(l);
            Check(Item(l,"GANTT_ITEM","bracket").Count==3,"Bottom milestone row accepts bracket");

            s=Spec("2024-02-29","2024-03-01");GanttItem exact=Add(s,"exact","2024-02-29T06:00:00","2024-03-01T12:00:00");exact.DateMode="Exact";exact.ShowDates=true;
            GanttItem exactMilestone=Add(s,"exactM","2024-02-29T18:00:00",null,"Milestone");exactMilestone.DateMode="Exact";
            l=Build(s);Frame(l);edge=Role(l,"GANTT_SCALE_EDGE")[0];bar=One(l,"GANTT_ITEM","exact");
            Near(bar.X,edge.X+edge.W*.125,"Exact start time preserved");Near(bar.W,edge.W*.625,"Exact end is exclusive");
            milestone=One(l,"GANTT_ITEM","exactM");Near(milestone.X+milestone.W/2,edge.X+edge.W*.375,"Exact milestone preserves time with Days scale");
            Check(One(l,"GANTT_ITEM_DATES","exact").Text.Contains("06:00:00")&&One(l,"GANTT_ITEM_DATES","exact").Text.Contains("12:00:00"),"Exact date labels preserve both full timestamps");
            s.Gantt.DateAnchors.Add(new GanttDateAnchor {Id="start",RowId="A",Index=0,Date="2024-02-29T12:00:00"});
            s.Gantt.DateAnchors.Add(new GanttDateAnchor {Id="end",RowId="A",Index=1,Date="2024-03-01T00:00:00"});
            exact.StartAnchor="start";exact.EndAnchor="end";l=Build(s);bar=One(l,"GANTT_ITEM","exact");
            Near(bar.X,edge.X+edge.W*.25,"Anchor overrides fallback date");Near(bar.W,edge.W*.25,"End anchor retains exact exclusivity");
            Check(exact.StartDate=="2024-02-29T06:00:00"&&s.Gantt.DateAnchors[0].Date=="2024-02-29T12:00:00","Rendering never writes anchored dates back to DTO");

            s=Spec();s.Gantt.WeekendDisplay="Shaded";l=Build(s);Frame(l);Check(Role(l,"GANTT_WEEKEND").Count==2,"Saturday/Sunday shade bands");
            s.Gantt.WeekendDisplay="Hidden";Add(s,"workday","2024-03-01","2024-03-01");GanttItem weekend=Add(s,"weekend","2024-03-02","2024-03-03");weekend.Label="Weekend retained";
            Add(s,"weekendM","2024-03-02",null,"Milestone").Label="Hidden milestone";
            l=Build(s);Frame(l);edge=Role(l,"GANTT_SCALE_EDGE")[0];
            Near(One(l,"GANTT_ITEM","workday").W,edge.W/5,"Hidden weekends compress calendar to weekdays");
            Check(Item(l,"GANTT_ITEM","weekend").Count==0&&Item(l,"GANTT_ITEM","weekendM").Count==0,"Weekend-only geometry omitted");
            Check(Item(l,"GANTT_ITEM_LABEL","weekend").Count==1&&Item(l,"GANTT_ITEM_LABEL","weekendM").Count==1,"Weekend-only labels retained inside frame");
            s.Gantt.Shades.Add(new GanttShade {StartDate="2024-02-29",EndDate="2024-03-01",Label="Holiday",Color="#F0E0D0"});
            l=Build(s);Frame(l);Near(Role(l,"GANTT_SHADE")[0].W,edge.W*2/5,"Custom holidays shade without further compression");
            s=Spec("2024-03-02","2024-03-03");s.Gantt.WeekendDisplay="Hidden";Reject(s,"All-hidden chart range rejected clearly");

            s=Spec();GanttItem clipped=Add(s,"clip","2024-02-01","2024-03-31");clipped.Label="Clipped";
            GanttItem outside=Add(s,"outside","2025-01-01","2025-01-10");outside.Label="Restore me";
            l=Build(s);Frame(l);edge=Role(l,"GANTT_SCALE_EDGE")[0];bar=One(l,"GANTT_ITEM","clip");
            Near(bar.X,edge.X,"Clipped item begins at chart boundary");Near(bar.W,edge.W,"Clipped item spans full range");
            Check(bar.Kind=="path"&&bar.Path.StartX>bar.X,"Clipped bar contour has visible arrow cap, not rectangle overpaint");
            Check(Item(l,"GANTT_OVERFLOW_START","clip").Count==1&&Item(l,"GANTT_OVERFLOW_END","clip").Count==1,"Both overflow caps tagged");
            Check(Item(l,"GANTT_ITEM","outside").Count==0&&Item(l,"GANTT_ITEM_LABEL","outside").Count==0,"Out-of-range item and label omitted");
            Check(outside.Label=="Restore me"&&outside.StartDate=="2025-01-01","Hidden item metadata retained");
            s.Gantt.StartDate="2025-01-01";s.Gantt.EndDate="2025-01-14";l=Build(s);Check(One(l,"GANTT_ITEM_LABEL","outside").Text=="Restore me","Range restoration restores original item label");

            s=Spec();s.Gantt.Activities[0].Responsible="Alice\nBob";s.Gantt.Activities[0].Remark="\t"+new String('W',60)+"\u754c\ud83d\ude00";s.Gantt.Activities[0].Label=new String('W',70)+"\t\u754c\ud83d\ude00";
            s.Gantt.Activities[0].Indent=10;s.Gantt.Activities[0].Color="#EEEEFF";s.Gantt.Activities[0].Separator=true;s.Gantt.ShowResponsible=true;s.Gantt.ShowRemark=true;
            Add(s,"long","2024-02-27","2024-03-01","Chevron").Label="Go";
            l=Build(s,320,240);Frame(l,320,240);
            Check(Role(l,"GANTT_RESPONSIBLE").Count==1&&Role(l,"GANTT_REMARK").Count==1,"Optional row columns rendered independently");
            Check(Role(l,"GANTT_ROW_SHADE").Count==1&&Role(l,"GANTT_ROW_SEPARATOR").Count==1,"Row fill and separator preserved");
            Near(One(l,"GANTT_ROW_SEPARATOR","A").Y,One(l,"GANTT_ROW_SHADE","A").Y,"Activity separator is above its row");
            Check(s.Gantt.Activities[0].Responsible=="Alice\nBob","Text sanitation does not change stored responsible names");

            s=Spec("2024-02-26","2024-03-04");
            string longProcessLabel=new String('W',80)+"\u754c\ud83d\ude00";
            GanttItem longProcess=Add(s,"process","2024-02-27","2024-03-01","Chevron");longProcess.Label=longProcessLabel;longProcess.ShowDates=true;longProcess.Color="#000000";
            l=Build(s,720,450);Frame(l,720,450);Primitive processShape=One(l,"GANTT_ITEM","process"),processLabel=One(l,"GANTT_ITEM_LABEL","process");
            Check(processShape.H>30,"Process arrow grows vertically for long content");
            Check(processLabel.X>processShape.X&&processLabel.Y>processShape.Y&&processLabel.X+processLabel.W<processShape.X+processShape.W&&processLabel.Y+processLabel.H<processShape.Y+processShape.H,"Process text lies strictly within shape bounding box");
            double processCap=Math.Min(processShape.H/2,processShape.W/3);
            Check(processLabel.X>=processShape.X+processCap&&processLabel.X+processLabel.W<=processShape.X+processShape.W-processCap,"Chevron internal text excludes both arrow caps");
            Check(processLabel.Text.Replace("\n","")==longProcessLabel&&longProcess.Label==longProcessLabel,"Long process text fully retained with surrogate-safe wrapping, DTO unchanged");
            Check(processLabel.Color==0xFFFFFF,"Dark process fill uses white internal text");
            Check(One(l,"GANTT_ITEM_DATES","process").Y>processShape.Y+processShape.H,"Process dates remain below expanded arrow");
            longProcess.Kind="Pentagon";longProcess.Label="Développement";longProcess.Color="#FFFFFF";
            l=Build(s,720,450);Frame(l,720,450);processShape=One(l,"GANTT_ITEM","process");processLabel=One(l,"GANTT_ITEM_LABEL","process");
            Check(processLabel.Color==0&&processLabel.Y>=processShape.Y&&processLabel.Y+processLabel.H<=processShape.Y+processShape.H,"Pentagon contains text with contrasting black on white");
            Check(processLabel.Text.Replace("\n","")=="Développement","Normal process label remains complete");
            longProcess.StartDate="2024-02-01";longProcess.EndDate="2024-03-31";l=Build(s,720,450);Frame(l,720,450);
            processShape=One(l,"GANTT_ITEM","process");processLabel=One(l,"GANTT_ITEM_LABEL","process");processCap=Math.Min(processShape.H/2,processShape.W/3);
            Check(processLabel.X>=processShape.X+processCap&&processLabel.X+processLabel.W<=processShape.X+processShape.W-processCap,"Clipped process text avoids overflow caps");
            longProcess.Label=new String('W',3000);Reject(s,"Process label exceeding frame height fails explicitly");
            s=Spec();GanttItem narrowProcess=Add(s,"narrow","2024-02-26T00:00:00","2024-02-26T01:00:00","Chevron");narrowProcess.DateMode="Exact";narrowProcess.Label="Narrow";
            Reject(s,"Process text cannot escape a narrow arrow into adjacent dates");
            s=Spec("2024-02-26","2024-03-04");GanttItem combining=Add(s,"combining","2024-02-26","2024-02-26","Chevron");combining.Label="e\u0301e\u0301e\u0301";
            Reject(s,"Narrow process refuses a whole combining element instead of splitting its accent");
            combining.EndDate="2024-02-27";combining.Label="e\u0301i\u0308o\u0300e\u0301i\u0308o\u0300";
            l=Build(s);Frame(l);processLabel=One(l,"GANTT_ITEM_LABEL","combining");
            Check(processLabel.Text.Replace("\n","")==combining.Label,"Decomposed accents preserved without normalization or mutation");
            foreach(string line in processLabel.Text.Split('\n'))
            {
                UnicodeCategory category=CharUnicodeInfo.GetUnicodeCategory(line,0);
                Check(category!=UnicodeCategory.NonSpacingMark&&category!=UnicodeCategory.SpacingCombiningMark&&category!=UnicodeCategory.EnclosingMark,"Wrapped line never starts with detached combining mark");
            }
            string joinedEmoji="\ud83d\udc69\u200D\ud83d\udcbb";combining.Label=joinedEmoji+joinedEmoji;
            l=Build(s);Frame(l);processLabel=One(l,"GANTT_ITEM_LABEL","combining");
            Check(processLabel.Text.Replace("\n","")==combining.Label,"ZWJ emoji content retained exactly");
            foreach(string line in processLabel.Text.Split('\n'))Check(line==joinedEmoji,"Explicit ZWJ chain stays on one line despite legacy StringInfo boundaries");
            s=Spec();GanttItem unchangedBar=Add(s,"barText","2024-02-27","2024-03-01");unchangedBar.Label="Below bar";
            l=Build(s);bar=One(l,"GANTT_ITEM","barText");Check(One(l,"GANTT_ITEM_LABEL","barText").Y>bar.Y+bar.H,"Ordinary bar label remains below bar");

            s=Spec();s.Gantt.Activities.Add(new GanttActivity {Id="B",Label="B"});
            Add(s,"pentagon","2024-02-26","2024-02-27","Pentagon");Add(s,"chevron","2024-02-29","2024-03-01","Chevron","B");
            s.Gantt.Links.Add(new GanttLink {FromId="pentagon",ToId="chevron",FromEdge="End",ToEdge="Start"});
            l=Build(s);Frame(l);Check(One(l,"GANTT_ITEM","pentagon").Path!=null&&One(l,"GANTT_ITEM","chevron").Path!=null,"Process arrows use native closed contours");
            Check(Role(l,"GANTT_LINK").Count==3&&Role(l,"GANTT_LINK_ARROW").Count==1,"Visual link routed with native elbow lines and arrow");
            Check(s.Gantt.Items[1].StartDate=="2024-02-29","Visual link performs no scheduling");
            s.Gantt.Items[1].StartDate="2025-01-01";s.Gantt.Items[1].EndDate="2025-01-02";l=Build(s);Check(Role(l,"GANTT_LINK").Count==0,"Link to hidden geometry omitted without changing model");

            foreach(string boundary in new[]{"0001-01-01","9999-12-31"})
            {
                s=Spec(boundary,boundary);Add(s,"boundary",boundary,boundary);l=Build(s);Frame(l);
                Near(One(l,"GANTT_ITEM","boundary").W,Role(l,"GANTT_SCALE_EDGE")[0].W,"Inclusive sentinel date renders full single day");
            }
            s=Spec("2023-12-15","2024-04-15");s.Gantt.Scales.Clear();s.Gantt.Scales.Add(new GanttScale {Unit="Years",Format="yyyy"});s.Gantt.Scales.Add(new GanttScale {Unit="Quarters",Format="'Q'q"});s.Gantt.Scales.Add(new GanttScale {Unit="Months",Format="mmm",Separators=false});
            l=Build(s);Frame(l);Check(Role(l,"GANTT_SCALE_EDGE").Count==3,"Multiple calendar scales cover year and quarter rollover");
            foreach(Primitive p in Role(l,"GANTT_SCALE_SEPARATOR"))Check(!p.EntityId.StartsWith("Months:"),"Scale separator option respected");
            foreach(Primitive p in Role(l,"GANTT_SCALE_SEPARATOR"))Check(p.DashStyle==(p.EntityId.StartsWith("Years:")?0:4),"Coarsest calendar separators solid; finer separators dashed");
            s.Gantt.Scales.Reverse();l=Build(s);
            Check(Role(l,"GANTT_SCALE_EDGE")[0].EntityId=="Months","Manual scale input order preserved");
            foreach(Primitive p in Role(l,"GANTT_SCALE_SEPARATOR"))Check(p.DashStyle==(p.EntityId.StartsWith("Years:")?0:4),"Dash granularity independent of scale input order");
            s.Gantt.ScaleMode="Auto";l=Build(s);Frame(l);Check(Role(l,"GANTT_SCALE_EDGE").Count>0,"Automatic scale selection works");
            foreach(Primitive p in Role(l,"GANTT_SCALE_LABEL"))Check(!p.Text.Contains("\u2026"),"Calendar dates never abbreviated into ellipses");
            s=Spec("2024-01-01","2024-12-31");s.Gantt.ScaleMode="Auto";l=Build(s,320,240);Frame(l,320,240);
            Check(!Role(l,"GANTT_SCALE_EDGE").Exists(delegate(Primitive p){return p.EntityId=="Days"||p.EntityId=="Weeks";}),"Small frame automatically chooses coarser complete date labels");
            s=Spec();s.Gantt.Scales[0].Format="'"+new String('W',50)+"'";Reject(s,"Manual full date labels cannot be silently abbreviated");
            s=Spec("2024-01-01","2024-12-31");Reject(s,"Manual daily scale too dense is explicit error");

            s=Spec();for(int i=1;i<100;i++)s.Gantt.Activities.Add(new GanttActivity {Id="R"+i,Label="Activity "+i});
            Validate(s);Check(true,"100 rows data validation independent of frame");Reject(s,"Too many rows for current frame rejected");
            s.Gantt.Activities.Add(new GanttActivity {Id="R100"});Reject(s,"101 ordinary rows rejected");
            s=Spec();for(int i=0;i<1001;i++)Add(s,"I"+i,"2024-02-26","2024-02-26");Reject(s,"1001 items rejected");
            s=Spec();Add(s,"duplicate","2024-02-26","2024-02-26");Add(s,"duplicate","2024-02-27","2024-02-27");Reject(s,"Duplicate item IDs rejected");
            s=Spec();s.Gantt.Activities.Add(new GanttActivity {Id="__milestones"});Reject(s,"Milestone row ID collision rejected");
            s=Spec();Add(s,"bad","2024-02-30","2024-03-01");Reject(s,"Invalid Gregorian date rejected");
            s=Spec();Add(s,"bad","2024-03-01","2024-02-29");Reject(s,"Reversed interval rejected");
            s=Spec();Add(s,"bad","2024-02-26","2024-02-26").DateMode="Exact";Reject(s,"Zero-duration Exact interval rejected");
            s=Spec();Add(s,"bad","2024-02-26","2024-02-27","Milestone");Reject(s,"Milestone cannot have two dates");
            s=Spec();Add(s,"bad","2024-02-26","2024-02-27","Bar","__milestones");Reject(s,"Bottom milestone row cannot contain bar");
            s=Spec();Add(s,"bad","2024-02-26","2024-02-27","Star");Reject(s,"Unknown item kind rejected");
            s=Spec();Add(s,"bad","2024-02-26","2024-02-27").Color="#BADHEX";Reject(s,"Invalid color rejected");
            s=Spec();Add(s,"bad","2024-02-26","2024-02-27","Bar","missing");Reject(s,"Missing activity rejected");
            s=Spec();s.Gantt.DateAnchors.Add(new GanttDateAnchor {Id="anchor",RowId="A",Index=0,Date="2024-02-29"});Add(s,"bad","2024-02-26","2024-03-01").StartAnchor="anchor";Reject(s,"Anchored endpoints require Exact mode");
            s.Gantt.Items[0].DateMode="Exact";s.Gantt.DateAnchors[0].RowId="__milestones";Reject(s,"Anchors cannot bind across rows");
            s=Spec();s.Gantt.DateAnchors.Add(new GanttDateAnchor {Id="a",RowId="A",Index=0,Date="2024-02-29"});s.Gantt.DateAnchors.Add(new GanttDateAnchor {Id="b",RowId="A",Index=0,Date="2024-03-01"});Reject(s,"Duplicate anchor slot rejected");
            s=Spec();s.Gantt.DateAnchors.Add(new GanttDateAnchor {Id="a",RowId="A",Index=32,Date="2024-02-29"});Reject(s,"Anchor index32 rejected");
            s=Spec();s.Gantt.Links.Add(new GanttLink {FromId="absent",ToId="alsoAbsent"});Reject(s,"Missing link endpoint rejected");
            s=Spec();s.Gantt.Scales[0].Unit="Hours";Reject(s,"Unsupported calendar scale rejected");
            s=Spec();s.Gantt.Scales[0].Format="unsupported";Reject(s,"Unsupported date format rejected");
            s=Spec();s.Gantt.FontSize=Double.NaN;Reject(s,"Nonfinite font rejected");
            s=Spec();s.Gantt.LabelFraction=.8;Reject(s,"Insufficient timeline fraction rejected");
            s=Spec();s.Rows.Add(new ChartRow());Reject(s,"Legacy scalar and dated Gantt ambiguity rejected");
            s=Spec();s.Scatter=new ScatterSpec();Reject(s,"Scatter/Gantt ambiguity rejected");
            s=Spec();s.Kind="Column";Reject(s,"Dated Gantt options require correct kind");
            s=GanttSamples.Create();l=Build(s,720,480);Frame(l,720,480);
            Console.WriteLine("GanttTests: "+checks+" checks passed; original dated layout, no Office execution.");return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
    }
}
