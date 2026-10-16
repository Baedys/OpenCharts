using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenSlideCharts
{
    public static partial class ChartEngine
    {
        private static readonly string[] GanttUnits={"Years","Quarters","Months","Weeks","Days"};
        private static readonly string[] GanttKinds={"Bar","Pentagon","Chevron","Milestone","Bracket"};

        // Calendar dates and item relationships are validated without requiring a frame.
        // Lane allocation, scale choice and text abbreviation are original layout policies.
        internal static void ValidateGantt(ChartSpec spec)
        {
            if(spec==null)throw new ArgumentNullException("spec");
            if(!String.Equals(spec.Kind,"Gantt",StringComparison.OrdinalIgnoreCase)||spec.Gantt==null)
                throw new ArgumentException("Dated Gantt data requires Kind=Gantt and Gantt options.");
            if(spec.Rows!=null&&spec.Rows.Count!=0 || spec.Series!=null&&spec.Series.Count!=0 || spec.Appearance!=null || spec.Waterfall!=null || spec.Pie!=null || spec.Scatter!=null || spec.LayoutMode!=null&&spec.LayoutMode!="Clustered")
                throw new ArgumentException("Dated Gantt data cannot be combined with scalar rows, series, appearance or other chart models.");
            GanttSpec g=spec.Gantt;
            if(!Finite(g.LabelFraction)||g.LabelFraction<.15||g.LabelFraction>.65)throw new ArgumentException("Gantt LabelFraction must be between .15 and .65.");
            if(!Finite(g.FontSize)||g.FontSize<6||g.FontSize>24)throw new ArgumentException("Gantt FontSize must be between 6 and 24 points.");
            if(g.ScaleMode!="Auto"&&g.ScaleMode!="Manual")throw new ArgumentException("Gantt ScaleMode must be Auto or Manual.");
            if(g.Activities==null||g.Activities.Count>100)throw new ArgumentException("Provide at most 100 ordinary Gantt activities.");
            if(g.MilestoneRow==null)throw new ArgumentException("The Gantt milestone row is required.");
            if(g.Items==null||g.Items.Count>1000)throw new ArgumentException("Provide at most 1000 Gantt items.");
            if(g.Shades!=null&&g.Shades.Count>100 || g.Links!=null&&g.Links.Count>1000)throw new ArgumentException("Gantt supports at most 100 shades and 1000 visual links.");
            if(g.DateAnchors!=null&&g.DateAnchors.Count>3232)throw new ArgumentException("Gantt supports at most 3232 date anchors.");
            GanttCalendar calendar=new GanttCalendar(g);
            if(g.ScaleMode=="Manual"&&(g.Scales==null||g.Scales.Count==0))throw new ArgumentException("Manual Gantt scale mode requires at least one scale.");
            if(g.Scales!=null)
            {
                if(g.Scales.Count>5)throw new ArgumentException("At most five Gantt calendar scales are supported.");
                HashSet<string> units=new HashSet<string>(StringComparer.Ordinal);
                foreach(GanttScale scale in g.Scales)
                {
                    if(scale==null||Array.IndexOf(GanttUnits,scale.Unit)<0||!units.Add(scale.Unit))throw new ArgumentException("Gantt scales must have distinct supported calendar units.");
                    if(scale.Format!=null)calendar.Format(new DateTime(2024,2,29),scale.Format);
                }
            }
            Dictionary<string,GanttActivity> rows=GanttRows(g);
            Dictionary<string,GanttDateAnchor> anchors=new Dictionary<string,GanttDateAnchor>(StringComparer.Ordinal);
            HashSet<string> slots=new HashSet<string>(StringComparer.Ordinal);
            if(g.DateAnchors!=null)foreach(GanttDateAnchor anchor in g.DateAnchors)
            {
                if(anchor==null)throw new ArgumentException("Gantt date anchors cannot be null.");
                GanttId(anchor.Id,"anchor");
                if(anchors.ContainsKey(anchor.Id)||anchor.RowId==null||!rows.ContainsKey(anchor.RowId)||anchor.Index<0||anchor.Index>=32||!slots.Add(anchor.RowId+"\u0001"+anchor.Index.ToString(Invariant)))
                    throw new ArgumentException("Gantt anchors require unique IDs and indices 0 to 31 in an existing row.");
                GanttDates.ParseInstant(anchor.Date);anchors.Add(anchor.Id,anchor);
            }
            Dictionary<string,GanttItem> items=new Dictionary<string,GanttItem>(StringComparer.Ordinal);
            foreach(GanttItem item in g.Items)
            {
                if(item==null)throw new ArgumentException("Gantt items cannot be null.");
                GanttId(item.Id,"item");
                if(items.ContainsKey(item.Id)||item.RowId==null||!rows.ContainsKey(item.RowId))throw new ArgumentException("Gantt items require unique IDs and an existing RowId.");
                if(Array.IndexOf(GanttKinds,item.Kind)<0)throw new ArgumentException("Unsupported Gantt item kind.");
                if(item.RowId==g.MilestoneRow.Id&&item.Kind!="Milestone"&&item.Kind!="Bracket")throw new ArgumentException("Only milestones and brackets can occupy the fixed milestone row.");
                if(item.DateMode!="InclusiveDays"&&item.DateMode!="Exact")throw new ArgumentException("Gantt DateMode must be InclusiveDays or Exact.");
                string start=GanttEndpoint(item,false,anchors),end=GanttEndpoint(item,true,anchors);
                DateTime first=item.DateMode=="Exact"?GanttDates.ParseInstant(start):GanttDates.ParseDay(start);
                if(item.Kind=="Milestone")
                {
                    if(!String.IsNullOrEmpty(end))
                    {
                        DateTime last=item.DateMode=="Exact"?GanttDates.ParseInstant(end):GanttDates.ParseDay(end);
                        if(last!=first)throw new ArgumentException("A Gantt milestone has one date; an optional end must equal its start.");
                    }
                }
                else
                {
                    DateTime last=item.DateMode=="Exact"?GanttDates.ParseInstant(end):GanttDates.ParseDay(end);
                    if(last<first||item.DateMode=="Exact"&&last==first)throw new ArgumentException("Gantt ranges must be ordered; Exact intervals require a positive duration.");
                }
                if(item.Color!=null)PieColor(item.Color);
                items.Add(item.Id,item);
            }
            if(g.Shades!=null)foreach(GanttShade shade in g.Shades)
            {
                if(shade==null||GanttDates.ParseDay(shade.EndDate)<GanttDates.ParseDay(shade.StartDate))throw new ArgumentException("Gantt shades require an ordered inclusive date range.");
                if(shade.Color!=null)PieColor(shade.Color);
            }
            if(g.Links!=null)foreach(GanttLink link in g.Links)
            {
                if(link==null||link.FromId==null||link.ToId==null||!items.ContainsKey(link.FromId)||!items.ContainsKey(link.ToId)||link.FromId==link.ToId)
                    throw new ArgumentException("Gantt visual links require two distinct existing item IDs.");
                if(link.FromEdge!="Start"&&link.FromEdge!="End" || link.ToEdge!="Start"&&link.ToEdge!="End")throw new ArgumentException("Gantt link edges must be Start or End.");
                if(link.Color!=null)PieColor(link.Color);
            }
        }

        private static void GanttId(string value,string kind)
        {
            if(String.IsNullOrWhiteSpace(value)||value!=value.Trim())throw new ArgumentException("Gantt "+kind+" IDs must be nonempty and have no surrounding whitespace.");
            foreach(char c in value)if(Char.IsControl(c))throw new ArgumentException("Gantt IDs cannot contain control characters.");
        }
        private static Dictionary<string,GanttActivity> GanttRows(GanttSpec g)
        {
            Dictionary<string,GanttActivity> result=new Dictionary<string,GanttActivity>(StringComparer.Ordinal);
            List<GanttActivity> all=new List<GanttActivity>(g.Activities);all.Add(g.MilestoneRow);
            foreach(GanttActivity row in all)
            {
                if(row==null)throw new ArgumentException("Gantt rows cannot be null.");
                GanttId(row.Id,"row");
                if(result.ContainsKey(row.Id))throw new ArgumentException("Gantt row IDs must be unique, including the milestone row.");
                if(row.Indent<0||row.Indent>10)throw new ArgumentException("Gantt row indentation must be between 0 and 10.");
                if(row.Color!=null)PieColor(row.Color);
                result.Add(row.Id,row);
            }
            return result;
        }
        private static string GanttEndpoint(GanttItem item,bool end,Dictionary<string,GanttDateAnchor> anchors)
        {
            string id=end?item.EndAnchor:item.StartAnchor;
            if(id==null)return end?item.EndDate:item.StartDate;
            GanttDateAnchor anchor;
            if(item.DateMode!="Exact"||!anchors.TryGetValue(id,out anchor)||anchor.RowId!=item.RowId)
                throw new ArgumentException("Anchored Gantt endpoints require Exact mode and an anchor in the same row.");
            return anchor.Date;
        }

        private sealed class GanttDrawing
        {
            internal GanttItem Item;
            internal double X1,X2,LabelX,LabelW,DateH,Height,Y,ShapeHeight,InnerX,InnerW;
            internal bool LeftOverflow,RightOverflow,DrawShape,InternalLabel;
            internal int Color,Lane;
            internal string Label,Dates;
        }
        private sealed class GanttLane
        {
            internal double Right,Height,Offset;
        }
        private sealed class GanttRowDrawing
        {
            internal GanttActivity Row;
            internal List<GanttDrawing> Items=new List<GanttDrawing>();
            internal List<GanttLane> Lanes=new List<GanttLane>();
            internal double Height,Y;
        }
        private sealed class GanttShadeDrawing
        {
            internal GanttShade Shade;internal double X,W;internal string Label,Id;internal int Lane;
        }

        private static List<GanttScale> GanttChosenScales(GanttSpec g,GanttCalendar calendar,double plotWidth)
        {
            if(g.ScaleMode=="Manual")return new List<GanttScale>(g.Scales);
            // Try successively coarser calendars. Complete date labels must fit at
            // six points; unlike categories they are not shortened into ellipses.
            string[][] choices={new[]{"Months","Days"},new[]{"Months","Weeks"},new[]{"Years","Months"},new[]{"Years","Quarters"},new[]{"Years"}};
            List<GanttScale> result=null;
            foreach(string[] choice in choices)
            {
                result=new List<GanttScale>();bool fits=true;
                foreach(string unit in choice)
                {
                    GanttScale scale=new GanttScale {Unit=unit};result.Add(scale);
                    if(!GanttScaleFits(calendar,scale,plotWidth))fits=false;
                }
                if(fits)return result;
            }
            return result;
        }
        private static bool GanttScaleFits(GanttCalendar calendar,GanttScale scale,double width)
        {
            List<GanttPeriod> periods=calendar.Periods(scale.Unit);int fitted=0;
            foreach(GanttPeriod period in periods)
            {
                double w=(calendar.BoundaryFraction(period.EndDayExclusive)-calendar.BoundaryFraction(period.StartDay))*width;
                string text=CleanPieText(calendar.Format(period.LabelDate,scale.Format??GanttDefaultFormat(scale.Unit)));
                if(w>=text.Length*6*1.2+4)fitted++;
                else if(period.StartDay!=calendar.StartDay&&period.EndDayExclusive!=calendar.EndDayExclusive)return false;
            }
            return fitted>0;
        }
        private static string GanttDefaultFormat(string unit)
        {
            if(unit=="Years")return "yyyy";
            if(unit=="Quarters")return "'Q'q";
            if(unit=="Months")return "mmm";
            if(unit=="Weeks")return "ww";
            return "dd";
        }

        private static void BuildGantt(Layout layout,ChartSpec spec,double width,double height)
        {
            GanttSpec g=spec.Gantt;GanttCalendar calendar=new GanttCalendar(g);
            double left=14,right=width-14,tableWidth=(width-28)*g.LabelFraction,plotX=left+tableWidth+8,plotW=right-plotX;
            double font=g.FontSize,textH=font*1.4,shapeH=Math.Max(8,Math.Min(16,font*1.2)),scaleH=Math.Max(18,textH+4);
            List<GanttScale> scales=GanttChosenScales(g,calendar,plotW);
            bool centerMilestones=false;
            foreach(GanttScale scale in scales)if(scale.Unit=="Days")centerMilestones=true;
            double scaleTop=44,bodyTop=scaleTop+scales.Count*scaleH+8;
            List<GanttShadeDrawing> shades=GanttPrepareShades(g,calendar,plotX,plotW,font);
            int shadeLanes=0;foreach(GanttShadeDrawing shade in shades)if(shade.Label.Length!=0)shadeLanes=Math.Max(shadeLanes,shade.Lane+1);
            bodyTop+=shadeLanes*(textH+2);
            double bodyBottom=height-14;
            Dictionary<string,GanttDateAnchor> anchors=new Dictionary<string,GanttDateAnchor>(StringComparer.Ordinal);
            if(g.DateAnchors!=null)foreach(GanttDateAnchor anchor in g.DateAnchors)anchors.Add(anchor.Id,anchor);
            List<GanttRowDrawing> rows=new List<GanttRowDrawing>();Dictionary<string,GanttRowDrawing> byRow=new Dictionary<string,GanttRowDrawing>(StringComparer.Ordinal);
            foreach(GanttActivity row in g.Activities){GanttRowDrawing draw=new GanttRowDrawing {Row=row};rows.Add(draw);byRow.Add(row.Id,draw);}
            GanttRowDrawing milestoneRow=new GanttRowDrawing {Row=g.MilestoneRow};rows.Add(milestoneRow);byRow.Add(g.MilestoneRow.Id,milestoneRow);
            foreach(GanttItem item in g.Items)
            {
                GanttDrawing draw=GanttPrepareItem(item,anchors,calendar,centerMilestones,plotX,plotW,font,shapeH);
                if(draw!=null)byRow[item.RowId].Items.Add(draw);
            }
            double minimumHeight=0;
            foreach(GanttRowDrawing row in rows)
            {
                List<GanttDrawing> ordered=new List<GanttDrawing>(row.Items);
                ordered.Sort(delegate(GanttDrawing a,GanttDrawing b){int c=Math.Min(a.X1,a.LabelX).CompareTo(Math.Min(b.X1,b.LabelX));return c!=0?c:StringComparer.Ordinal.Compare(a.Item.Id,b.Item.Id);});
                foreach(GanttDrawing item in ordered)
                {
                    double markerRadius=item.Item.Kind=="Milestone"?item.ShapeHeight/2:0;
                    double start=Math.Min(item.X1-markerRadius,item.LabelX),end=Math.Max(item.X2+markerRadius,item.LabelX+item.LabelW);
                    int lane=0;while(lane<row.Lanes.Count&&row.Lanes[lane].Right+4>start)lane++;
                    if(lane==row.Lanes.Count)row.Lanes.Add(new GanttLane());
                    item.Lane=lane;row.Lanes[lane].Right=end;row.Lanes[lane].Height=Math.Max(row.Lanes[lane].Height,item.Height);
                }
                double stack=0;foreach(GanttLane lane in row.Lanes){lane.Offset=stack;stack+=lane.Height+4;}
                row.Height=Math.Max(textH+10,stack+8);minimumHeight+=row.Height;
            }
            if(plotW<60||bodyBottom-bodyTop<minimumHeight)throw new ArgumentException("Gantt chart frame is too small for its activity rows, overlapping items and labels. Increase the frame or reduce row/item density.");
            double extra=(bodyBottom-bodyTop-minimumHeight)/rows.Count,y=bodyTop;
            foreach(GanttRowDrawing row in rows){row.Y=y;row.Height+=extra;y+=row.Height;}
            // Backgrounds first: row colors, weekend bands, then custom shade ranges.
            foreach(GanttRowDrawing row in rows)if(row.Row.Color!=null)GanttRect(layout,"GANTT_ROW_SHADE",row.Row.Id,row.Row.Id,left,row.Y,right-left,row.Height,PieColor(row.Row.Color));
            if(g.WeekendDisplay=="Shaded")
            {
                for(int day=calendar.StartDay;day<calendar.EndDayExclusive;day++)
                {
                    DateTime date=new DateTime((long)day*TimeSpan.TicksPerDay);
                    if(g.WeekendDays.Contains((int)date.DayOfWeek))
                    {
                        double x=plotX+calendar.BoundaryFraction(day)*plotW,w=(calendar.BoundaryFraction(day+1)-calendar.BoundaryFraction(day))*plotW;
                        if(w>0)GanttRect(layout,"GANTT_WEEKEND",GanttDates.Iso(date),null,x,bodyTop,w,bodyBottom-bodyTop,0xEFEFEF);
                    }
                }
            }
            foreach(GanttShadeDrawing shade in shades)
            {
                GanttRect(layout,"GANTT_SHADE",shade.Id,null,shade.X,bodyTop,shade.W,bodyBottom-bodyTop,shade.Shade.Color==null?0xE4EEDD:PieColor(shade.Shade.Color));
                if(shade.Label.Length!=0)GanttText(layout,"GANTT_SHADE_LABEL",shade.Id,null,shade.Label,shade.X,bodyTop-(shadeLanes-shade.Lane)*(textH+2),shade.W,textH,font,Ink,"center",true);
            }
            GanttDrawScales(layout,calendar,scales,plotX,plotW,scaleTop,scaleH,bodyBottom,font);
            int columns=1+(g.ShowResponsible?1:0)+(g.ShowRemark?1:0);
            double columnW=tableWidth/columns;
            GanttTextFit(layout,"GANTT_HEADER","activity",null,g.ActivityHeader,left,bodyTop-textH-4,columnW-4,textH,font,Ink,"left");
            int col=1;
            if(g.ShowResponsible){GanttTextFit(layout,"GANTT_HEADER","responsible",null,g.ResponsibleHeader,left+col*columnW,bodyTop-textH-4,columnW-4,textH,font,Ink,"left");col++;}
            if(g.ShowRemark)GanttTextFit(layout,"GANTT_HEADER","remark",null,g.RemarkHeader,left+col*columnW,bodyTop-textH-4,columnW-4,textH,font,Ink,"left");
            Dictionary<string,GanttDrawing> visibleItems=new Dictionary<string,GanttDrawing>(StringComparer.Ordinal);
            foreach(GanttRowDrawing row in rows)
            {
                double indent=Math.Min(row.Row.Indent*font*.8,Math.Max(0,columnW-4-font*1.2)),labelY=row.Y+4;
                GanttTextFit(layout,"GANTT_ACTIVITY",row.Row.Id,row.Row.Id,row.Row.Label,left+indent,labelY,columnW-4-indent,textH,font,Ink,"left");
                col=1;
                if(g.ShowResponsible){GanttTextFit(layout,"GANTT_RESPONSIBLE",row.Row.Id,row.Row.Id,row.Row.Responsible,left+col*columnW,labelY,columnW-4,textH,font,Ink,"left");col++;}
                if(g.ShowRemark)GanttTextFit(layout,"GANTT_REMARK",row.Row.Id,row.Row.Id,row.Row.Remark,left+col*columnW,labelY,columnW-4,textH,font,Ink,"left");
                if(row.Row.Separator)GanttLine(layout,"GANTT_ROW_SEPARATOR",row.Row.Id,row.Row.Id,left,row.Y,right-left,0,Gray);
                foreach(GanttDrawing item in row.Items)
                {
                    item.Y=row.Y+6+row.Lanes[item.Lane].Offset+item.ShapeHeight/2;
                    GanttDrawItem(layout,item,font);
                    if(item.DrawShape)visibleItems.Add(item.Item.Id,item);
                }
            }
            if(g.Links!=null)for(int i=0;i<g.Links.Count;i++)
            {
                GanttLink link=g.Links[i];GanttDrawing from,to;
                if(visibleItems.TryGetValue(link.FromId,out from)&&visibleItems.TryGetValue(link.ToId,out to))GanttDrawLink(layout,link,i,from,to,plotX,plotW,bodyTop,bodyBottom);
            }
        }

        private static GanttDrawing GanttPrepareItem(GanttItem item,Dictionary<string,GanttDateAnchor> anchors,GanttCalendar calendar,bool dayScale,double x,double w,double font,double shapeH)
        {
            string startText=GanttEndpoint(item,false,anchors),endText=GanttEndpoint(item,true,anchors);
            bool exact=item.DateMode=="Exact",milestone=item.Kind=="Milestone";
            DateTime start=exact?GanttDates.ParseInstant(startText):GanttDates.ParseDay(startText);
            DateTime end=milestone?start:exact?GanttDates.ParseInstant(endText):GanttDates.ParseDay(endText);
            int startDay=GanttDates.Day(start),endDay=GanttDates.Day(end);
            bool endsBefore=exact?endDay<calendar.StartDay||endDay==calendar.StartDay&&end.TimeOfDay.Ticks==0:endDay<calendar.StartDay;
            if(milestone&&(startDay<calendar.StartDay||startDay>=calendar.EndDayExclusive) || !milestone&&(endsBefore||startDay>=calendar.EndDayExclusive))return null;
            double first=exact?calendar.Fraction(start):calendar.BoundaryFraction(startDay),last=milestone?first:exact?calendar.Fraction(end):calendar.BoundaryFraction(endDay+1);
            if(milestone&&!exact&&dayScale)first=last=calendar.CenterFraction(startDay);
            bool visible=milestone?calendar.VisibleDay(startDay):last>first;
            bool rightOverflow=exact?endDay>calendar.EndDayExclusive||endDay==calendar.EndDayExclusive&&end.TimeOfDay.Ticks>0:endDay+1>calendar.EndDayExclusive;
            GanttDrawing result=new GanttDrawing {Item=item,X1=x+first*w,X2=x+last*w,DrawShape=visible,LeftOverflow=startDay<calendar.StartDay,RightOverflow=rightOverflow,Color=item.Color==null?Blue:PieColor(item.Color),Label=CleanPieText(item.Label),Dates="",ShapeHeight=shapeH};
            if(item.ShowDates)result.Dates=milestone?startText:startText+" – "+endText;
            if(!visible&&result.Label.Length==0&&result.Dates.Length==0)return null;
            result.InternalLabel=visible&&result.Label.Length!=0&&(item.Kind=="Pentagon"||item.Kind=="Chevron");
            if(result.InternalLabel)GanttFitProcessLabel(result,font);
            double desired=Math.Min(w,Math.Max(result.InternalLabel?0:result.Label.Length*font*1.2,result.Dates.Length==0?0:Math.Min(w,result.Dates.Length*font*1.2)));
            result.LabelW=Math.Max(0,desired);result.LabelX=Math.Max(x,Math.Min(x+w-desired,result.X1));
            if(!result.InternalLabel)result.Label=FitPieText(result.Label,result.LabelW,font);
            result.DateH=result.Dates.Length==0?0:Math.Ceiling(result.Dates.Length*font*1.2/Math.Max(1,result.LabelW))*font*1.4;
            result.Height=result.ShapeHeight+(result.InternalLabel||result.Label.Length==0?0:font*1.4+2)+(result.DateH==0?0:result.DateH+2);
            return result;
        }
        private static void GanttFitProcessLabel(GanttDrawing item,double font)
        {
            string original=item.Label;double width=item.X2-item.X1;
            List<string> elements=new List<string>();TextElementEnumerator enumerator=StringInfo.GetTextElementEnumerator(original);
            while(enumerator.MoveNext())
            {
                string element=enumerator.GetTextElement();
                // Framework 4 text elements preserve combining marks and surrogate
                // pairs, but predate modern extended grapheme rules. Preserve explicit
                // ZWJ chains as well; this is not a complete Unicode UAX #29 segmenter.
                if(elements.Count>0&&(element[0]=='\u200D'||elements[elements.Count-1].EndsWith("\u200D",StringComparison.Ordinal)))elements[elements.Count-1]+=element;
                else elements.Add(element);
            }
            for(int iteration=0;iteration<64;iteration++)
            {
                double cap=Math.Min(item.ShapeHeight/2,width/3);
                double leftInset=item.Item.Kind=="Chevron"||item.LeftOverflow?cap:0;
                item.InnerX=item.X1+leftInset+2;item.InnerW=width-leftInset-cap-4;
                int capacity=(int)Math.Floor(item.InnerW/(font*1.2));
                if(capacity<1)throw new ArgumentException("Gantt process-arrow text cannot fit inside the arrow at its configured font size. Increase chart width or the item's displayed duration.");
                List<string> lines=new List<string>();string line="";
                foreach(string element in elements)
                {
                    if(element.Length>capacity)throw new ArgumentException("Gantt process-arrow text cannot fit a complete text element. Increase chart width or the item's displayed duration.");
                    if(line.Length+element.Length>capacity){lines.Add(line);line="";}
                    line+=element;
                }
                if(line.Length!=0)lines.Add(line);
                double needed=lines.Count*font*1.4+4;
                if(needed<=item.ShapeHeight)
                {
                    item.Label=String.Join("\n",lines.ToArray());return;
                }
                item.ShapeHeight=needed;
            }
            throw new ArgumentException("Gantt process-arrow text is too dense. Increase chart width or simplify the label.");
        }
        private static List<GanttShadeDrawing> GanttPrepareShades(GanttSpec g,GanttCalendar calendar,double x,double w,double font)
        {
            List<GanttShadeDrawing> result=new List<GanttShadeDrawing>();List<double> laneEnds=new List<double>();
            if(g.Shades==null)return result;
            for(int i=0;i<g.Shades.Count;i++)
            {
                GanttShade shade=g.Shades[i];double a=calendar.BoundaryFraction(GanttDates.Day(GanttDates.ParseDay(shade.StartDate))),b=calendar.BoundaryFraction(GanttDates.Day(GanttDates.ParseDay(shade.EndDate))+1);
                if(b<=a)continue;
                GanttShadeDrawing draw=new GanttShadeDrawing {Shade=shade,X=x+a*w,W=(b-a)*w,Id=i.ToString(Invariant)};
                draw.Label=FitPieText(CleanPieText(shade.Label),draw.W,font);
                if(draw.Label.Length!=0)
                {
                    int lane=0;while(lane<laneEnds.Count&&laneEnds[lane]+4>draw.X)lane++;
                    if(lane==laneEnds.Count)laneEnds.Add(0);laneEnds[lane]=draw.X+draw.W;draw.Lane=lane;
                }
                result.Add(draw);
            }
            return result;
        }
        private static void GanttDrawScales(Layout layout,GanttCalendar calendar,List<GanttScale> scales,double x,double w,double y,double h,double bottom,double font)
        {
            int coarsestRank=GanttUnits.Length;
            foreach(GanttScale shown in scales)coarsestRank=Math.Min(coarsestRank,Array.IndexOf(GanttUnits,shown.Unit));
            for(int index=0;index<scales.Count;index++)
            {
                GanttScale scale=scales[index];List<GanttPeriod> periods=calendar.Periods(scale.Unit);
                if(!GanttScaleFits(calendar,scale,w))throw new ArgumentException("Gantt calendar scale is too dense for complete date labels at 6 points. Use a coarser scale or increase chart width.");
                double yy=y+index*h;
                foreach(GanttPeriod period in periods)
                {
                    double xx=x+calendar.BoundaryFraction(period.StartDay)*w,ww=(calendar.BoundaryFraction(period.EndDayExclusive)-calendar.BoundaryFraction(period.StartDay))*w;
                    if(ww<=0)continue;
                    string id=scale.Unit+":"+period.StartDay.ToString(Invariant),text=CleanPieText(calendar.Format(period.LabelDate,scale.Format??GanttDefaultFormat(scale.Unit)));
                    double f=Math.Max(6,Math.Min(font,(ww-4)/Math.Max(1,text.Length*1.2)));
                    if(ww>=text.Length*6*1.2+4)GanttText(layout,"GANTT_SCALE_LABEL",id,null,text,xx+1,yy+2,Math.Max(0,ww-2),h-4,f,Ink,"left",true);
                    if(scale.Separators&&period.StartDay>calendar.StartDay)
                    {
                        GanttLine(layout,"GANTT_SCALE_SEPARATOR",id,null,xx,yy,0,bottom-yy,Gray);
                        // Original relative-granularity policy: the broadest visible
                        // unit is solid; finer units use Office msoLineDash (4).
                        if(Array.IndexOf(GanttUnits,scale.Unit)>coarsestRank)layout.Items[layout.Items.Count-1].DashStyle=4;
                    }
                }
                GanttLine(layout,"GANTT_SCALE_EDGE",scale.Unit,null,x,yy+h,w,0,Ink);
            }
        }
        private static void GanttDrawItem(Layout layout,GanttDrawing item,double font)
        {
            double h=item.ShapeHeight;
            string id=item.Item.Id,row=item.Item.RowId;double x=item.X1,w=item.X2-item.X1,y=item.Y-h/2;
            if(item.DrawShape)
            {
                if(item.Item.Kind=="Milestone")GanttPolygon(layout,"GANTT_ITEM",id,row,item.Color,new double[,]{{x,item.Y-h/2},{x+h/2,item.Y},{x,item.Y+h/2},{x-h/2,item.Y}});
                else if(item.Item.Kind=="Bracket")
                {
                    GanttLine(layout,"GANTT_ITEM",id,row,x,y,w,0,item.Color);
                    GanttLine(layout,"GANTT_ITEM",id,row,x,y,0,h,item.Color);
                    GanttLine(layout,"GANTT_ITEM",id,row,x+w,y,0,h,item.Color);
                }
                else if(item.LeftOverflow||item.RightOverflow)
                {
                    double cap=Math.Min(h/2,w/3),leftInset=item.LeftOverflow?cap:0;
                    bool rightTip=item.RightOverflow||item.Item.Kind!="Bar";
                    double rightInset=rightTip?cap:0;
                    GanttPolygon(layout,"GANTT_ITEM",id,row,item.Color,new double[,]{{x+leftInset,y},{x+w-rightInset,y},{x+w,rightTip?item.Y:y},{x+w-rightInset,y+h},{x+leftInset,y+h},{x,item.LeftOverflow?item.Y:y+h}});
                }
                else if(item.Item.Kind=="Bar")GanttRect(layout,"GANTT_ITEM",id,row,x,y,w,h,item.Color);
                else
                {
                    double tip=Math.Min(h/2,w/3);
                    if(item.Item.Kind=="Pentagon")GanttPolygon(layout,"GANTT_ITEM",id,row,item.Color,new double[,]{{x,y},{x+w-tip,y},{x+w,y+h/2},{x+w-tip,y+h},{x,y+h}});
                    else GanttPolygon(layout,"GANTT_ITEM",id,row,item.Color,new double[,]{{x,y},{x+w-tip,y},{x+w,y+h/2},{x+w-tip,y+h},{x,y+h},{x+tip,y+h/2}});
                }
                if(item.Item.Kind!="Milestone")
                {
                    double cap=Math.Min(h/2,w/3);
                    if(item.LeftOverflow)GanttPolygon(layout,"GANTT_OVERFLOW_START",id,row,item.Color,new double[,]{{x,item.Y},{x+cap,y},{x+cap,y+h}});
                    if(item.RightOverflow)GanttPolygon(layout,"GANTT_OVERFLOW_END",id,row,item.Color,new double[,]{{x+w,item.Y},{x+w-cap,y},{x+w-cap,y+h}});
                }
            }
            double textY=y+h+2;
            if(item.InternalLabel)GanttText(layout,"GANTT_ITEM_LABEL",id,row,item.Label,item.InnerX,y+2,item.InnerW,h-4,font,GanttContrast(item.Color),"center",false);
            else if(item.Label.Length!=0){GanttText(layout,"GANTT_ITEM_LABEL",id,row,item.Label,item.LabelX,textY,item.LabelW,font*1.4,font,Ink,"left",true);textY+=font*1.4+2;}
            if(item.Dates.Length!=0)GanttText(layout,"GANTT_ITEM_DATES",id,row,item.Dates,item.LabelX,textY,item.LabelW,item.DateH,font,Ink,"left",false);
        }
        private static int GanttContrast(int color)
        {
            double[] channels={(color&255)/255.0,((color>>8)&255)/255.0,((color>>16)&255)/255.0};
            for(int i=0;i<channels.Length;i++)channels[i]=channels[i]<=.04045?channels[i]/12.92:Math.Pow((channels[i]+.055)/1.055,2.4);
            return channels[0]*.2126+channels[1]*.7152+channels[2]*.0722>.179?0:0xFFFFFF;
        }
        private static void GanttDrawLink(Layout layout,GanttLink link,int index,GanttDrawing from,GanttDrawing to,double left,double width,double top,double bottom)
        {
            double x1=link.FromEdge=="Start"?from.X1:from.X2,x2=link.ToEdge=="Start"?to.X1:to.X2;
            double direction=link.ToEdge=="Start"?1:-1,elbow=x2-direction*8;
            elbow=Math.Max(left-6,Math.Min(left+width+6,elbow));
            int color=link.Color==null?Ink:PieColor(link.Color);string id=index.ToString(Invariant),row=from.Item.RowId;
            if(Math.Abs(from.Y-to.Y)<1e-7)
            {
                double bypass=Math.Min(bottom-3,Math.Max(top+3,from.Y-from.ShapeHeight/2-3));
                GanttLine(layout,"GANTT_LINK",id,row,x1,from.Y,0,bypass-from.Y,color);
                GanttLine(layout,"GANTT_LINK",id,row,x1,bypass,elbow-x1,0,color);
                GanttLine(layout,"GANTT_LINK",id,row,elbow,bypass,0,to.Y-bypass,color);
            }
            else
            {
                GanttLine(layout,"GANTT_LINK",id,row,x1,from.Y,elbow-x1,0,color);
                GanttLine(layout,"GANTT_LINK",id,row,elbow,from.Y,0,to.Y-from.Y,color);
            }
            GanttLine(layout,"GANTT_LINK",id,row,elbow,to.Y,x2-elbow,0,color);
            GanttPolygon(layout,"GANTT_LINK_ARROW",id,row,color,new double[,]{{x2,to.Y},{x2-direction*5,to.Y-2.5},{x2-direction*5,to.Y+2.5}});
        }
        private static void GanttRect(Layout layout,string role,string id,string row,double x,double y,double w,double h,int color)
        {layout.Items.Add(new Primitive {Kind="rect",Role=role,EntityId=id,ColumnId=row,X=x,Y=y,W=w,H=h,Color=color});}
        private static void GanttLine(Layout layout,string role,string id,string row,double x,double y,double w,double h,int color)
        {layout.Items.Add(new Primitive {Kind="line",Role=role,EntityId=id,ColumnId=row,X=x,Y=y,W=w,H=h,Color=color,LineWidth=.75});}
        private static void GanttTextFit(Layout layout,string role,string id,string row,string text,double x,double y,double w,double h,double font,int color,string align)
        {GanttText(layout,role,id,row,FitPieText(CleanPieText(text),w,font),x,y,w,h,font,color,align,true);}
        private static void GanttText(Layout layout,string role,string id,string row,string text,double x,double y,double w,double h,double font,int color,string align,bool nowrap)
        {if(!String.IsNullOrEmpty(text))layout.Items.Add(new Primitive {Kind="text",Role=role,EntityId=id,ColumnId=row,Text=text,X=x,Y=y,W=w,H=h,FontSize=font,Color=color,Alignment=align,VerticalAlignment="center",NoWrap=nowrap,MarginLeft=0,MarginRight=0});}
        private static void GanttPolygon(Layout layout,string role,string id,string row,int color,double[,] vertices)
        {
            ShapePath path=new ShapePath {StartX=vertices[0,0],StartY=vertices[0,1]};double left=path.StartX,right=left,top=path.StartY,bottom=top;
            for(int i=1;i<vertices.GetLength(0);i++){double x=vertices[i,0],y=vertices[i,1];path.Segments.Add(new PathSegment {X=x,Y=y});left=Math.Min(left,x);right=Math.Max(right,x);top=Math.Min(top,y);bottom=Math.Max(bottom,y);}
            path.Segments.Add(new PathSegment {X=path.StartX,Y=path.StartY});
            layout.Items.Add(new Primitive {Kind="path",Role=role,EntityId=id,ColumnId=row,Path=path,X=left,Y=top,W=right-left,H=bottom-top,Color=color});
        }
    }
}
