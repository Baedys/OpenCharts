using System;
using System.Collections.Generic;
using System.Globalization;

namespace OpenSlideCharts
{
    public sealed class GanttSpec
    {
        public string StartDate, EndDate;
        public string Culture = "fr-FR", WeekendDisplay = "Shaded", ScaleMode = "Auto";
        public int WeekStartsOn = 1;
        public List<int> WeekendDays = new List<int> { 0, 6 };
        public List<GanttScale> Scales = new List<GanttScale>();
        public List<GanttActivity> Activities = new List<GanttActivity>();
        public GanttActivity MilestoneRow = new GanttActivity { Id = "__milestones", Label = "Jalons" };
        public List<GanttDateAnchor> DateAnchors = new List<GanttDateAnchor>();
        public List<GanttItem> Items = new List<GanttItem>();
        public List<GanttShade> Shades = new List<GanttShade>();
        public List<GanttLink> Links = new List<GanttLink>();
        public bool ShowResponsible, ShowRemark;
        public string ActivityHeader = "Activité", ResponsibleHeader = "Responsable", RemarkHeader = "Remarque";
        public double LabelFraction = .28, FontSize = 10;
    }
    public sealed class GanttScale
    {
        public string Unit = "Months", Format;
        public bool Separators = true;
    }
    public sealed class GanttActivity
    {
        public string Id, Label, Responsible, Remark, Color;
        public int Indent;
        public bool Separator;
    }
    public sealed class GanttItem
    {
        public string Id, RowId, Kind = "Bar", StartDate, EndDate, Label, Color, DateMode = "InclusiveDays", StartAnchor, EndAnchor;
        public bool ShowDates;
    }
    public sealed class GanttShade
    {
        public string StartDate, EndDate, Label, Color;
    }
    public sealed class GanttLink
    {
        public string FromId, ToId, FromEdge = "End", ToEdge = "Start", Color;
    }
    public sealed class GanttDateAnchor
    {
        public string Id, RowId, Date;
        public int Index;
    }
    public static class GanttSamples
    {
        public static ChartSpec Create()
        {
            DateTime start = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            // Three calendar months, matching the documented initial date range.
            if (start.Year == 9999 && start.Month > 9) start = new DateTime(9999, 10, 1);
            DateTime end = start.Year == 9999 && start.Month == 10 ? new DateTime(9999, 12, 31) : start.AddMonths(3).AddDays(-1);
            GanttSpec g = new GanttSpec { StartDate = Iso(start), EndDate = Iso(end) };
            g.Activities.Add(new GanttActivity { Id = "design", Label = "Conception", Responsible = "Produit" });
            g.Activities.Add(new GanttActivity { Id = "build", Label = "Réalisation", Responsible = "Équipe" });
            g.Activities.Add(new GanttActivity { Id = "release", Label = "Livraison" });
            int days = (end - start).Days;
            g.Items.Add(new GanttItem { Id = "a", RowId = "design", StartDate = Iso(start.AddDays(2)), EndDate = Iso(start.AddDays(days / 3)), Label = "Étude" });
            g.Items.Add(new GanttItem { Id = "b", RowId = "build", Kind = "Chevron", StartDate = Iso(start.AddDays(days / 3 + 1)), EndDate = Iso(start.AddDays(days * 2 / 3)), Label = "Développement", Color = "#468C75" });
            g.Items.Add(new GanttItem { Id = "m", RowId = "release", Kind = "Milestone", StartDate = Iso(start.AddDays(days * 2 / 3 + 1)), Label = "Version", Color = "#C05A50" });
            return new ChartSpec { Kind = "Gantt", Title = "Calendrier du projet", Gantt = g };
        }
        private static string Iso(DateTime value) { return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); }
    }
}
