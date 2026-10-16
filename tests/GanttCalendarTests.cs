using System;
using System.Collections.Generic;
using OpenSlideCharts;

public static class GanttCalendarTests
{
    private static int scenarios, assertions;
    public static int Main()
    {
        try
        {
            StrictDates(); Instants(); LeapYears(); IsoWeeks(); InclusiveProjection(); HiddenWeekends(); CustomWeekdays();
            NaturalPeriods(); CompressedPeriods(); DateLimits(); RangePolicy(); CultureAndTokens(); LiteralFormats();
            InvalidCalendars(); InvalidFormats(); SnapshotIsolation();
            Console.WriteLine("PASS: " + scenarios + " Gantt calendar scenarios, " + assertions + " assertions"); return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
    private static void Check(bool condition, string message) { assertions++; if (!condition) throw new Exception(message); }
    private static void Near(double value, double expected, string message)
    { Check(!Double.IsNaN(value) && Math.Abs(value - expected) < 1e-13, message + ": " + value + " vs " + expected); }
    private static T Reject<T>(Action action, string message) where T : Exception
    { try { action(); } catch (T error) { assertions++; return error; } throw new Exception(message); }
    private static DateTime D(string date) { return GanttDates.ParseDay(date); }
    private static int Day(string date) { return GanttDates.Day(D(date)); }
    private static GanttSpec Spec(string first, string last, string weekend = "Plain")
    { return new GanttSpec { StartDate = first, EndDate = last, WeekendDisplay = weekend }; }
    private static void StrictDates()
    {
        scenarios++; DateTime date = D("2024-02-29"); Check(date.Year == 2024 && date.Month == 2 && date.Day == 29 && date.Kind == DateTimeKind.Unspecified, "strict Gregorian date");
        Check(GanttDates.Iso(date) == "2024-02-29" && GanttDates.Iso(date.AddHours(19)) == "2024-02-29T19:00:00", "Iso preserves nonmidnight time");
        Check(GanttDates.Day(DateTime.MinValue) == 0 && GanttDates.Day(DateTime.MinValue.AddDays(1)) == 1, "zero-based ordinal");
        foreach (string invalid in new string[] { "2024-2-29", "24-02-29", " 2024-02-29", "2024-02-29 ", "2024-02-29\n", "2023-02-29", "1900-02-29", "0000-01-01", "10000-01-01", "2024-13-01", "2024-01-00", "2024-01-01T00:00:00", "２０２４-01-01" })
            Reject<FormatException>(() => D(invalid), "invalid day accepted: " + invalid);
        Reject<ArgumentNullException>(() => D(null), "null day");
    }
    private static void Instants()
    {
        scenarios++; DateTime instant = GanttDates.ParseInstant("2024-01-01T12:34:56.1234567");
        Check(instant.TimeOfDay.Ticks == (12 * 3600L + 34 * 60 + 56) * TimeSpan.TicksPerSecond + 1234567, "seven fractional-second digits preserved");
        Check(instant.Kind == DateTimeKind.Unspecified, "no time-zone conversion");
        Check(GanttDates.Iso(instant) == "2024-01-01T12:34:56.1234567" && GanttDates.ParseInstant(GanttDates.Iso(instant)).Ticks == instant.Ticks, "ISO export preserves all intraday ticks");
        Check(GanttDates.Iso(instant.AddTicks(-1234567 + 1000000)) == "2024-01-01T12:34:56.1", "fraction output trims trailing zeros only");
        Check(GanttDates.ParseInstant(GanttDates.Iso(DateTime.MaxValue)) == DateTime.MaxValue, "maximum instant ISO roundtrip");
        Check(GanttDates.ParseInstant("2024-01-01") == D("2024-01-01"), "date accepted as midnight instant");
        Check(GanttDates.ParseInstant("2024-01-01T00:00:00.1").Ticks % TimeSpan.TicksPerSecond == 1000000, "short fractional seconds");
        GanttCalendar calendar = new GanttCalendar(Spec("2024-01-01", "2024-01-01"));
        Near(calendar.Fraction(DateTime.FromOADate(45292.25)), .25, "Excel serial quarter day preserved");
        Near(calendar.Fraction(GanttDates.ParseInstant("2024-01-01T12:00:00")), .5, "instant noon differs from day boundary");
        Near(calendar.Fraction(D("2024-01-01")), 0, "day is midnight boundary");
        Check(calendar.Fraction(D("2024-01-01").AddTicks(1)) > 0, "single intraday tick not truncated");
        foreach (string invalid in new string[] { "2024-01-01 12:00:00", "2024-01-01t12:00:00", "2024-01-01T12:00", "2024-01-01T24:00:00", "2024-01-01T12:00:60", "2024-01-01T12:00:00.", "2024-01-01T12:00:00.12345678", "2024-01-01T12:00:00Z", "2024-01-01T12:00:00+01:00", "2024-01-01T12:00:00\n" })
            Reject<FormatException>(() => GanttDates.ParseInstant(invalid), "invalid instant accepted: " + invalid);
    }
    private static void LeapYears()
    {
        scenarios++; GanttCalendar leap = new GanttCalendar(Spec("2000-02-28", "2000-03-01"));
        GanttCalendar normal = new GanttCalendar(Spec("1900-02-28", "1900-03-01"));
        Near(leap.TotalUnits, 3, "year 2000 leap"); Near(normal.TotalUnits, 2, "year 1900 common");
        Check(leap.Periods("Days")[1].LabelDate == D("2000-02-29"), "leap day period exists");
        Check(normal.Periods("Days")[1].LabelDate == D("1900-03-01"), "century exception honored");
    }
    private static void IsoWeeks()
    {
        scenarios++;
        string[] dates = { "0001-01-01", "2015-12-31", "2016-01-01", "2016-01-04", "2017-01-01", "2019-12-30", "2021-01-01", "9999-01-01", "9999-12-31" };
        int[] years = { 1, 2015, 2015, 2016, 2016, 2020, 2020, 9998, 9999 }, weeks = { 1, 53, 53, 1, 52, 1, 53, 53, 52 };
        for (int i = 0; i < dates.Length; i++)
        { Check(GanttDates.IsoWeekYear(D(dates[i])) == years[i], "ISO week year " + dates[i]); Check(GanttDates.IsoWeekNumber(D(dates[i])) == weeks[i], "ISO week number " + dates[i]); }
        // Every Gregorian year has Jan 4 in ISO week 1 and Dec 28 in its last week.
        for (int year = 1; year <= 9999; year++)
        {
            DateTime jan4 = new DateTime(year, 1, 4), dec28 = new DateTime(year, 12, 28);
            Check(GanttDates.IsoWeekNumber(jan4) == 1 && GanttDates.IsoWeekYear(jan4) == year, "ISO first week invariant");
            int last = GanttDates.IsoWeekNumber(dec28); Check((last == 52 || last == 53) && GanttDates.IsoWeekYear(dec28) == year, "ISO last week invariant");
        }
    }
    private static void InclusiveProjection()
    {
        scenarios++; GanttCalendar calendar = new GanttCalendar(Spec("2024-01-01", "2024-01-03"));
        Check(calendar.StartDay == Day("2024-01-01") && calendar.EndDayExclusive == Day("2024-01-04"), "inclusive chart end converted to exclusive ordinal");
        Near(calendar.TotalUnits, 3, "all inclusive days count"); Near(calendar.BoundaryFraction(Day("2024-01-02")), 1.0 / 3, "boundary fraction");
        Near(calendar.CenterFraction(Day("2024-01-02")), .5, "day center");
        Near(calendar.BoundaryFraction(Int32.MinValue), 0, "lower integer clamp"); Near(calendar.BoundaryFraction(Int32.MaxValue), 1, "upper integer clamp");
        Near(calendar.Fraction(DateTime.MinValue), 0, "earlier instant clamped"); Near(calendar.Fraction(DateTime.MaxValue), 1, "later instant clamped");
        Check(!calendar.VisibleDay(Day("2023-12-31")) && !calendar.VisibleDay(calendar.EndDayExclusive), "outside range invisible");
        Near(calendar.CenterFraction(Int32.MinValue), 0, "center lower clamp"); Near(calendar.CenterFraction(Int32.MaxValue), 1, "center upper clamp");
    }
    private static void HiddenWeekends()
    {
        scenarios++; GanttSpec spec = Spec("2024-03-01", "2024-03-04", "Hidden");
        spec.Shades.Add(new GanttShade { StartDate = "2024-03-01", EndDate = "2024-03-01", Label = "Holiday shade" });
        GanttCalendar calendar = new GanttCalendar(spec); Near(calendar.TotalUnits, 2, "only weekend weekdays compressed; holiday shade not compressed");
        Check(calendar.VisibleDay(Day("2024-03-01")) && !calendar.VisibleDay(Day("2024-03-02")) && !calendar.VisibleDay(Day("2024-03-03")), "Friday visible, weekend hidden");
        Near(calendar.BoundaryFraction(Day("2024-03-02")), .5, "Saturday compressed boundary"); Near(calendar.BoundaryFraction(Day("2024-03-04")), .5, "Monday shares compressed boundary");
        Near(calendar.Fraction(GanttDates.ParseInstant("2024-03-01T12:00:00")), .25, "Friday intraday position");
        Near(calendar.Fraction(GanttDates.ParseInstant("2024-03-03T23:59:59")), .5, "all hidden-day times collapse");
        Near(calendar.Fraction(GanttDates.ParseInstant("2024-03-04T18:00:00")), .875, "Monday intraday position");
        Near(calendar.CenterFraction(Day("2024-03-02")), .5, "hidden center collapses"); Near(calendar.CenterFraction(Day("2024-03-04")), .75, "visible center");
        foreach (string display in new string[] { "Plain", "Shaded" }) Near(new GanttCalendar(Spec("2024-03-01", "2024-03-04", display)).TotalUnits, 4, "display mode preserves weekends");
        Reject<ArgumentException>(() => new GanttCalendar(Spec("2024-03-02", "2024-03-03", "Hidden")), "no visible day rejected");
    }
    private static void CustomWeekdays()
    {
        scenarios++; GanttSpec spec = Spec("2024-01-01", "2024-01-05", "Hidden"); spec.WeekendDays = new List<int> { 2, 3 };
        GanttCalendar calendar = new GanttCalendar(spec); Near(calendar.TotalUnits, 3, "custom Tuesday/Wednesday weekends");
        Check(!calendar.VisibleDay(Day("2024-01-02")) && calendar.VisibleDay(Day("2024-01-04")), "custom hidden weekdays mapped correctly");
        GanttSpec sunday = Spec("2024-01-08", "2024-01-15"); sunday.WeekStartsOn = 0;
        GanttCalendar sundayCalendar = new GanttCalendar(sunday); List<GanttPeriod> periods = sundayCalendar.Periods("Weeks");
        Check(periods.Count == 2 && periods[0].LabelDate == D("2024-01-07") && periods[1].LabelDate == D("2024-01-14"), "Sunday week boundaries");
        Check(sundayCalendar.Format(periods[0].LabelDate, "ww") == "01", "ISO number remains independent of week start");
        GanttSpec thursday = Spec("2024-01-01", "2024-01-15"); thursday.WeekStartsOn = 4;
        periods = new GanttCalendar(thursday).Periods("Weeks"); Check(periods.Count == 3 && periods[0].LabelDate == D("2023-12-28"), "Thursday period starts before range");
    }
    private static void NaturalPeriods()
    {
        scenarios++; GanttCalendar calendar = new GanttCalendar(Spec("2023-12-15", "2024-04-03"));
        List<GanttPeriod> months = calendar.Periods("Months"), quarters = calendar.Periods("Quarters"), years = calendar.Periods("Years");
        Check(months.Count == 5 && months[0].LabelDate == D("2023-12-01") && months[4].LabelDate == D("2024-04-01"), "calendar months retain natural starts");
        Check(months[0].StartDay == calendar.StartDay && months[4].EndDayExclusive == calendar.EndDayExclusive, "period bounds intersect chart");
        Check(quarters.Count == 3 && quarters[0].LabelDate == D("2023-10-01") && quarters[1].LabelDate == D("2024-01-01"), "quarter boundaries");
        Check(years.Count == 2 && years[0].LabelDate == D("2023-01-01") && years[1].LabelDate == D("2024-01-01"), "year boundaries");
        for (int i = 1; i < months.Count; i++) Check(months[i - 1].EndDayExclusive == months[i].StartDay, "months cover without gaps");
        Reject<ArgumentException>(() => calendar.Periods("Fortnights"), "unknown period unit rejected");
    }
    private static void CompressedPeriods()
    {
        scenarios++; GanttCalendar calendar = new GanttCalendar(Spec("2024-03-01", "2024-03-04", "Hidden"));
        List<GanttPeriod> days = calendar.Periods("Days"); Check(days.Count == 2 && days[0].LabelDate == D("2024-03-01") && days[1].LabelDate == D("2024-03-04"), "hidden day periods omitted");
        foreach (string unit in new string[] { "Years", "Quarters", "Months", "Weeks", "Days" })
        {
            double sum = 0; foreach (GanttPeriod period in calendar.Periods(unit)) { double width = calendar.BoundaryFraction(period.EndDayExclusive) - calendar.BoundaryFraction(period.StartDay); Check(width > 0, "no zero-width compressed period"); sum += width; }
            Near(sum, 1, "compressed periods cover visible domain");
        }
        GanttSpec emptyWeekend = Spec("2024-03-02", "2024-03-03", "Hidden"); emptyWeekend.WeekendDays.Clear(); Near(new GanttCalendar(emptyWeekend).TotalUnits, 2, "empty weekend list hides nothing");
    }
    private static void DateLimits()
    {
        scenarios++; GanttSpec minimum = Spec("0001-01-01", "0001-01-02"); minimum.WeekStartsOn = 0;
        GanttCalendar first = new GanttCalendar(minimum); Check(first.Periods("Weeks")[0].LabelDate == DateTime.MinValue, "unrepresentable year-zero week label clamped safely");
        GanttCalendar last = new GanttCalendar(Spec("9999-12-30", "9999-12-31"));
        Check(last.EndDayExclusive == GanttDates.Day(DateTime.MaxValue) + 1, "last day inclusive without DateTime overflow");
        Near(last.BoundaryFraction(last.EndDayExclusive), 1, "maximum exclusive boundary");
        Check(last.Fraction(DateTime.MaxValue) < 1 && last.Fraction(DateTime.MaxValue) > .999999999, "last representable instant retains time");
        foreach (string unit in new string[] { "Years", "Quarters", "Months", "Weeks", "Days" })
        {
            List<GanttPeriod> periods = last.Periods(unit); Check(periods.Count > 0 && periods[periods.Count - 1].EndDayExclusive == last.EndDayExclusive, "period reaches max date safely");
        }
        Check(last.Periods("Quarters")[0].LabelDate == D("9999-10-01"), "last quarter label");
    }
    private static void RangePolicy()
    {
        scenarios++; DateTime start = D("2000-01-01"); GanttCalendar cap = new GanttCalendar(Spec(GanttDates.Iso(start), GanttDates.Iso(start.AddDays(36599))));
        Near(cap.TotalUnits, 36600, "original 36600-day cap inclusive");
        Reject<ArgumentException>(() => new GanttCalendar(Spec(GanttDates.Iso(start), GanttDates.Iso(start.AddDays(36600)))), "range beyond cap rejected");
        Reject<ArgumentException>(() => new GanttCalendar(Spec("2024-01-02", "2024-01-01")), "reversed dates rejected");
    }
    private static void CultureAndTokens()
    {
        scenarios++; GanttCalendar french = new GanttCalendar(Spec("2024-03-01", "2024-03-01")); DateTime date = D("2024-03-01");
        Check(french.Format(date, "y yy yyyy q qq m mm mmm mmmm d dd ddd dddd w ww") == "2024 24 2024 1 i 3 03 mars mars 1 01 ven. vendredi 9 09", "French localized tokens and Roman quarter");
        Check(french.Format(date, "YYYY-MM-DD") == "2024-03-01", "tokens case insensitive");
        Check(french.Format(date, "qqq www wwww ddddd mmmmm") == "I Semaine 9 Semaine 09 v m", "Roman quarter, localized week words and initials");
        Check(french.Format(new DateTime(7, 11, 1), "y yy yyyy q qq") == "7 07 0007 4 iv", "unpadded original y extension and quarter IV");
        GanttSpec english = Spec("2024-03-01", "2024-03-01"); english.Culture = "en-US";
        Check(new GanttCalendar(english).Format(date, "mmm mmmm ddd dddd") == "Mar March Fri Friday", "English month and weekday names");
        Check(new GanttCalendar(english).Format(date, "www wwww") == "Week 9 Week 09", "English week words");
        GanttSpec arabic = Spec("2024-03-01", "2024-03-01"); arabic.Culture = "ar-SA";
        Check(new GanttCalendar(arabic).Format(date, "yyyy-mm-dd") == "2024-03-01", "Gregorian fields preserved in non-Gregorian default culture");
    }
    private static void LiteralFormats()
    {
        scenarios++; GanttCalendar calendar = new GanttCalendar(Spec("2024-03-01", "2024-03-01")); DateTime date = D("2024-03-01");
        Check(calendar.Format(date, "'Q'q 'l''année' yyyy 'mm/dd literal'") == "Q1 l'année 2024 mm/dd literal", "quoted text and doubled apostrophe preserved");
        Check(calendar.Format(date, "yyyy / mm - dd (q)") == "2024 / 03 - 01 (1)", "punctuation preserved");
    }
    private static void InvalidCalendars()
    {
        scenarios++; Reject<ArgumentNullException>(() => new GanttCalendar(null), "null spec");
        GanttSpec spec = Spec("2024-01-01", "2024-01-10"); spec.WeekStartsOn = 7; Reject<ArgumentException>(() => new GanttCalendar(spec), "invalid week start");
        spec.WeekStartsOn = 1; spec.WeekendDays = new List<int> { 0, 0 }; Reject<ArgumentException>(() => new GanttCalendar(spec), "duplicate weekend day");
        spec.WeekendDays = new List<int> { -1 }; Reject<ArgumentException>(() => new GanttCalendar(spec), "negative weekday");
        spec.WeekendDays = null; Reject<ArgumentException>(() => new GanttCalendar(spec), "null weekend list");
        spec.WeekendDays = new List<int> { 0, 1, 2, 3, 4, 5, 6 }; spec.WeekendDisplay = "Hidden"; Reject<ArgumentException>(() => new GanttCalendar(spec), "all weekdays hidden");
        spec.WeekendDisplay = "Unknown"; Reject<ArgumentException>(() => new GanttCalendar(spec), "invalid weekend mode");
        spec.WeekendDisplay = "Plain"; spec.Culture = "invalid-culture-name-xyz"; Reject<ArgumentException>(() => new GanttCalendar(spec), "invalid culture");
        spec.Culture = "fr-FR"; spec.Scales.Add(new GanttScale { Unit = "Unknown" }); Reject<ArgumentException>(() => new GanttCalendar(spec), "invalid scale unit");
    }
    private static void InvalidFormats()
    {
        scenarios++; GanttCalendar calendar = new GanttCalendar(Spec("2024-01-01", "2024-01-02")); DateTime date = D("2024-01-01");
        foreach (string format in new string[] { "", "yyy", "qqqq", "mmmmmm", "dddddd", "wwwww", "yyyy 'open", "yyyy\nmm", "'line\nbreak'", "unquoted", new String('d', 257) })
            Reject<ArgumentException>(() => calendar.Format(date, format), "invalid format accepted: " + format);
        Reject<ArgumentNullException>(() => calendar.Format(date, null), "null format");
        GanttSpec spec = Spec("2024-01-01", "2024-01-02"); spec.Scales.Add(new GanttScale { Format = "yyy" });
        Reject<ArgumentException>(() => new GanttCalendar(spec), "scale format validated during construction");
    }
    private static void SnapshotIsolation()
    {
        scenarios++; GanttSpec spec = Spec("2024-03-01", "2024-03-04", "Hidden"); GanttCalendar calendar = new GanttCalendar(spec);
        Check(spec.WeekendDays.Count == 2 && spec.StartDate == "2024-03-01", "constructor leaves input unchanged");
        spec.WeekendDays.Clear(); spec.WeekendDisplay = "Plain"; spec.WeekStartsOn = 4; spec.Culture = "en-US";
        Near(calendar.TotalUnits, 2, "projection snapshot independent of later input changes");
        Check(calendar.Periods("Weeks")[0].LabelDate.DayOfWeek == DayOfWeek.Monday && calendar.Format(D("2024-03-01"), "mmmm") == "mars", "week and culture snapshot independent");
    }
}
