using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OpenSlideCharts
{
    public static class GanttDates
    {
        private static readonly Regex DayPattern = new Regex(@"^[0-9]{4}-[0-9]{2}-[0-9]{2}$", RegexOptions.CultureInvariant);
        private static readonly Regex InstantPattern = new Regex(@"^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(\.[0-9]{1,7})?$", RegexOptions.CultureInvariant);
        public static DateTime ParseDay(string value)
        {
            if (value == null) throw new ArgumentNullException("value");
            DateTime result;
            if (value.Length != 10 || !DayPattern.IsMatch(value) || !DateTime.TryParseExact(value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                throw new FormatException("A Gregorian ISO date yyyy-MM-dd between 0001-01-01 and 9999-12-31 is required.");
            return DateTime.SpecifyKind(result, DateTimeKind.Unspecified);
        }
        public static DateTime ParseInstant(string value)
        {
            if (value == null) throw new ArgumentNullException("value");
            if (value.Length == 10) return ParseDay(value);
            DateTime result;
            string pattern = value.Length == 19 ? "yyyy-MM-dd'T'HH:mm:ss" : "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF";
            if (value.Length < 19 || value.Length > 27 || !InstantPattern.IsMatch(value) || !DateTime.TryParseExact(value, pattern, CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                throw new FormatException("An ISO date or yyyy-MM-ddTHH:mm:ss[.fffffff] without a time zone is required.");
            return DateTime.SpecifyKind(result, DateTimeKind.Unspecified);
        }
        public static int Day(DateTime value) { return (int)(value.Ticks / TimeSpan.TicksPerDay); }
        // Midnight is represented as a date; all other wall-clock ticks survive
        // import/export exactly. No zone suffix or local/UTC conversion is added.
        public static string Iso(DateTime value)
        {
            if (value.TimeOfDay.Ticks == 0) return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            string result = value.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
            long fraction = value.Ticks % TimeSpan.TicksPerSecond;
            if (fraction != 0) result += "." + fraction.ToString("D7", CultureInfo.InvariantCulture).TrimEnd('0');
            return result;
        }
        public static int IsoWeekYear(DateTime value)
        {
            int week = ProvisionalWeek(value);
            if (week < 1) return value.Year - 1;
            return week > WeeksInYear(value.Year) ? value.Year + 1 : value.Year;
        }
        public static int IsoWeekNumber(DateTime value)
        {
            int week = ProvisionalWeek(value);
            if (week < 1) return WeeksInYear(value.Year - 1);
            return week > WeeksInYear(value.Year) ? 1 : week;
        }
        private static int ProvisionalWeek(DateTime value)
        { return (value.DayOfYear - ((int)value.DayOfWeek + 6) % 7 + 9) / 7; }
        private static int WeeksInYear(int year)
        {
            DayOfWeek first = new DateTime(year, 1, 1).DayOfWeek;
            return first == DayOfWeek.Thursday || first == DayOfWeek.Wednesday && DateTime.IsLeapYear(year) ? 53 : 52;
        }
    }

    public sealed class GanttPeriod
    {
        // Bounds intersect the chart. LabelDate remains the natural period start,
        // except a week beginning before year 1 is represented by DateTime.MinValue.
        public int StartDay, EndDayExclusive;
        public DateTime LabelDate;
    }

    /// <summary>
    /// Original Gregorian calendar projection. StartDate/EndDate are inclusive.
    /// The 36600-day range cap is a product policy, not a vendor compatibility claim.
    /// Only Hidden weekend weekdays are compressed; shades/holidays never are.
    /// DateTime values are wall-clock values; no time-zone conversion is performed.
    /// </summary>
    public sealed class GanttCalendar
    {
        public readonly int StartDay, EndDayExclusive;
        public readonly double TotalUnits;
        private static readonly int MaximumDayExclusive = GanttDates.Day(DateTime.MaxValue) + 1;
        private readonly int[] prefix;
        private readonly bool[] visible;
        private readonly int weekStartsOn;
        private readonly CultureInfo culture;

        public GanttCalendar(GanttSpec spec)
        {
            if (spec == null) throw new ArgumentNullException("spec");
            DateTime start = GanttDates.ParseDay(spec.StartDate), end = GanttDates.ParseDay(spec.EndDate);
            StartDay = GanttDates.Day(start); EndDayExclusive = GanttDates.Day(end) + 1;
            int length = EndDayExclusive - StartDay;
            if (length < 1 || length > 36600) throw new ArgumentException("Gantt range must contain 1 to 36600 inclusive calendar days.");
            if (spec.WeekStartsOn < 0 || spec.WeekStartsOn > 6) throw new ArgumentException("WeekStartsOn must be from 0 (Sunday) to 6 (Saturday).");
            weekStartsOn = spec.WeekStartsOn;
            if (spec.WeekendDisplay != "Plain" && spec.WeekendDisplay != "Shaded" && spec.WeekendDisplay != "Hidden")
                throw new ArgumentException("WeekendDisplay must be Plain, Shaded or Hidden.");
            if (spec.WeekendDays == null) throw new ArgumentException("WeekendDays is required; use an empty list for no weekend days.");
            bool[] weekends = new bool[7];
            foreach (int day in spec.WeekendDays)
            {
                if (day < 0 || day > 6 || weekends[day]) throw new ArgumentException("WeekendDays must contain unique weekday numbers from 0 to 6.");
                weekends[day] = true;
            }
            if (spec.Culture == null) throw new ArgumentException("A culture name is required; empty selects the invariant culture.");
            try { culture = (CultureInfo)CultureInfo.GetCultureInfo(spec.Culture).Clone(); }
            catch (CultureNotFoundException error) { throw new ArgumentException("Unknown Gantt culture.", "spec", error); }
            GregorianCalendar gregorian = null;
            foreach (Calendar calendar in culture.OptionalCalendars) if (calendar is GregorianCalendar) { gregorian = (GregorianCalendar)calendar; break; }
            if (gregorian == null) throw new ArgumentException("The selected culture must support a Gregorian calendar.");
            culture.DateTimeFormat.Calendar = gregorian;
            prefix = new int[length + 1]; visible = new bool[length];
            for (int i = 0; i < length; i++)
            {
                visible[i] = spec.WeekendDisplay != "Hidden" || !weekends[(StartDay + i + 1) % 7];
                prefix[i + 1] = prefix[i] + (visible[i] ? 1 : 0);
            }
            TotalUnits = prefix[length];
            if (TotalUnits == 0) throw new ArgumentException("The range contains no visible days after hiding weekend weekdays.");
            if (spec.Scales != null) foreach (GanttScale scale in spec.Scales)
            {
                if (scale == null || !KnownUnit(scale.Unit)) throw new ArgumentException("Unknown or missing Gantt scale unit.");
                if (scale.Format != null) Format(start, scale.Format);
            }
        }

        public double BoundaryFraction(int day)
        {
            if (day <= StartDay) return 0;
            if (day >= EndDayExclusive) return 1;
            return prefix[day - StartDay] / TotalUnits;
        }
        public bool VisibleDay(int day)
        { return day >= StartDay && day < EndDayExclusive && visible[day - StartDay]; }
        public double CenterFraction(int day)
        {
            if (day < StartDay) return 0;
            if (day >= EndDayExclusive) return 1;
            int index = day - StartDay;
            return (prefix[index] + (visible[index] ? .5 : 0)) / TotalUnits;
        }
        public double Fraction(DateTime instant)
        {
            int day = GanttDates.Day(instant);
            if (day < StartDay) return 0;
            if (day >= EndDayExclusive) return 1;
            int index = day - StartDay;
            double time = visible[index] ? (instant.Ticks % TimeSpan.TicksPerDay) / (double)TimeSpan.TicksPerDay : 0;
            return (prefix[index] + time) / TotalUnits;
        }

        public List<GanttPeriod> Periods(string unit)
        {
            if (!KnownUnit(unit)) throw new ArgumentException("Unit must be Years, Quarters, Months, Weeks or Days.", "unit");
            DateTime start = FromDay(StartDay); int natural;
            if (unit == "Years") natural = GanttDates.Day(new DateTime(start.Year, 1, 1));
            else if (unit == "Quarters") natural = GanttDates.Day(new DateTime(start.Year, ((start.Month - 1) / 3) * 3 + 1, 1));
            else if (unit == "Months") natural = GanttDates.Day(new DateTime(start.Year, start.Month, 1));
            else if (unit == "Weeks") natural = StartDay - (((StartDay + 1) % 7 - weekStartsOn + 7) % 7);
            else natural = StartDay;
            List<GanttPeriod> result = new List<GanttPeriod>();
            while (natural < EndDayExclusive)
            {
                int next;
                if (unit == "Days") next = natural + 1;
                else if (unit == "Weeks") next = natural + 7;
                else
                {
                    DateTime label = FromDay(natural);
                    int months = unit == "Years" ? 12 : unit == "Quarters" ? 3 : 1;
                    int monthIndex = (label.Year - 1) * 12 + label.Month - 1 + months;
                    next = monthIndex >= 9999 * 12 ? MaximumDayExclusive : GanttDates.Day(new DateTime(monthIndex / 12 + 1, monthIndex % 12 + 1, 1));
                }
                int clippedStart = Math.Max(StartDay, Math.Max(0, natural));
                int clippedEnd = Math.Min(EndDayExclusive, Math.Min(MaximumDayExclusive, next));
                if (clippedEnd > clippedStart && prefix[clippedEnd - StartDay] > prefix[clippedStart - StartDay])
                    result.Add(new GanttPeriod { StartDay = clippedStart, EndDayExclusive = clippedEnd, LabelDate = FromDay(Math.Max(0, natural)) });
                natural = next;
            }
            return result;
        }
        private static DateTime FromDay(int ordinal) { return new DateTime((long)ordinal * TimeSpan.TicksPerDay, DateTimeKind.Unspecified); }
        private static bool KnownUnit(string value)
        { return value == "Years" || value == "Quarters" || value == "Months" || value == "Weeks" || value == "Days"; }

        /// <summary>
        /// Supported case-insensitive tokens: y/yy/yyyy, q/qq/qqq, m through
        /// mmmmm, d through ddddd, w through wwww. Letters outside tokens must be quoted
        /// with apostrophes; doubled apostrophes inside a literal preserve one.
        /// q is numeric; qq/qqq are lowercase/uppercase Roman. www/wwww prefix
        /// a localized Week word (English fallback for unsupported languages).
        /// The single y token is an original
        /// extension for the unpadded full year; yy retains the last two digits.
        /// Invalid tokens/unclosed literals are rejected rather than discarded.
        /// </summary>
        public string Format(DateTime date, string format)
        {
            if (format == null) throw new ArgumentNullException("format");
            if (format.Length == 0 || format.Length > 256) throw new ArgumentException("A date format must contain 1 to 256 characters.", "format");
            StringBuilder result = new StringBuilder();
            for (int i = 0; i < format.Length;)
            {
                char c = format[i];
                if (Char.IsControl(c)) throw new ArgumentException("Control characters are not supported in date formats.", "format");
                if (c == '\'')
                {
                    i++; bool closed = false;
                    while (i < format.Length)
                    {
                        c = format[i++];
                        if (Char.IsControl(c)) throw new ArgumentException("Control characters are not supported in date literals.", "format");
                        if (c == '\'')
                        {
                            if (i < format.Length && format[i] == '\'') { result.Append('\''); i++; }
                            else { closed = true; break; }
                        }
                        else result.Append(c);
                    }
                    if (!closed) throw new ArgumentException("Unclosed date-format literal.", "format");
                    continue;
                }
                if (!Char.IsLetter(c)) { result.Append(c); i++; continue; }
                int end = i + 1; while (end < format.Length && Char.ToLowerInvariant(format[end]) == Char.ToLowerInvariant(c)) end++;
                int count = end - i; string token = format.Substring(i, count).ToLowerInvariant(); string text;
                switch (token)
                {
                    case "y": text = date.Year.ToString(culture); break;
                    case "yy": text = (date.Year % 100).ToString("D2", culture); break;
                    case "yyyy": text = date.Year.ToString("D4", culture); break;
                    case "q": text = ((date.Month - 1) / 3 + 1).ToString(culture); break;
                    case "qq": text = new string[] { "i", "ii", "iii", "iv" }[(date.Month - 1) / 3]; break;
                    case "qqq": text = new string[] { "I", "II", "III", "IV" }[(date.Month - 1) / 3]; break;
                    case "m": text = date.Month.ToString(culture); break;
                    case "mm": text = date.Month.ToString("D2", culture); break;
                    case "mmm": text = culture.DateTimeFormat.GetAbbreviatedMonthName(date.Month); break;
                    case "mmmm": text = culture.DateTimeFormat.GetMonthName(date.Month); break;
                    case "mmmmm": text = Initial(culture.DateTimeFormat.GetMonthName(date.Month)); break;
                    case "d": text = date.Day.ToString(culture); break;
                    case "dd": text = date.Day.ToString("D2", culture); break;
                    case "ddd": text = culture.DateTimeFormat.GetAbbreviatedDayName(date.DayOfWeek); break;
                    case "dddd": text = culture.DateTimeFormat.GetDayName(date.DayOfWeek); break;
                    case "ddddd": text = Initial(culture.DateTimeFormat.GetDayName(date.DayOfWeek)); break;
                    case "w": text = GanttDates.IsoWeekNumber(date).ToString(culture); break;
                    case "ww": text = GanttDates.IsoWeekNumber(date).ToString("D2", culture); break;
                    case "www": text = WeekWord() + " " + GanttDates.IsoWeekNumber(date).ToString(culture); break;
                    case "wwww": text = WeekWord() + " " + GanttDates.IsoWeekNumber(date).ToString("D2", culture); break;
                    default: throw new ArgumentException("Unsupported date-format token: " + token + ". Quote literal text with apostrophes.", "format");
                }
                result.Append(text); i = end;
            }
            if (result.Length > 1024) throw new ArgumentException("Formatted date exceeds 1024 characters.", "format");
            return result.ToString();
        }
        private static string Initial(string value) { return value.Length == 0 ? "" : StringInfo.GetNextTextElement(value); }
        private string WeekWord()
        {
            switch (culture.TwoLetterISOLanguageName)
            {
                case "fr": return "Semaine";
                case "de": return "Woche";
                case "es": return "Semana";
                case "it": return "Settimana";
                case "pt": return "Semana";
                case "nl": return "Week";
                case "da": return "Uge";
                case "sv": return "Vecka";
                case "no": case "nb": return "Uke";
                default: return "Week";
            }
        }
    }
}
