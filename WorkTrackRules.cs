using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OutlookListener
{
    // Work Track's reminder rules, ported from the Work Track app (web/src/core/rules.ts and
    // schedule.ts in github.com/ridhun30/work-track) so the laptop shows the same reminders as
    // the phone. Pure logic, no Outlook: tested by tests\WorkTrackRulesTests.cs against the same
    // shared vectors the web and Android apps use. Dates are calendar days (DateTime.Date) in
    // the user's Work Track timezone; instants are epoch milliseconds.

    public class WtRecurrence
    {
        public string Freq;          // daily | workdays | weekly | monthly
        public int Interval = 1;
        public List<int> ByDay = new List<int>();   // 0 = Sunday … 6 = Saturday
        public int? DayOfMonth;
    }

    public class WtTask
    {
        public string Id, Title, Notes, Project, Person, Status = "open", DueTime, WaitingOn, SeriesId, BlockedReason, AgendaFor, Source;
        public DateTime? DueDate, TodayDate, FollowUpDate, SnoozedUntil;
        public int? TodayRank;
        public WtRecurrence Recurrence;
        public bool Blocked, Deleted;
        public long CreatedAt, UpdatedAt, LastTouchedAt;
        public long? CompletedAt;
        public Dictionary<string, object> Raw = new Dictionary<string, object>();

        public bool IsClosed
        {
            get { return Deleted || Status == "done" || Status == "dropped"; }
        }

        public static WtTask FromPlain(string id, Dictionary<string, object> d)
        {
            var t = new WtTask
            {
                Id = id, Raw = d,
                Title = Str(d, "title") ?? "(untitled)", Notes = Str(d, "notes"), Project = Str(d, "project"), Person = Str(d, "person"),
                Status = Str(d, "status") ?? "open", DueDate = Date(d, "dueDate"), DueTime = Str(d, "dueTime"),
                TodayDate = Date(d, "todayDate"), TodayRank = (int?)Long(d, "todayRank"), WaitingOn = Str(d, "waitingOn"),
                FollowUpDate = Date(d, "followUpDate"), SnoozedUntil = Date(d, "snoozedUntil"), SeriesId = Str(d, "seriesId"),
                Blocked = Bool(d, "blocked"), BlockedReason = Str(d, "blockedReason"), AgendaFor = Str(d, "agendaFor"), Source = Str(d, "source"),
                CreatedAt = Long(d, "createdAt") ?? 0, UpdatedAt = Long(d, "updatedAt") ?? 0, CompletedAt = Long(d, "completedAt"),
                Deleted = Bool(d, "deleted")
            };
            t.LastTouchedAt = Long(d, "lastTouchedAt") ?? t.CreatedAt;
            var r = Get(d, "recurrence") as Dictionary<string, object>;
            if (r != null && Str(r, "freq") != null)
            {
                t.Recurrence = new WtRecurrence { Freq = Str(r, "freq"), Interval = (int)(Long(r, "interval") ?? 1), DayOfMonth = (int?)Long(r, "dayOfMonth") };
                var days = Get(r, "byDay") as System.Collections.IList;
                if (days != null) foreach (object x in days) t.Recurrence.ByDay.Add(Convert.ToInt32(x, CultureInfo.InvariantCulture));
            }
            return t;
        }

        internal static object Get(Dictionary<string, object> d, string k)
        {
            object v;
            return d != null && d.TryGetValue(k, out v) ? v : null;
        }

        internal static string Str(Dictionary<string, object> d, string k)
        {
            var s = Get(d, k) as string;
            return string.IsNullOrEmpty(s) ? null : s;
        }

        internal static long? Long(Dictionary<string, object> d, string k)
        {
            object v = Get(d, k);
            if (v == null || v is string || v is bool) return null;
            try { return Convert.ToInt64(v, CultureInfo.InvariantCulture); } catch { return null; }
        }

        internal static bool Bool(Dictionary<string, object> d, string k)
        {
            return Get(d, k) is bool && (bool)Get(d, k);
        }

        internal static DateTime? Date(Dictionary<string, object> d, string k)
        {
            DateTime dt;
            string s = Str(d, k);
            return s != null && DateTime.TryParseExact(s, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out dt) ? dt : (DateTime?)null;
        }
    }

    public class WtPrep
    {
        public int DayOffset;
        public string Time;
    }

    public class WtMeeting
    {
        public string Id, Name, Person, Time;
        public int Weekday;
        public List<WtPrep> Prep = new List<WtPrep>();
    }

    public class WtSettings
    {
        public string Timezone = "Asia/Riyadh", WorkStart = "08:00", WorkEnd = "17:00", EveningReview = "16:45", RecapTime = "16:00", PlanTime = "08:00";
        public List<int> WorkDays = new List<int> { 0, 1, 2, 3, 4 };
        public int NagHours = 2, StaleDays = 3, TodayCap = 3, RecapWeekday = 4, PlanWeekday = 0;
        public List<WtMeeting> Meetings = new List<WtMeeting>
        {
            new WtMeeting { Id = "boss", Name = "Boss 1:1", Person = "Boss", Weekday = 3, Time = "13:00",
                Prep = new List<WtPrep> { new WtPrep { DayOffset = -1, Time = "16:00" }, new WtPrep { DayOffset = 0, Time = "11:00" } } }
        };

        public static WtSettings FromPlain(Dictionary<string, object> d)
        {
            var s = new WtSettings();
            if (d == null) return s;
            s.Timezone = WtTask.Str(d, "timezone") ?? s.Timezone;
            s.WorkStart = WtTask.Str(d, "workStart") ?? s.WorkStart;
            s.WorkEnd = WtTask.Str(d, "workEnd") ?? s.WorkEnd;
            s.EveningReview = WtTask.Str(d, "eveningReview") ?? s.EveningReview;
            s.NagHours = (int)(WtTask.Long(d, "nagHours") ?? s.NagHours);
            s.StaleDays = (int)(WtTask.Long(d, "staleDays") ?? s.StaleDays);
            s.TodayCap = (int)(WtTask.Long(d, "todayCap") ?? s.TodayCap);
            var wd = WtTask.Get(d, "workDays") as System.Collections.IList;
            if (wd != null) { s.WorkDays = new List<int>(); foreach (object x in wd) s.WorkDays.Add(Convert.ToInt32(x, CultureInfo.InvariantCulture)); }
            var recap = WtTask.Get(d, "weeklyRecap") as Dictionary<string, object>;
            if (recap != null) { s.RecapWeekday = (int)(WtTask.Long(recap, "weekday") ?? s.RecapWeekday); s.RecapTime = WtTask.Str(recap, "time") ?? s.RecapTime; }
            var plan = WtTask.Get(d, "weekPlan") as Dictionary<string, object>;
            if (plan != null) { s.PlanWeekday = (int)(WtTask.Long(plan, "weekday") ?? s.PlanWeekday); s.PlanTime = WtTask.Str(plan, "time") ?? s.PlanTime; }
            var meetings = WtTask.Get(d, "meetings") as System.Collections.IList;
            if (meetings != null)
            {
                s.Meetings = new List<WtMeeting>();
                foreach (object o in meetings)
                {
                    var m = o as Dictionary<string, object>;
                    if (m == null || WtTask.Str(m, "id") == null) continue;
                    var mt = new WtMeeting
                    {
                        Id = WtTask.Str(m, "id"), Name = WtTask.Str(m, "name") ?? "Meeting", Person = WtTask.Str(m, "person") ?? "",
                        Weekday = (int)(WtTask.Long(m, "weekday") ?? 0), Time = WtTask.Str(m, "time") ?? "10:00"
                    };
                    var prep = WtTask.Get(m, "prep") as System.Collections.IList;
                    if (prep != null)
                        foreach (object p in prep)
                        {
                            var pm = p as Dictionary<string, object>;
                            if (pm != null) mt.Prep.Add(new WtPrep { DayOffset = (int)(WtTask.Long(pm, "dayOffset") ?? 0), Time = WtTask.Str(pm, "time") ?? "09:00" });
                        }
                    s.Meetings.Add(mt);
                }
            }
            return s;
        }
    }

    public class WtDay
    {
        public bool MorningPlanDone, EveningReviewDone;
    }

    public class WtCalEvent
    {
        public string Id, Title;
        public long Start, End;
        public int Attendees;
    }

    // ---------- time ----------

    public static class WtTime
    {
        // Work Track stores IANA names; .NET Framework on Windows wants Windows ids.
        private static readonly Dictionary<string, string> Windows = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Asia/Riyadh", "Arab Standard Time" }, { "Asia/Kuwait", "Arab Standard Time" }, { "Asia/Qatar", "Arab Standard Time" },
            { "Asia/Bahrain", "Arab Standard Time" }, { "Asia/Baghdad", "Arabic Standard Time" }, { "Asia/Dubai", "Arabian Standard Time" },
            { "Asia/Muscat", "Arabian Standard Time" }, { "Asia/Kolkata", "India Standard Time" }, { "Asia/Karachi", "Pakistan Standard Time" },
            { "Africa/Cairo", "Egypt Standard Time" }, { "Europe/London", "GMT Standard Time" }, { "Europe/Berlin", "W. Europe Standard Time" },
            { "Europe/Paris", "Romance Standard Time" }, { "America/New_York", "Eastern Standard Time" }, { "UTC", "UTC" }
        };

        public static TimeZoneInfo Resolve(string iana)
        {
            if (!string.IsNullOrEmpty(iana))
            {
                try { return TimeZoneInfo.FindSystemTimeZoneById(iana); } catch { }
                string win;
                if (Windows.TryGetValue(iana, out win)) { try { return TimeZoneInfo.FindSystemTimeZoneById(win); } catch { } }
            }
            return TimeZoneInfo.Local;
        }

        private static readonly DateTime Epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        public static int Minutes(string hhmm)
        {
            var p = hhmm.Split(':');
            return int.Parse(p[0], CultureInfo.InvariantCulture) * 60 + int.Parse(p[1], CultureInfo.InvariantCulture);
        }

        public static string Time(int minutes)
        {
            return (minutes / 60).ToString("00") + ":" + (minutes % 60).ToString("00");
        }

        public static long ToEpoch(DateTime date, string hhmm, TimeZoneInfo tz)
        {
            var local = DateTime.SpecifyKind(date.Date.AddMinutes(Minutes(hhmm)), DateTimeKind.Unspecified);
            return (long)(TimeZoneInfo.ConvertTimeToUtc(local, tz) - Epoch).TotalMilliseconds;
        }

        public static DateTime Local(long ms, TimeZoneInfo tz)
        {
            return TimeZoneInfo.ConvertTimeFromUtc(Epoch.AddMilliseconds(ms), tz);
        }

        public static DateTime DateOf(long ms, TimeZoneInfo tz) { return Local(ms, tz).Date; }

        public static int MinutesOf(long ms, TimeZoneInfo tz)
        {
            var l = Local(ms, tz);
            return l.Hour * 60 + l.Minute;
        }

        public static int Wd(DateTime d) { return (int)d.DayOfWeek; }   // Sunday = 0, same as the app

        public static DateTime WeekStart(DateTime d) { return d.Date.AddDays(-Wd(d)); }

        public static string Iso(DateTime d) { return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); }
    }

    public class WtCtx
    {
        public DateTime Today;
        public WtSettings Settings;
        public TimeZoneInfo Tz;

        public WtCtx(DateTime today, WtSettings settings, TimeZoneInfo tz)
        {
            Today = today.Date;
            Settings = settings;
            Tz = tz;
        }

        public DateTime DateOf(long ms) { return WtTime.DateOf(ms, Tz); }
    }

    // ---------- rules ----------

    public static class WtRules
    {
        public static bool IsWorkDay(DateTime d, IList<int> workDays) { return workDays.Contains(WtTime.Wd(d)); }

        /// <summary>Number of work days d with a &lt; d &lt;= b.</summary>
        public static int WorkDaysBetween(DateTime a, DateTime b, IList<int> workDays)
        {
            int n = 0;
            for (var d = a.Date.AddDays(1); d <= b.Date; d = d.AddDays(1)) if (IsWorkDay(d, workDays)) n++;
            return n;
        }

        public static DateTime AddWorkDays(DateTime d, int n, IList<int> workDays)
        {
            if (workDays.Count == 0) return d.Date.AddDays(n);
            var cur = d.Date;
            while (n > 0)
            {
                cur = cur.AddDays(1);
                if (IsWorkDay(cur, workDays)) n--;
            }
            return cur;
        }

        private static int Mod(long x, int m) { return (int)(((x % m) + m) % m); }

        private static int MonthIndex(DateTime d) { return d.Year * 12 + d.Month - 1; }

        private static bool Matches(WtRecurrence r, DateTime d, DateTime anchor, IList<int> workDays)
        {
            int interval = Math.Max(1, r.Interval);
            switch (r.Freq)
            {
                case "daily":
                    return Mod((long)(d - anchor).TotalDays, interval) == 0;
                case "workdays":
                    return IsWorkDay(d, workDays.Count > 0 ? workDays : new List<int> { 0, 1, 2, 3, 4, 5, 6 });
                case "weekly":
                    var days = r.ByDay.Count > 0 ? r.ByDay : new List<int> { WtTime.Wd(anchor) };
                    if (!days.Contains(WtTime.Wd(d))) return false;
                    return Mod((long)((WtTime.WeekStart(d) - WtTime.WeekStart(anchor)).TotalDays / 7), interval) == 0;
                case "monthly":
                    int want = Math.Min(r.DayOfMonth ?? anchor.Day, DateTime.DaysInMonth(d.Year, d.Month));
                    return d.Day == want && Mod(MonthIndex(d) - MonthIndex(anchor), interval) == 0;
            }
            return false;
        }

        public static DateTime NextOccurrence(WtRecurrence r, DateTime after, DateTime anchor, IList<int> workDays)
        {
            if (r.Freq == "monthly")
            {
                int interval = Math.Max(1, r.Interval);
                var first = new DateTime(after.Year, after.Month, 1);
                for (int i = 0; i < 12 * interval + 2; i++)
                {
                    var m = first.AddMonths(i);
                    var cand = new DateTime(m.Year, m.Month, Math.Min(r.DayOfMonth ?? anchor.Day, DateTime.DaysInMonth(m.Year, m.Month)));
                    if (cand > after && Matches(r, cand, anchor, workDays)) return cand;
                }
                return after.AddDays(30);
            }
            for (int i = 1; i <= 7 * 60; i++)
            {
                var cand = after.Date.AddDays(i);
                if (Matches(r, cand, anchor, workDays)) return cand;
            }
            return after.Date.AddDays(1);
        }

        /// <summary>Next instance of a recurring task finished on completedDate: (doc id, due date), or null.</summary>
        public static KeyValuePair<string, DateTime>? NextInstance(WtTask t, DateTime completedDate, IList<int> workDays)
        {
            if (t.Recurrence == null) return null;
            var due = t.DueDate ?? completedDate.Date;
            var baseDate = due > completedDate.Date ? due : completedDate.Date;
            var next = NextOccurrence(t.Recurrence, baseDate, due, workDays);
            string series = string.IsNullOrEmpty(t.SeriesId) ? t.Id : t.SeriesId;
            return new KeyValuePair<string, DateTime>(series + "_" + WtTime.Iso(next), next);
        }

        public static bool IsStale(string status, DateTime lastTouched, DateTime? due, DateTime? snoozedUntil, DateTime? todayDate, bool deleted, DateTime today, int staleDays, IList<int> workDays)
        {
            if (deleted) return false;
            if (status != "open" && status != "inbox") return false;
            if (snoozedUntil.HasValue && snoozedUntil.Value > today) return false;
            if (todayDate.HasValue && todayDate.Value == today) return false;
            if (due.HasValue && due.Value >= today) return false;
            if (due.HasValue && WorkDaysBetween(due.Value, today, workDays) >= staleDays) return true;
            return WorkDaysBetween(lastTouched, today, workDays) >= staleDays;
        }

        private static IEnumerable<WtTask> Live(IEnumerable<WtTask> tasks) { return tasks.Where(t => !t.Deleted); }

        private static IEnumerable<WtTask> ByDue(IEnumerable<WtTask> tasks)
        {
            return tasks.OrderBy(t => t.DueDate ?? DateTime.MaxValue).ThenBy(t => t.DueTime ?? "99").ThenBy(t => t.CreatedAt);
        }

        public static List<WtTask> OpenToday(IEnumerable<WtTask> tasks, WtCtx c)
        {
            return Live(tasks).Where(t => t.TodayDate == c.Today && t.Status != "dropped" && !t.IsClosed)
                .OrderBy(t => t.TodayRank ?? 99).ThenBy(t => t.CreatedAt).ToList();
        }

        public static List<WtTask> Stale(IEnumerable<WtTask> tasks, WtCtx c)
        {
            return Live(tasks).Where(t => IsStale(t.Status, c.DateOf(t.LastTouchedAt), t.DueDate, t.SnoozedUntil, t.TodayDate, t.Deleted, c.Today, c.Settings.StaleDays, c.Settings.WorkDays)).ToList();
        }

        public static List<WtTask> Inbox(IEnumerable<WtTask> tasks) { return Live(tasks).Where(t => t.Status == "inbox").ToList(); }

        public static List<WtTask> DueNow(IEnumerable<WtTask> tasks, WtCtx c)
        {
            return ByDue(Live(tasks).Where(t => (t.Status == "open" || t.Status == "inbox") && t.DueDate.HasValue && t.DueDate.Value <= c.Today
                && !(t.SnoozedUntil.HasValue && t.SnoozedUntil.Value > c.Today))).ToList();
        }

        public static List<WtTask> Overdue(IEnumerable<WtTask> tasks, WtCtx c) { return DueNow(tasks, c).Where(t => t.DueDate.Value < c.Today).ToList(); }

        public static List<WtTask> FollowUpsDue(IEnumerable<WtTask> tasks, WtCtx c)
        {
            return Live(tasks).Where(t => t.Status == "waiting" && t.FollowUpDate.HasValue && t.FollowUpDate.Value <= c.Today)
                .OrderBy(t => t.FollowUpDate).ToList();
        }

        public static List<WtTask> DoneBetween(IEnumerable<WtTask> tasks, long fromMs, long toMs)
        {
            return Live(tasks).Where(t => t.Status == "done" && t.CompletedAt.HasValue && t.CompletedAt.Value >= fromMs && t.CompletedAt.Value < toMs).ToList();
        }

        public static bool SameName(string a, string b)
        {
            return !string.IsNullOrWhiteSpace(a) && !string.IsNullOrWhiteSpace(b) && string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        public static List<WtTask> MeetingAsks(IEnumerable<WtTask> tasks, WtMeeting m) { return Live(tasks).Where(t => SameName(t.Person, m.Person)).ToList(); }

        public static bool IsAgendaFor(WtTask t, WtMeeting m)
        {
            string a = (t.AgendaFor ?? "").ToLowerInvariant();
            return a.Length > 0 && (a == m.Id.ToLowerInvariant() || a == m.Person.ToLowerInvariant());
        }

        public static List<WtTask> MeetingBlockers(IEnumerable<WtTask> tasks, WtMeeting m)
        {
            return tasks.Where(t => !t.IsClosed && (t.Blocked || IsAgendaFor(t, m))).ToList();
        }
    }

    // ---------- schedule ----------

    public class WtSlot
    {
        public string Key, Kind, MeetingId, TaskId, Label;
        public long At;
        public DateTime Date;
        public int? DayOffset;
    }

    public class WtContent
    {
        public string Title, Body, Route;
        public List<string> TaskIds = new List<string>();
    }

    public static class WtSchedule
    {
        public static readonly string[] CaptureTimes = { "11:00", "15:00" };

        public static List<WtSlot> Slots(long fromMs, long toMs, WtSettings s, TimeZoneInfo tz, IEnumerable<WtTask> tasks, bool calendarAvailable)
        {
            var output = new List<WtSlot>();
            var startDate = WtTime.DateOf(fromMs, tz);
            var endDate = WtTime.DateOf(toMs, tz);
            Action<DateTime, string, string, string, int?, string> push = (date, time, kind, meetingId, dayOffset, taskId) =>
            {
                long at = WtTime.ToEpoch(date, time, tz);
                if (at < fromMs || at > toMs) return;
                output.Add(new WtSlot
                {
                    Key = string.Join("|", new[] { kind, WtTime.Iso(date), time, meetingId ?? "", taskId ?? "" }),
                    At = at, Date = date, Kind = kind, MeetingId = meetingId, DayOffset = dayOffset, TaskId = taskId
                });
            };
            int start = WtTime.Minutes(s.WorkStart), end = WtTime.Minutes(s.WorkEnd);
            for (var d = startDate; d <= endDate; d = d.AddDays(1))
            {
                if (WtRules.IsWorkDay(d, s.WorkDays))
                {
                    push(d, s.WorkStart, "morning", null, null, null);
                    for (int m = start + 60; m < end; m += 60) push(d, WtTime.Time(m), "morning-repeat", null, null, null);
                    int step = Math.Max(1, s.NagHours) * 60;
                    for (int m = start + step; m < end; m += step) push(d, WtTime.Time(m), "nag", null, null, null);
                    if (start + 90 < end) push(d, WtTime.Time(start + 90), "followups", null, null, null);
                    if (!calendarAvailable) foreach (string t in CaptureTimes) push(d, t, "capture", null, null, null);
                    push(d, s.EveningReview, "evening", null, null, null);
                    if (WtTime.Wd(d) == s.RecapWeekday) push(d, s.RecapTime, "weekly-recap", null, null, null);
                }
                foreach (var mt in s.Meetings)
                {
                    foreach (var p in mt.Prep)
                        if (WtTime.Wd(d.AddDays(-p.DayOffset)) == mt.Weekday) push(d, p.Time, "meeting-prep", mt.Id, p.DayOffset, null);
                    if (WtTime.Wd(d) == mt.Weekday) push(d, mt.Time, "meeting-start", mt.Id, null, null);
                }
            }
            foreach (var t in tasks)
            {
                if (t.IsClosed || !t.DueDate.HasValue || t.DueTime == null || t.Status == "someday" || t.Status == "waiting") continue;
                if (t.DueDate.Value < startDate || t.DueDate.Value > endDate) continue;
                push(t.DueDate.Value, t.DueTime, "due-time", null, null, t.Id);
            }
            return output.OrderBy(x => x.At).ThenBy(x => x.Key, StringComparer.Ordinal).ToList();
        }

        /// <summary>"Any action items?" 2 minutes after real meetings (2+ attendees) ending in work hours.</summary>
        public static List<WtSlot> MeetingEndSlots(IEnumerable<WtCalEvent> events, WtSettings s, TimeZoneInfo tz, long fromMs, long toMs)
        {
            var seen = new HashSet<string>();
            var output = new List<WtSlot>();
            foreach (var e in events)
            {
                long at = e.End + 2 * 60000;
                var date = WtTime.DateOf(e.End, tz);
                int min = WtTime.MinutesOf(e.End, tz);
                if (e.Attendees < 2 || at < fromMs || at > toMs || !WtRules.IsWorkDay(date, s.WorkDays)) continue;
                if (min < WtTime.Minutes(s.WorkStart) || min > WtTime.Minutes(s.WorkEnd) + 30) continue;
                if (!seen.Add((e.Title ?? "").Trim().ToLowerInvariant() + "|" + e.End)) continue;
                output.Add(new WtSlot { Key = "meeting-end|" + e.Id + "|" + e.End, At = at, Date = date, Kind = "meeting-end", Label = e.Title });
            }
            return output;
        }

        private static string Plural(int n, string word) { return n + " " + word + (n == 1 ? "" : "s"); }

        private static string List(IList<WtTask> tasks, int max)
        {
            var lines = tasks.Take(max).Select(t => "• " + t.Title).ToList();
            if (tasks.Count > max) lines.Add("…and " + (tasks.Count - max) + " more");
            return string.Join("\n", lines.ToArray());
        }

        public static WtContent Content(WtSlot slot, IList<WtTask> tasks, WtCtx c, WtDay day)
        {
            var s = c.Settings;
            switch (slot.Kind)
            {
                case "morning":
                case "morning-repeat":
                {
                    if (day.MorningPlanDone) return null;
                    bool weekPlan = WtTime.Wd(c.Today) == s.PlanWeekday;
                    string body = string.Join(" · ", new[]
                    {
                        Plural(WtRules.Overdue(tasks, c).Count, "overdue"),
                        WtRules.DueNow(tasks, c).Count(t => t.DueDate == c.Today) + " due today",
                        WtRules.Stale(tasks, c).Count + " to decide",
                        WtRules.Inbox(tasks).Count + " in inbox"
                    });
                    return new WtContent { Title = weekPlan ? "Plan your week — and pick today’s 3" : "Pick today’s " + s.TodayCap, Body = body, Route = "plan" };
                }
                case "nag":
                {
                    var due = WtRules.DueNow(tasks, c);
                    if (due.Count == 0) return null;
                    int over = due.Count(t => t.DueDate.Value < c.Today);
                    var content = new WtContent
                    {
                        Title = over > 0 ? Plural(over, "overdue task") + ", " + (due.Count - over) + " due today" : Plural(due.Count, "task") + " due today",
                        Body = List(due, 4), Route = "today"
                    };
                    content.TaskIds.AddRange(due.Select(t => t.Id));
                    return content;
                }
                case "followups":
                {
                    var f = WtRules.FollowUpsDue(tasks, c);
                    if (f.Count == 0) return null;
                    var content = new WtContent
                    {
                        Title = "Time to chase " + Plural(f.Count, "person"),
                        Body = string.Join("\n", f.Take(4).Select(t => "• " + t.WaitingOn + ": " + t.Title).ToArray()),
                        Route = "waiting"
                    };
                    content.TaskIds.AddRange(f.Select(t => t.Id));
                    return content;
                }
                case "capture":
                    return new WtContent { Title = "Anything new to capture?", Body = "Asks from meetings, chats or email since the last check.", Route = "add" };
                case "evening":
                {
                    if (day.EveningReviewDone) return null;
                    long dayStart = WtTime.ToEpoch(c.Today, "00:00", c.Tz);
                    int done = WtRules.DoneBetween(tasks, dayStart, dayStart + 86400000L).Count;
                    int left = WtRules.OpenToday(tasks, c).Count;
                    return new WtContent { Title = "Evening review (2 min)", Body = done + " done today · " + left + " of today’s left · " + WtRules.Inbox(tasks).Count + " in inbox", Route = "review" };
                }
                case "weekly-recap":
                    return new WtContent { Title = "Weekly recap is ready", Body = "What got done, what slipped, and next week’s load.", Route = "weekly" };
                case "meeting-prep":
                case "meeting-start":
                {
                    var mt = s.Meetings.FirstOrDefault(m => m.Id == slot.MeetingId);
                    if (mt == null) return null;
                    if (slot.Kind == "meeting-start")
                        return new WtContent { Title = mt.Name + " now", Body = "Open meeting mode to jot down new asks.", Route = "meeting/" + mt.Id };
                    int asks = WtRules.MeetingAsks(tasks, mt).Count(t => !t.IsClosed);
                    int blockers = WtRules.MeetingBlockers(tasks, mt).Count;
                    string when = (slot.DayOffset ?? 0) < 0 ? "tomorrow" : "today at " + mt.Time;
                    return new WtContent
                    {
                        Title = mt.Name + " " + when,
                        Body = Plural(asks, "open ask") + " from " + mt.Person + " · " + Plural(blockers, "blocker") + " · tap for your brief",
                        Route = "brief/" + mt.Id
                    };
                }
                case "due-time":
                {
                    var t = tasks.FirstOrDefault(x => x.Id == slot.TaskId);
                    if (t == null || t.IsClosed || t.Status == "waiting" || t.Status == "someday" || t.DueDate != slot.Date) return null;
                    var content = new WtContent { Title = "Due now: " + t.Title, Body = t.Notes ?? "", Route = "today" };
                    content.TaskIds.Add(t.Id);
                    return content;
                }
                case "meeting-end":
                    return new WtContent { Title = "“" + (slot.Label ?? "Meeting") + "” just ended", Body = "Any action items? Add them in Work Track.", Route = "add" };
            }
            return null;
        }
    }
}
