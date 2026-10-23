using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using OutlookListener;

// Tests for WorkTrackRules.cs (the C# port of Work Track's reminder rules). Runs the shared
// vectors from the Work Track repo plus mirrors of its schedule tests (web/tests/schedule.test.ts).
//   csc /out:WorkTrackRulesTests.exe /r:System.Web.Extensions.dll tests\WorkTrackRulesTests.cs WorkTrackRules.cs
//   WorkTrackRulesTests.exe tests\vectors\rules-cases.json
internal static class WorkTrackRulesTests
{
    private static int _failed, _passed;

    private static void Eq<T>(string name, T actual, T expected)
    {
        if (Equals(actual, expected)) { _passed++; return; }
        _failed++;
        Console.WriteLine("FAIL " + name + ": got " + actual + ", expected " + expected);
    }

    private static DateTime D(string s) { return DateTime.ParseExact(s, "yyyy-MM-dd", null); }
    private static List<int> Ints(object o) { return ((IEnumerable)o).Cast<object>().Select(Convert.ToInt32).ToList(); }

    private static int Main(string[] args)
    {
        string path = args.Length > 0 ? args[0] : Path.Combine("tests", Path.Combine("vectors", "rules-cases.json"));
        var v = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(File.ReadAllText(path));
        var wd = Ints(v["workDays"]);

        foreach (Dictionary<string, object> c in (IEnumerable)v["workDaysBetween"])
            Eq("workDaysBetween " + c["a"] + " " + c["b"], WtRules.WorkDaysBetween(D((string)c["a"]), D((string)c["b"]), wd), Convert.ToInt32(c["expect"]));
        foreach (Dictionary<string, object> c in (IEnumerable)v["addWorkDays"])
            Eq("addWorkDays " + c["date"], WtRules.AddWorkDays(D((string)c["date"]), Convert.ToInt32(c["n"]), wd), D((string)c["expect"]));
        foreach (Dictionary<string, object> c in (IEnumerable)v["nextOccurrence"])
        {
            var r = (Dictionary<string, object>)c["rec"];
            var rec = new WtRecurrence { Freq = (string)r["freq"], Interval = Convert.ToInt32(r["interval"]), ByDay = Ints(r["byDay"]), DayOfMonth = r["dayOfMonth"] == null ? (int?)null : Convert.ToInt32(r["dayOfMonth"]) };
            Eq("nextOccurrence " + r["freq"] + " after " + c["after"], WtRules.NextOccurrence(rec, D((string)c["after"]), D((string)c["anchor"]), wd), D((string)c["expect"]));
        }
        var st = (Dictionary<string, object>)v["isStale"];
        foreach (Dictionary<string, object> c in (IEnumerable)st["cases"])
        {
            var t = (Dictionary<string, object>)c["task"];
            Func<string, DateTime?> opt = k => t.ContainsKey(k) && t[k] != null ? D((string)t[k]) : (DateTime?)null;
            Eq("isStale: " + c["name"], WtRules.IsStale((string)t["status"], D((string)t["lastTouchedDate"]), opt("dueDate"), opt("snoozedUntil"), opt("todayDate"), false, D((string)st["today"]), Convert.ToInt32(st["staleDays"]), wd), (bool)c["expect"]);
        }

        // ---- schedule (mirrors web/tests/schedule.test.ts) ----
        var s = new WtSettings();
        var tz = WtTime.Resolve(s.Timezone);
        Func<string, string, long> at = (d, t) => WtTime.ToEpoch(D(d), t, tz);
        Eq("Riyadh is UTC+3", at("2026-10-28", "13:00"), (long)(new DateTime(2026, 10, 28, 10, 0, 0, DateTimeKind.Utc) - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds);
        Eq("date of late UTC evening", WtTime.DateOf(at("2026-10-29", "00:30"), tz), D("2026-10-29"));

        var empty = new List<WtTask>();
        var sunday = WtSchedule.Slots(at("2026-10-25", "00:00"), at("2026-10-25", "23:59"), s, tz, empty, true);
        Eq("nags at 10,12,14,16", string.Join(",", sunday.Where(x => x.Kind == "nag").Select(x => (WtTime.MinutesOf(x.At, tz) / 60).ToString()).ToArray()), "10,12,14,16");
        Eq("hourly morning repeats", sunday.Count(x => x.Kind == "morning-repeat"), 8);
        Eq("followups 09:30", sunday.Count(x => x.Kind == "followups" && WtTime.MinutesOf(x.At, tz) == 570), 1);
        Eq("evening 16:45", sunday.Count(x => x.Kind == "evening" && WtTime.MinutesOf(x.At, tz) == 1005), 1);
        Eq("no capture with calendar", sunday.Count(x => x.Kind == "capture"), 0);
        Eq("weekend is quiet", WtSchedule.Slots(at("2026-10-30", "00:00"), at("2026-10-31", "23:59"), s, tz, empty, true).Count, 0);
        var boss = WtSchedule.Slots(at("2026-10-25", "00:00"), at("2026-10-31", "23:59"), s, tz, empty, true).Where(x => x.MeetingId == "boss").ToList();
        Eq("boss prep/start", string.Join(";", boss.Select(x => x.Kind + "@" + WtTime.Iso(x.Date) + " " + WtTime.MinutesOf(x.At, tz)).ToArray()),
            "meeting-prep@2026-10-27 960;meeting-prep@2026-10-28 660;meeting-start@2026-10-28 780");
        Eq("weekly recap Thu 16:00", WtSchedule.Slots(at("2026-10-25", "00:00"), at("2026-10-31", "23:59"), s, tz, empty, true).Count(x => x.Kind == "weekly-recap" && WtTime.Iso(x.Date) == "2026-10-29"), 1);

        long t0 = at("2026-10-20", "09:00");
        int n = 0;
        Func<string, string, WtTask> task = (due, person) => new WtTask { Id = "t" + (++n), Title = "task " + n, DueDate = due == null ? (DateTime?)null : D(due), Person = person, CreatedAt = t0, UpdatedAt = t0, LastTouchedAt = t0 };
        var ctx = new WtCtx(D("2026-10-27"), s, tz);
        var nag = new WtSlot { Kind = "nag", Date = D("2026-10-27") };
        Eq("nag silent when nothing due", WtSchedule.Content(nag, new List<WtTask> { task(null, null) }, ctx, new WtDay()) == null, true);
        Eq("nag title", WtSchedule.Content(nag, new List<WtTask> { task("2026-10-26", null), task("2026-10-27", null) }, ctx, new WtDay()).Title, "1 overdue task, 1 due today");
        Eq("morning suppressed once planned", WtSchedule.Content(new WtSlot { Kind = "morning-repeat" }, empty, ctx, new WtDay { MorningPlanDone = true }) == null, true);
        Eq("morning route", WtSchedule.Content(new WtSlot { Kind = "morning" }, empty, ctx, new WtDay()).Route, "plan");
        var prepTasks = new List<WtTask> { task(null, "boss"), task(null, null), task(null, null) };
        prepTasks[1].Blocked = true;
        prepTasks[2].AgendaFor = "boss";
        var prep = WtSchedule.Content(new WtSlot { Kind = "meeting-prep", MeetingId = "boss", DayOffset = -1 }, prepTasks, ctx, new WtDay());
        Eq("prep title", prep.Title, "Boss 1:1 tomorrow");
        Eq("prep body", prep.Body, "1 open ask from Boss · 2 blockers · tap for your brief");

        var events = new List<WtCalEvent>
        {
            new WtCalEvent { Id = "a", Title = "Weekly sync", Start = at("2026-10-27", "10:00"), End = at("2026-10-27", "10:30"), Attendees = 4 },
            new WtCalEvent { Id = "b", Title = "Focus time", Start = at("2026-10-27", "11:00"), End = at("2026-10-27", "12:00"), Attendees = 1 },
            new WtCalEvent { Id = "c", Title = "Late call", Start = at("2026-10-27", "18:00"), End = at("2026-10-27", "19:00"), Attendees = 3 }
        };
        var ends = WtSchedule.MeetingEndSlots(events, s, tz, at("2026-10-27", "00:00"), at("2026-10-28", "00:00"));
        Eq("meeting-end only for real meetings in work hours", ends.Count == 1 && ends[0].Label == "Weekly sync" && ends[0].At == at("2026-10-27", "10:32"), true);

        var recurring = new WtTask { Id = "abc", Title = "weekly report", DueDate = D("2026-10-29"), Recurrence = new WtRecurrence { Freq = "weekly", Interval = 1, ByDay = new List<int> { 4 } } };
        var next = WtRules.NextInstance(recurring, D("2026-10-29"), s.WorkDays).Value;
        Eq("next instance id", next.Key, "abc_2026-11-05");

        var plain = new Dictionary<string, object> { { "workDays", new List<object> { 0L, 1L } }, { "meetings", new List<object> { new Dictionary<string, object> { { "id", "boss" }, { "name", "1:1" }, { "person", "Ahmed" }, { "weekday", 2L }, { "time", "10:00" }, { "prep", new List<object>() } } } } };
        var parsed = WtSettings.FromPlain(plain);
        Eq("settings parse", parsed.WorkDays.Count + "|" + parsed.Meetings[0].Person + "|" + parsed.Meetings[0].Weekday, "2|Ahmed|2");

        Console.WriteLine(_failed == 0 ? "All " + _passed + " WorkTrackRules tests passed." : _failed + " of " + (_passed + _failed) + " tests failed.");
        return _failed == 0 ? 0 : 1;
    }
}
