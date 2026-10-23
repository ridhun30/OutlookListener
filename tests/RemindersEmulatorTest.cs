using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using OutlookListener;

// Integration test for the desktop-reminder data path against the Firebase emulators:
// full read, incremental read by updatedAt, settings with nested meetings, and "Done" from the
// tray writing the task + its next recurring instance (deterministic id), as the phone does.
internal static class RemindersEmulatorTest
{
    private static int Main()
    {
        WorkTrackConnector.AuthBase = "http://127.0.0.1:9099/securetoken.googleapis.com";
        WorkTrackConnector.FirestoreBase = "http://127.0.0.1:8080";
        var json = new JavaScriptSerializer();
        string idToken = "{\"sub\":\"tester@example.com\",\"email\":\"tester@example.com\",\"email_verified\":true}";
        var signIn = json.Deserialize<Dictionary<string, object>>(Post(
            "http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signInWithIdp?key=demo-key",
            json.Serialize(new Dictionary<string, object> { { "postBody", "id_token=" + Uri.EscapeDataString(idToken) + "&providerId=google.com" }, { "requestUri", "http://localhost" }, { "returnSecureToken", true } })));
        var conn = new WorkTrackConnector(new ConnectorCode { ApiKey = "demo-key", ProjectId = "demo-worktrack", Uid = (string)signIn["localId"], RefreshToken = (string)signIn["refreshToken"] });

        var tz = WtTime.Resolve("Asia/Riyadh");
        var today = WtTime.DateOf(WorkTrackConnector.NowMs(), tz);
        conn.SetDocument("settings/main", new Dictionary<string, object>
        {
            { "timezone", "Asia/Riyadh" }, { "workDays", new List<object> { 0, 1, 2, 3, 4 } }, { "todayCap", 3 },
            { "meetings", new List<object> { new Dictionary<string, object> { { "id", "boss" }, { "name", "Boss 1:1" }, { "person", "Boss" }, { "weekday", 3 }, { "time", "13:00" },
                { "prep", new List<object> { new Dictionary<string, object> { { "dayOffset", -1 }, { "time", "16:00" } } } } } } }
        });
        var settings = WtSettings.FromPlain(conn.GetDocument("settings/main"));
        Assert(settings.Meetings.Count == 1 && settings.Meetings[0].Prep[0].DayOffset == -1 && settings.WorkDays.Count == 5, "settings with nested meetings read back");
        Assert(conn.GetDocument("days/1999-01-01") == null, "missing document reads as null");

        long before = WorkTrackConnector.NowMs() - 1;
        var task = WorkTrackConnector.NewTask("Weekly ops report", "", null, "open", WtTime.Iso(today), null, false);
        task["source"] = "web";
        task["todayDate"] = WtTime.Iso(today);
        task["todayRank"] = 1;
        task["recurrence"] = new Dictionary<string, object> { { "freq", "weekly" }, { "interval", 1 }, { "byDay", new List<object> { WtTime.Wd(today) } }, { "dayOfMonth", null } };
        Assert(conn.CreateTask("rec-1", task), "recurring task created");

        var all = conn.AllTasks();
        Assert(all.Any(r => (string)r["id"] == "rec-1"), "full read finds it");
        var changed = conn.TasksUpdatedSince(before);
        Assert(changed.Count >= 1 && changed.All(r => (long)r["updatedAt"] > before), "incremental read returns only newer tasks");
        Assert(conn.TasksUpdatedSince(WorkTrackConnector.NowMs() + 60000).Count == 0, "incremental read is empty when nothing changed");

        var row = changed.First(r => (string)r["id"] == "rec-1");
        row.Remove("id");
        var t = WtTask.FromPlain("rec-1", row);
        Assert(t.Recurrence != null && t.Recurrence.Freq == "weekly" && t.TodayDate == today, "task parsed with recurrence and Today");
        var ctx = new WtCtx(today, settings, tz);
        Assert(WtRules.OpenToday(new[] { t }, ctx).Count == 1, "it is today's #1");

        // What "Done" in the tray writes:
        var next = WtRules.NextInstance(t, today, settings.WorkDays).Value;
        long now = WorkTrackConnector.NowMs();
        conn.UpdateTask("rec-1", new Dictionary<string, object> { { "status", "done" }, { "completedAt", now }, { "updatedAt", now }, { "lastTouchedAt", now } });
        var nextFields = new Dictionary<string, object>(t.Raw);
        nextFields["seriesId"] = "rec-1"; nextFields["status"] = "open"; nextFields["dueDate"] = WtTime.Iso(next.Value);
        nextFields["todayDate"] = null; nextFields["todayRank"] = null; nextFields["completedAt"] = null;
        nextFields["createdAt"] = now; nextFields["updatedAt"] = now; nextFields["lastTouchedAt"] = now;
        conn.SetDocument("tasks/" + next.Key, nextFields);
        var after = conn.AllTasks();
        Assert(after.Any(r => (string)r["id"] == "rec-1" && (string)r["status"] == "done"), "task is done");
        Assert(after.Any(r => (string)r["id"] == next.Key && (string)r["dueDate"] == WtTime.Iso(next.Value) && r["recurrence"] is Dictionary<string, object>), "next instance " + next.Key + " exists with its recurrence");

        Console.WriteLine("Reminders emulator test passed. next=" + next.Key);
        return 0;
    }

    private static void Assert(bool ok, string what)
    {
        Console.WriteLine((ok ? "ok   " : "FAIL ") + what);
        if (!ok) Environment.Exit(1);
    }

    private static string Post(string url, string body)
    {
        var req = (HttpWebRequest)WebRequest.Create(url);
        req.Method = "POST"; req.ContentType = "application/json"; req.Proxy = null;
        byte[] b = Encoding.UTF8.GetBytes(body);
        using (var s = req.GetRequestStream()) s.Write(b, 0, b.Length);
        using (var res = req.GetResponse()) using (var r = new StreamReader(res.GetResponseStream())) return r.ReadToEnd();
    }
}
