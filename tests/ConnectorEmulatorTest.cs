using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;
using OutlookListener;

// Integration test for WorkTrackConnector against the Firebase emulators (Auth :9099,
// Firestore :8080, project demo-worktrack). Signs in as the same fake Google user the
// Work Track web e2e tests use, then exercises every REST call the listener makes.
internal static class ConnectorEmulatorTest
{
    private static int Main()
    {
        WorkTrackConnector.AuthBase = "http://127.0.0.1:9099/securetoken.googleapis.com";
        WorkTrackConnector.FirestoreBase = "http://127.0.0.1:8080";

        var json = new JavaScriptSerializer();
        string idToken = "{\"sub\":\"tester@example.com\",\"email\":\"tester@example.com\",\"email_verified\":true}";
        string body = json.Serialize(new Dictionary<string, object>
        {
            { "postBody", "id_token=" + Uri.EscapeDataString(idToken) + "&providerId=google.com" },
            { "requestUri", "http://localhost" }, { "returnSecureToken", true }
        });
        var signIn = json.Deserialize<Dictionary<string, object>>(Post(
            "http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signInWithIdp?key=demo-key", body));

        var code = new ConnectorCode
        {
            ApiKey = "demo-key", ProjectId = "demo-worktrack",
            Uid = (string)signIn["localId"], RefreshToken = (string)signIn["refreshToken"]
        };
        // The code survives a round trip through its text form.
        code = ConnectorCode.Parse(code.ToText());
        var conn = new WorkTrackConnector(code);
        conn.Test();

        string entry = "00000000ABCDEF-test-entry-id";
        string id = WorkTrackConnector.DocIdFor(entry);
        var task = WorkTrackConnector.NewTask("Approve PO 4411", "Outlook mail from Sara", "Sara", "inbox", null, entry, false);
        Assert(conn.CreateTask(id, task), "first create succeeds");
        Assert(!conn.CreateTask(id, task), "second create is a no-op (no duplicate)");

        string flagged = WorkTrackConnector.DocIdFor("flag-entry");
        Assert(conn.CreateTask(flagged, WorkTrackConnector.NewTask("Reply to auditors", "Flagged in Outlook", "Boss", "open", "2026-10-29", "flag-entry", true)), "flagged task created");
        conn.UpdateTask(flagged, new Dictionary<string, object> { { "status", "done" }, { "completedAt", WorkTrackConnector.NowMs() } });
        var finished = conn.FinishedFlaggedTasks();
        Assert(finished.Count == 1 && (string)finished[0]["id"] == flagged && (string)finished[0]["outlookEntryId"] == "flag-entry", "finished flagged task is found");
        conn.UpdateTask(flagged, new Dictionary<string, object> { { "outlookFlagSynced", true } });
        Assert(conn.FinishedFlaggedTasks().Count == 0, "synced task is no longer returned");
        conn.UpdateTask("does-not-exist", new Dictionary<string, object> { { "status", "done" } }); // ignored, no throw

        long start = WorkTrackConnector.ToMs(DateTime.UtcNow.AddHours(1));
        conn.SetDocument("calendar/" + DateTime.Now.ToString("yyyy-MM-dd"), new Dictionary<string, object>
        {
            { "events", new List<object> { new Dictionary<string, object> { { "id", "e1" }, { "title", "Weekly sync" }, { "start", start }, { "end", start + 1800000 }, { "attendees", 4 } } } },
            { "updatedAt", WorkTrackConnector.NowMs() }, { "source", "outlook" }
        });

        // Someone else's code can't write into this user's data.
        var other = new ConnectorCode { ApiKey = "demo-key", ProjectId = "demo-worktrack", Uid = "someone-else", RefreshToken = code.RefreshToken };
        bool denied = false;
        try { new WorkTrackConnector(other).CreateTask("x", task); } catch (WorkTrackException) { denied = true; }
        Assert(denied, "writing to another user's data is denied by firestore.rules");

        Console.WriteLine("Connector emulator test passed. uid=" + code.Uid);
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
        req.Method = "POST";
        req.ContentType = "application/json";
        req.Proxy = null;
        byte[] b = Encoding.UTF8.GetBytes(body);
        using (var s = req.GetRequestStream()) s.Write(b, 0, b.Length);
        using (var res = req.GetResponse())
        using (var r = new StreamReader(res.GetResponseStream())) return r.ReadToEnd();
    }
}
