using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace OutlookListener
{
    // Talks to the user's own Work Track database (Cloud Firestore) over its REST API.
    // Signs in with the connector code's refresh token, so no Microsoft or company API is
    // involved and nothing needs IT approval. Firestore's security rules limit this to the
    // user's own users/{uid}/ documents. All calls are synchronous: run them off the UI thread.
    public class WorkTrackConnector
    {
        private readonly ConnectorCode _code;
        private string _idToken;
        private DateTime _idTokenExpires = DateTime.MinValue;
        private readonly object _lock = new object();
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        static WorkTrackConnector()
        {
            // TLS 1.2 (3072) for older .NET Framework defaults; use the system proxy with the
            // Windows login, which is what corporate proxies expect.
            ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072;
            if (WebRequest.DefaultWebProxy != null) WebRequest.DefaultWebProxy.Credentials = CredentialCache.DefaultCredentials;
        }

        // Overridable only for tests against the local Firebase emulators.
        public static string AuthBase = "https://securetoken.googleapis.com";
        public static string FirestoreBase = "https://firestore.googleapis.com";

        public WorkTrackConnector(ConnectorCode code)
        {
            _code = code;
        }

        private string DocsRoot
        {
            get { return FirestoreBase + "/v1/projects/" + _code.ProjectId + "/databases/(default)/documents"; }
        }

        private string UserPath
        {
            get { return DocsRoot + "/users/" + _code.Uid; }
        }

        // ---------- auth ----------

        private string IdToken()
        {
            lock (_lock)
            {
                if (_idToken != null && DateTime.UtcNow < _idTokenExpires) return _idToken;
                string body = "grant_type=refresh_token&refresh_token=" + Uri.EscapeDataString(_code.RefreshToken);
                var res = Send("POST", AuthBase + "/v1/token?key=" + Uri.EscapeDataString(_code.ApiKey),
                    body, "application/x-www-form-urlencoded", false);
                if (res.Status != 200) throw new WorkTrackException("Sign-in failed (" + res.Status + "). Create a new connection code in Work Track → Settings.");
                var d = Json.Deserialize<Dictionary<string, object>>(res.Body);
                _idToken = (string)d["id_token"];
                int seconds;
                int.TryParse(Convert.ToString(d["expires_in"], CultureInfo.InvariantCulture), out seconds);
                _idTokenExpires = DateTime.UtcNow.AddSeconds(Math.Max(60, seconds - 120));
                return _idToken;
            }
        }

        // ---------- public operations ----------

        /// <summary>Checks the code works by reading the user's settings document.</summary>
        public void Test()
        {
            var res = Send("GET", UserPath + "/settings/main", null, null, true);
            if (res.Status != 200 && res.Status != 404) throw new WorkTrackException("Work Track answered " + res.Status + ": " + Short(res.Body));
        }

        /// <summary>Creates a task with a fixed id. Returns false if it already existed (no duplicate is made).</summary>
        public bool CreateTask(string docId, IDictionary<string, object> fields)
        {
            var res = Send("POST", UserPath + "/tasks?documentId=" + Uri.EscapeDataString(docId),
                Json.Serialize(new Dictionary<string, object> { { "fields", Fields(fields) } }), "application/json", true);
            if (res.Status == 200) return true;
            if (res.Status == 409) return false;
            throw new WorkTrackException("Creating a task failed (" + res.Status + "): " + Short(res.Body));
        }

        /// <summary>Updates only the given fields of an existing task.</summary>
        public void UpdateTask(string docId, IDictionary<string, object> fields)
        {
            var url = new StringBuilder(UserPath + "/tasks/" + Uri.EscapeDataString(docId) + "?currentDocument.exists=true");
            foreach (string k in fields.Keys) url.Append("&updateMask.fieldPaths=").Append(Uri.EscapeDataString(k));
            var res = Send("PATCH", url.ToString(),
                Json.Serialize(new Dictionary<string, object> { { "fields", Fields(fields) } }), "application/json", true);
            if (res.Status == 200 || res.Status == 404) return;
            throw new WorkTrackException("Updating a task failed (" + res.Status + "): " + Short(res.Body));
        }

        /// <summary>Writes (replaces) a whole document under the user, e.g. "calendar/2026-10-28".</summary>
        public void SetDocument(string relativePath, IDictionary<string, object> fields)
        {
            var res = Send("PATCH", UserPath + "/" + relativePath,
                Json.Serialize(new Dictionary<string, object> { { "fields", Fields(fields) } }), "application/json", true);
            if (res.Status != 200) throw new WorkTrackException("Saving " + relativePath + " failed (" + res.Status + "): " + Short(res.Body));
        }

        /// <summary>Tasks that came from flagged mail and were finished in Work Track, but whose flag isn't cleared yet.</summary>
        public List<Dictionary<string, object>> FinishedFlaggedTasks()
        {
            return Query(new ArrayList
            {
                FieldFilter("source", "EQUAL", "outlook"),
                FieldFilter("status", "EQUAL", "done"),
                FieldFilter("outlookFlagSynced", "EQUAL", false)
            }, 50);
        }

        /// <summary>Tasks changed after the given time (epoch ms), for incremental refreshes.</summary>
        public List<Dictionary<string, object>> TasksUpdatedSince(long sinceMs)
        {
            return Query(new ArrayList { FieldFilter("updatedAt", "GREATER_THAN", sinceMs) }, 500);
        }

        /// <summary>Every task (paged). Each row has an "id" entry plus the task's fields.</summary>
        public List<Dictionary<string, object>> AllTasks()
        {
            var rows = new List<Dictionary<string, object>>();
            string token = null;
            do
            {
                string url = UserPath + "/tasks?pageSize=300" + (token != null ? "&pageToken=" + Uri.EscapeDataString(token) : "");
                var res = Send("GET", url, null, null, true);
                if (res.Status != 200) throw new WorkTrackException("Reading tasks failed (" + res.Status + "): " + Short(res.Body));
                var d = Json.Deserialize<Dictionary<string, object>>(res.Body);
                var docs = d.ContainsKey("documents") ? d["documents"] as IList : null;
                if (docs != null) foreach (object o in docs) rows.Add(Row((Dictionary<string, object>)o));
                token = d.ContainsKey("nextPageToken") ? d["nextPageToken"] as string : null;
            } while (!string.IsNullOrEmpty(token));
            return rows;
        }

        /// <summary>One document under the user (e.g. "settings/main"), or null if it doesn't exist.</summary>
        public Dictionary<string, object> GetDocument(string relativePath)
        {
            var res = Send("GET", UserPath + "/" + relativePath, null, null, true);
            if (res.Status == 404) return null;
            if (res.Status != 200) throw new WorkTrackException("Reading " + relativePath + " failed (" + res.Status + "): " + Short(res.Body));
            return Row(Json.Deserialize<Dictionary<string, object>>(res.Body));
        }

        private List<Dictionary<string, object>> Query(ArrayList filters, int limit)
        {
            object where = filters.Count == 1
                ? filters[0]
                : new Dictionary<string, object> { { "compositeFilter", new Dictionary<string, object> { { "op", "AND" }, { "filters", filters } } } };
            var query = new Dictionary<string, object>
            {
                { "structuredQuery", new Dictionary<string, object>
                    {
                        { "from", new ArrayList { new Dictionary<string, object> { { "collectionId", "tasks" } } } },
                        { "where", where },
                        { "limit", limit }
                    }
                }
            };
            var res = Send("POST", UserPath + ":runQuery", Json.Serialize(query), "application/json", true);
            if (res.Status != 200) throw new WorkTrackException("Query failed (" + res.Status + "): " + Short(res.Body));
            var rows = new List<Dictionary<string, object>>();
            foreach (object row in Json.Deserialize<ArrayList>(res.Body))
            {
                var r = row as Dictionary<string, object>;
                if (r != null && r.ContainsKey("document")) rows.Add(Row((Dictionary<string, object>)r["document"]));
            }
            return rows;
        }

        private static Dictionary<string, object> Row(Dictionary<string, object> doc)
        {
            var plain = FromFields(doc.ContainsKey("fields") ? (Dictionary<string, object>)doc["fields"] : new Dictionary<string, object>());
            string name = (string)doc["name"];
            plain["id"] = name.Substring(name.LastIndexOf('/') + 1);
            return plain;
        }

        // ---------- helpers ----------

        public static string DocIdFor(string outlookEntryId)
        {
            using (var sha = SHA1.Create())
            {
                byte[] h = sha.ComputeHash(Encoding.UTF8.GetBytes(outlookEntryId ?? ""));
                var sb = new StringBuilder("ol-");
                for (int i = 0; i < 12; i++) sb.Append(h[i].ToString("x2"));
                return sb.ToString();
            }
        }

        public static long NowMs()
        {
            return ToMs(DateTime.UtcNow);
        }

        public static long ToMs(DateTime utc)
        {
            return (long)(utc - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
        }

        /// <summary>A complete new task document, matching Work Track's data model (shared/DATA_MODEL.md).</summary>
        public static Dictionary<string, object> NewTask(string title, string notes, string person, string status, string dueDate, string entryId, bool fromFlag)
        {
            long now = NowMs();
            var t = new Dictionary<string, object>
            {
                { "title", title }, { "notes", notes }, { "project", null }, { "person", person },
                { "status", status }, { "dueDate", dueDate }, { "dueTime", null },
                { "todayDate", null }, { "todayRank", null }, { "waitingOn", null }, { "followUpDate", null },
                { "snoozedUntil", null }, { "recurrence", null }, { "seriesId", null },
                { "blocked", false }, { "blockedReason", null }, { "agendaFor", null },
                { "source", "outlook" }, { "createdAt", now }, { "updatedAt", now }, { "lastTouchedAt", now },
                { "completedAt", null }, { "deleted", false }, { "outlookEntryId", entryId }
            };
            if (fromFlag) t["outlookFlagSynced"] = false;
            return t;
        }

        private static Dictionary<string, object> FieldFilter(string field, string op, object value)
        {
            return new Dictionary<string, object>
            {
                { "fieldFilter", new Dictionary<string, object>
                    {
                        { "field", new Dictionary<string, object> { { "fieldPath", field } } },
                        { "op", op },
                        { "value", Value(value) }
                    }
                }
            };
        }

        // Firestore REST uses typed values: {"stringValue": "..."}, {"integerValue": "12"}, …
        public static Dictionary<string, object> Fields(IDictionary<string, object> plain)
        {
            var f = new Dictionary<string, object>();
            foreach (var kv in plain) f[kv.Key] = Value(kv.Value);
            return f;
        }

        public static Dictionary<string, object> Value(object v)
        {
            if (v == null) return new Dictionary<string, object> { { "nullValue", null } };
            if (v is string) return new Dictionary<string, object> { { "stringValue", v } };
            if (v is bool) return new Dictionary<string, object> { { "booleanValue", v } };
            if (v is int || v is long) return new Dictionary<string, object> { { "integerValue", Convert.ToInt64(v).ToString(CultureInfo.InvariantCulture) } };
            if (v is double || v is float) return new Dictionary<string, object> { { "doubleValue", Convert.ToDouble(v) } };
            var map = v as IDictionary<string, object>;
            if (map != null) return new Dictionary<string, object> { { "mapValue", new Dictionary<string, object> { { "fields", Fields(map) } } } };
            var list = v as IEnumerable;
            if (list != null)
            {
                var values = new ArrayList();
                foreach (object x in list) values.Add(Value(x));
                return new Dictionary<string, object> { { "arrayValue", new Dictionary<string, object> { { "values", values } } } };
            }
            return new Dictionary<string, object> { { "stringValue", v.ToString() } };
        }

        public static Dictionary<string, object> FromFields(Dictionary<string, object> fields)
        {
            var plain = new Dictionary<string, object>();
            if (fields == null) return plain;
            foreach (var kv in fields) plain[kv.Key] = Plain(kv.Value);
            return plain;
        }

        // Typed value -> plain value; maps and arrays are converted recursively.
        public static object Plain(object typedValue)
        {
            var tv = typedValue as Dictionary<string, object>;
            if (tv == null || tv.Count == 0) return null;
            foreach (var t in tv)
            {
                switch (t.Key)
                {
                    case "integerValue":
                        return long.Parse(Convert.ToString(t.Value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                    case "doubleValue":
                        return Convert.ToDouble(t.Value, CultureInfo.InvariantCulture);
                    case "mapValue":
                        var m = t.Value as Dictionary<string, object>;
                        return FromFields(m != null && m.ContainsKey("fields") ? (Dictionary<string, object>)m["fields"] : null);
                    case "arrayValue":
                        var list = new List<object>();
                        var a = t.Value as Dictionary<string, object>;
                        var values = a != null && a.ContainsKey("values") ? a["values"] as IList : null;
                        if (values != null) foreach (object v in values) list.Add(Plain(v));
                        return list;
                    default:
                        return t.Value; // stringValue, booleanValue, nullValue, timestampValue
                }
            }
            return null;
        }

        private struct Response
        {
            public int Status;
            public string Body;
        }

        private Response Send(string method, string url, string body, string contentType, bool auth)
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = method;
            req.Timeout = 20000;
            req.UserAgent = "OutlookListener-WorkTrack/1.0";
            if (auth) req.Headers["Authorization"] = "Bearer " + IdToken();
            if (body != null)
            {
                byte[] bytes = Encoding.UTF8.GetBytes(body);
                req.ContentType = contentType;
                req.ContentLength = bytes.Length;
                using (var s = req.GetRequestStream()) s.Write(bytes, 0, bytes.Length);
            }
            try
            {
                using (var res = (HttpWebResponse)req.GetResponse())
                using (var r = new StreamReader(res.GetResponseStream(), Encoding.UTF8))
                {
                    return new Response { Status = (int)res.StatusCode, Body = r.ReadToEnd() };
                }
            }
            catch (WebException ex)
            {
                var res = ex.Response as HttpWebResponse;
                if (res == null) throw new WorkTrackException("Can't reach Work Track (" + ex.Status + "). Check the internet connection.");
                using (res)
                using (var r = new StreamReader(res.GetResponseStream(), Encoding.UTF8))
                {
                    return new Response { Status = (int)res.StatusCode, Body = r.ReadToEnd() };
                }
            }
        }

        private static string Short(string s)
        {
            if (s == null) return "";
            s = s.Replace("\r", " ").Replace("\n", " ");
            return s.Length > 200 ? s.Substring(0, 200) + "…" : s;
        }
    }

    public class WorkTrackException : Exception
    {
        public WorkTrackException(string message) : base(message) { }
    }
}
