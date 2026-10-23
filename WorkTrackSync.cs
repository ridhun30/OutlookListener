using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookListener
{
    // Connects Outlook to Work Track:
    //   - flagged mail becomes a task; completing/clearing the flag finishes it, and
    //     finishing the task in Work Track completes the flag (checked every 5 minutes);
    //   - new mail that looks like an ask (or is from the boss / a VIP) lands in the inbox;
    //   - today's and tomorrow's meetings are sent to the phone for "any action items?" nudges.
    // Outlook (COM) is only touched on the listener's main thread; network calls run in the
    // background, and anything they need to do back in Outlook is queued to the main thread.
    public class WorkTrackSync
    {
        private readonly Outlook.NameSpace _ns;
        private readonly Outlook.Items _inbox;
        private Outlook.Items _calendar;
        private readonly Action<string> _log;
        private WorkTrackSettings _settings;
        private WorkTrackConnector _conn;
        private string _myName = "";
        private string _myAddress = "";

        // EntryID -> "flagStatus|dueDate", so only real flag changes are acted on.
        private readonly Dictionary<string, string> _flagState = new Dictionary<string, string>();
        private readonly ConcurrentQueue<Action> _mainThread = new ConcurrentQueue<Action>();
        private System.Windows.Forms.Timer _pump, _calendarTimer, _pollTimer;
        private DateTime _calendarDue = DateTime.MaxValue;
        private FileSystemWatcher _watcher;

        public WorkTrackSync(Outlook.NameSpace ns, Outlook.Items inboxItems, Action<string> log)
        {
            _ns = ns;
            _inbox = inboxItems;
            _log = log;
        }

        public void Start()
        {
            try
            {
                _myName = _ns.CurrentUser.Name ?? "";
                _myAddress = SmtpOf(_ns.CurrentUser.AddressEntry);
            }
            catch { }

            _inbox.ItemChange += OnInboxItemChange;
            try
            {
                _calendar = _ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).Items;
                _calendar.ItemAdd += OnCalendarChange;
                _calendar.ItemChange += OnCalendarChange;
            }
            catch (Exception ex) { _log("[Work Track] Calendar not available: " + ex.Message); }

            _pump = new System.Windows.Forms.Timer { Interval = 1000 };
            _pump.Tick += (s, e) => Pump();
            _pump.Start();
            _calendarTimer = new System.Windows.Forms.Timer { Interval = 30 * 60 * 1000 };
            _calendarTimer.Tick += (s, e) => PushCalendar();
            _calendarTimer.Start();
            _pollTimer = new System.Windows.Forms.Timer { Interval = 5 * 60 * 1000 };
            _pollTimer.Tick += (s, e) => PollFinishedTasks();
            _pollTimer.Start();

            string dir = Path.GetDirectoryName(WorkTrackSettings.SettingsPath);
            Directory.CreateDirectory(dir);
            _watcher = new FileSystemWatcher(dir, Path.GetFileName(WorkTrackSettings.SettingsPath));
            _watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName;
            FileSystemEventHandler reload = (s, e) => _mainThread.Enqueue(Reload);
            _watcher.Changed += reload;
            _watcher.Created += reload;
            _watcher.EnableRaisingEvents = true;

            Reload();
        }

        public void Stop()
        {
            try
            {
                if (_watcher != null) _watcher.Dispose();
                foreach (var t in new[] { _pump, _calendarTimer, _pollTimer }) if (t != null) t.Stop();
                _inbox.ItemChange -= OnInboxItemChange;
                if (_calendar != null)
                {
                    _calendar.ItemAdd -= OnCalendarChange;
                    _calendar.ItemChange -= OnCalendarChange;
                }
            }
            catch { }
        }

        private bool Active
        {
            get { return _conn != null && _settings != null && _settings.Enabled; }
        }

        private void Reload()
        {
            Thread.Sleep(150); // the dashboard may still be writing
            _settings = WorkTrackSettings.Load();
            var code = _settings.GetCode();
            _conn = code != null ? new WorkTrackConnector(code) : null;
            if (!_settings.Enabled || _conn == null)
            {
                _log("[Work Track] Off. Connect it from the Rules Dashboard (Work Track button).");
                return;
            }
            _log("[Work Track] On: " + Describe());
            if (_settings.FlagSync) ScanFlagged();
            if (_settings.CalendarPush) PushCalendar();
        }

        private string Describe()
        {
            var on = new List<string>();
            if (_settings.FlagSync) on.Add("flagged mail");
            if (_settings.AskDetection) on.Add("asks");
            if (_settings.CalendarPush) on.Add("calendar");
            return on.Count == 0 ? "nothing selected" : string.Join(", ", on.ToArray());
        }

        // ---------- threading ----------

        private void Pump()
        {
            Action a;
            while (_mainThread.TryDequeue(out a))
            {
                try { a(); } catch (Exception ex) { _log("[Work Track] " + ex.Message); }
            }
            if (DateTime.Now >= _calendarDue)
            {
                _calendarDue = DateTime.MaxValue;
                PushCalendar();
            }
        }

        private void Background(string what, Action work)
        {
            var conn = _conn;
            if (conn == null) return;
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try { work(); }
                catch (Exception ex) { _log("[Work Track] " + what + " failed: " + ex.Message); }
            });
        }

        // ---------- new mail: asks ----------

        public void OnNewMail(Outlook.MailItem mail)
        {
            if (!Active) return;
            if (_settings.FlagSync && (int)mail.FlagStatus == (int)Outlook.OlFlagStatus.olFlagMarked)
            {
                HandleFlag(mail, false);
                return;
            }
            if (!_settings.AskDetection) return;

            string senderName = Safe(() => mail.SenderName);
            string sender = senderName + " " + Safe(() => mail.SenderEmailAddress);
            string subject = Safe(() => mail.Subject);
            string body = Safe(() => mail.Body);
            bool boss = AskDetector.NameMatches(sender, _settings.BossNames);
            bool vip = AskDetector.NameMatches(sender, _settings.VipNames);
            string matched = null;
            bool ask = boss || vip || (IsToMe(mail) && AskDetector.LooksLikeAsk(subject, body, _settings.AskPhrases, _settings.IgnorePhrases, out matched));
            if (!ask) return;

            string entryId = mail.EntryID;
            string title = AskDetector.IsTeamsDigest(sender, subject) ? AskDetector.TeamsTitle(body, subject) : AskDetector.CleanSubject(subject);
            string why = boss ? "from your boss" : vip ? "from a VIP" : "looks like an ask (\"" + matched + "\")";
            string notes = "Outlook mail from " + senderName + ", " + Safe(() => mail.ReceivedTime.ToString("ddd d MMM HH:mm", CultureInfo.InvariantCulture))
                + " — " + why + "\n\n" + Snippet(AskDetector.StripQuoted(body), 300);
            var task = WorkTrackConnector.NewTask(title, notes, boss ? _settings.BossPerson : senderName, "inbox", null, entryId, false);
            Background("Adding \"" + title + "\"", () =>
            {
                if (_conn.CreateTask(WorkTrackConnector.DocIdFor(entryId), task))
                    _log("  -> Work Track: added to Inbox (" + why + "): " + title);
            });
        }

        private bool IsToMe(Outlook.MailItem mail)
        {
            try
            {
                foreach (Outlook.Recipient r in mail.Recipients)
                {
                    if (r.Type != (int)Outlook.OlMailRecipientType.olTo) continue;
                    string name = Safe(() => r.Name);
                    string addr = SmtpOf(r.AddressEntry);
                    if ((_myAddress.Length > 0 && string.Equals(addr, _myAddress, StringComparison.OrdinalIgnoreCase))
                        || (_myName.Length > 0 && string.Equals(name, _myName, StringComparison.OrdinalIgnoreCase))) return true;
                }
            }
            catch { }
            string to = Safe(() => mail.To);
            return _myName.Length > 0 && to.IndexOf(_myName, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // ---------- flags ----------

        private void OnInboxItemChange(object item)
        {
            if (!Active || !_settings.FlagSync) return;
            var mail = item as Outlook.MailItem;
            if (mail != null) HandleFlag(mail, false);
        }

        private void ScanFlagged()
        {
            try
            {
                var flagged = _inbox.Restrict("[FlagStatus] = 2");
                int n = 0;
                foreach (object o in flagged)
                {
                    var mail = o as Outlook.MailItem;
                    if (mail == null) continue;
                    HandleFlag(mail, true);
                    n++;
                }
                _log("[Work Track] Checked " + n + " flagged mail(s).");
            }
            catch (Exception ex) { _log("[Work Track] Scanning flagged mail failed: " + ex.Message); }
        }

        // fromScan: the startup scan only creates missing tasks. A live change (the user just
        // flagged the mail) also turns an existing task - e.g. one captured as an ask - into an
        // open flagged task; the scan must not do that, or it would reopen tasks finished in
        // Work Track whose flag hasn't been completed yet.
        private void HandleFlag(Outlook.MailItem mail, bool fromScan)
        {
            string entryId = mail.EntryID;
            int flag = (int)mail.FlagStatus;
            string due = null;
            try
            {
                DateTime d = mail.TaskDueDate;
                if (d.Year < 4000) due = d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            catch { }
            string state = flag + "|" + due;
            string prev;
            bool known = _flagState.TryGetValue(entryId, out prev);
            if (known && prev == state) return;
            _flagState[entryId] = state;
            int prevFlag = known ? int.Parse(prev.Split('|')[0]) : -1;
            string docId = WorkTrackConnector.DocIdFor(entryId);
            long now = WorkTrackConnector.NowMs();

            if (flag == (int)Outlook.OlFlagStatus.olFlagMarked)
            {
                string title = AskDetector.CleanSubject(Safe(() => mail.Subject));
                string senderName = Safe(() => mail.SenderName);
                string sender = senderName + " " + Safe(() => mail.SenderEmailAddress);
                string person = AskDetector.NameMatches(sender, _settings.BossNames) ? _settings.BossPerson : senderName;
                string notes = "Flagged in Outlook — mail from " + senderName + ", "
                    + Safe(() => mail.ReceivedTime.ToString("ddd d MMM HH:mm", CultureInfo.InvariantCulture))
                    + "\n\n" + Snippet(AskDetector.StripQuoted(Safe(() => mail.Body)), 300);
                var task = WorkTrackConnector.NewTask(title, notes, person, "open", due, entryId, true);
                bool wasFlagged = prevFlag == (int)Outlook.OlFlagStatus.olFlagMarked;
                Background("Adding flagged \"" + title + "\"", () =>
                {
                    if (_conn.CreateTask(docId, task))
                    {
                        _log("  -> Work Track: flagged mail added: " + title);
                    }
                    else if (!fromScan)
                    {
                        // Already a task (captured as an ask, or the due date changed): make it a live flagged task.
                        var patch = new Dictionary<string, object> { { "status", "open" }, { "outlookFlagSynced", false }, { "updatedAt", now }, { "lastTouchedAt", now } };
                        if (due != null) patch["dueDate"] = due;
                        if (!wasFlagged) patch["completedAt"] = null;
                        _conn.UpdateTask(docId, patch);
                    }
                });
            }
            else if (prevFlag == (int)Outlook.OlFlagStatus.olFlagMarked)
            {
                bool completed = flag == (int)Outlook.OlFlagStatus.olFlagComplete;
                var patch = new Dictionary<string, object>
                {
                    { "status", completed ? "done" : "dropped" }, { "completedAt", now }, { "outlookFlagSynced", true },
                    { "updatedAt", now }, { "lastTouchedAt", now }, { "todayDate", null }, { "todayRank", null }
                };
                Background("Finishing a flagged task", () =>
                {
                    _conn.UpdateTask(docId, patch);
                    _log("  -> Work Track: flag " + (completed ? "completed" : "cleared") + ", task " + (completed ? "done" : "dropped"));
                });
            }
        }

        // Tasks finished in Work Track: complete their Outlook flags.
        private void PollFinishedTasks()
        {
            if (!Active || !_settings.FlagSync) return;
            Background("Checking finished tasks", () =>
            {
                foreach (var row in _conn.FinishedFlaggedTasks())
                {
                    var r = row;
                    _mainThread.Enqueue(() => CompleteFlag(r));
                }
            });
        }

        private void CompleteFlag(Dictionary<string, object> row)
        {
            string id = (string)row["id"];
            object entryObj;
            string entryId = row.TryGetValue("outlookEntryId", out entryObj) ? entryObj as string : null;
            if (!string.IsNullOrEmpty(entryId))
            {
                try
                {
                    var mail = _ns.GetItemFromID(entryId) as Outlook.MailItem;
                    if (mail != null && (int)mail.FlagStatus == (int)Outlook.OlFlagStatus.olFlagMarked)
                    {
                        _flagState[entryId] = (int)Outlook.OlFlagStatus.olFlagComplete + "|"; // don't echo back
                        mail.FlagStatus = Outlook.OlFlagStatus.olFlagComplete;
                        mail.Save();
                        _log("  -> Work Track: task finished on the phone/web — completed the Outlook flag: " + Safe(() => mail.Subject));
                    }
                }
                catch (Exception ex) { _log("[Work Track] Couldn't update the Outlook flag (mail moved or deleted?): " + ex.Message); }
            }
            Background("Marking flag synced", () => _conn.UpdateTask(id, new Dictionary<string, object> { { "outlookFlagSynced", true } }));
        }

        // ---------- calendar ----------

        private void OnCalendarChange(object item)
        {
            _calendarDue = DateTime.Now.AddSeconds(10); // debounce bursts of changes
        }

        private void PushCalendar()
        {
            if (!Active || !_settings.CalendarPush || _calendar == null) return;
            var days = new List<KeyValuePair<string, List<object>>>();
            try
            {
                for (int i = 0; i < 2; i++)
                {
                    DateTime start = DateTime.Today.AddDays(i);
                    days.Add(new KeyValuePair<string, List<object>>(start.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), ReadDay(start)));
                }
            }
            catch (Exception ex)
            {
                _log("[Work Track] Reading the calendar failed: " + ex.Message);
                return;
            }
            Background("Sending meetings", () =>
            {
                long now = WorkTrackConnector.NowMs();
                foreach (var d in days)
                    _conn.SetDocument("calendar/" + d.Key, new Dictionary<string, object> { { "events", d.Value }, { "updatedAt", now }, { "source", "outlook" } });
            });
        }

        private List<object> ReadDay(DateTime day)
        {
            var events = new List<object>();
            _calendar.IncludeRecurrences = true;
            _calendar.Sort("[Start]");
            string filter = "[Start] < '" + day.AddDays(1).ToString("g") + "' AND [End] > '" + day.ToString("g") + "'";
            var items = _calendar.Restrict(filter);
            object o = items.GetFirst();
            while (o != null)
            {
                var a = o as Outlook.AppointmentItem;
                if (a != null && !a.AllDayEvent && a.BusyStatus != Outlook.OlBusyStatus.olFree)
                {
                    events.Add(new Dictionary<string, object>
                    {
                        { "id", Safe(() => a.GlobalAppointmentID) + "-" + a.Start.ToString("yyyyMMddHHmm", CultureInfo.InvariantCulture) },
                        { "title", Safe(() => a.Subject) },
                        { "start", WorkTrackConnector.ToMs(a.Start.ToUniversalTime()) },
                        { "end", WorkTrackConnector.ToMs(a.End.ToUniversalTime()) },
                        { "attendees", Math.Max(1, a.Recipients.Count) }
                    });
                }
                o = items.GetNext();
            }
            return events;
        }

        // ---------- helpers ----------

        private static string SmtpOf(Outlook.AddressEntry e)
        {
            if (e == null) return "";
            try
            {
                var ex = e.GetExchangeUser();
                if (ex != null && !string.IsNullOrEmpty(ex.PrimarySmtpAddress)) return ex.PrimarySmtpAddress;
            }
            catch { }
            return Safe(() => e.Address);
        }

        private static string Snippet(string s, int max)
        {
            s = (s ?? "").Trim();
            return s.Length > max ? s.Substring(0, max) + "…" : s;
        }

        private static string Safe(Func<string> get)
        {
            try { return get() ?? ""; } catch { return ""; }
        }
    }
}
