using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace OutlookListener
{
    // Work Track reminders on this computer: the same nudges the phone shows (morning plan,
    // overdue nags, follow-ups, evening review, weekly recap, 1:1 prep, due times, meeting ends),
    // as Windows notifications from a tray icon. The tray icon is the listener's own icon with
    // today's count on it; hovering shows the #1 task, right-click offers Done / Next / Snooze.
    //
    // Tasks are kept in memory: one full read, then every 2 minutes only tasks changed since the
    // last read (plus the settings and today's plan status), so it stays far inside Firebase's
    // free daily reads. Everything here runs on the listener's main thread except network calls.
    public class WorkTrackReminders
    {
        private const long CatchUpMs = 30 * 60 * 1000L;
        private const long OverlapMs = 10 * 60 * 1000L;   // tolerates devices whose clocks differ

        private readonly Action<string> _log;
        private readonly Action<Action> _post;
        private WorkTrackConnector _conn;
        private string _appUrl = ConnectorCode.DefaultAppUrl;
        private bool _enabled;

        private readonly Dictionary<string, WtTask> _tasks = new Dictionary<string, WtTask>();
        private WtSettings _wt = new WtSettings();
        private WtDay _day = new WtDay();
        private List<WtCalEvent> _calendar = new List<WtCalEvent>();
        private bool _loaded, _polling;
        private long _lastSeenUpdated;
        private DateTime _lastFull = DateTime.MinValue;

        private NotifyIcon _tray;
        private Icon _baseIcon, _badgeIcon;
        private System.Windows.Forms.Timer _tick, _poll, _pump;
        private readonly Queue<WtContent> _pending = new Queue<WtContent>();
        private DateTime _lastBalloon = DateTime.MinValue;
        private string _lastRoute = "today";
        private State _state;

        private class State
        {
            public long LastFired { get; set; }
            public long SnoozeUntil { get; set; }
            public string PausedDate { get; set; }
        }

        private static readonly string StatePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookListener", "worktrack-state.json");

        public WorkTrackReminders(Action<string> log, Action<Action> postToMainThread)
        {
            _log = log;
            _post = postToMainThread;
        }

        /// <summary>Turns reminders on or off (called whenever worktrack.json changes).</summary>
        public void Configure(bool enabled, ConnectorCode code)
        {
            _enabled = enabled && code != null;
            if (!_enabled)
            {
                Stop();
                return;
            }
            _conn = new WorkTrackConnector(code);
            _appUrl = string.IsNullOrEmpty(code.AppUrl) ? ConnectorCode.DefaultAppUrl : code.AppUrl;
            if (_tray == null) Start();
            _lastFull = DateTime.MinValue;
            Poll();
        }

        public void SetCalendar(List<WtCalEvent> events)
        {
            _calendar = events ?? new List<WtCalEvent>();
        }

        private void Start()
        {
            _state = LoadState();
            try { _baseIcon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            if (_baseIcon == null) _baseIcon = SystemIcons.Information;
            _tray = new NotifyIcon { Icon = _baseIcon, Text = "Work Track", Visible = true, ContextMenuStrip = new ContextMenuStrip() };
            _tray.ContextMenuStrip.Opening += (s, e) => { BuildMenu(); e.Cancel = false; };
            _tray.BalloonTipClicked += (s, e) => Open(_lastRoute);
            _tray.DoubleClick += (s, e) => Open("today");

            _tick = new System.Windows.Forms.Timer { Interval = 30 * 1000 };
            _tick.Tick += (s, e) => Tick();
            _tick.Start();
            _poll = new System.Windows.Forms.Timer { Interval = 2 * 60 * 1000 };
            _poll.Tick += (s, e) => Poll();
            _poll.Start();
            _pump = new System.Windows.Forms.Timer { Interval = 1000 };
            _pump.Tick += (s, e) => ShowNext();
            _pump.Start();
            _log("[Work Track] Reminders on this computer: on.");
        }

        public void Stop()
        {
            foreach (var t in new[] { _tick, _poll, _pump }) if (t != null) { t.Stop(); t.Dispose(); }
            _tick = _poll = _pump = null;
            if (_tray != null)
            {
                _tray.Visible = false;
                _tray.Dispose();
                _tray = null;
                _log("[Work Track] Reminders on this computer: off.");
            }
            _pending.Clear();
        }

        // ---------- data ----------

        private TimeZoneInfo Tz
        {
            get { return WtTime.Resolve(_wt.Timezone); }
        }

        private WtCtx Ctx()
        {
            var tz = Tz;
            return new WtCtx(WtTime.DateOf(WorkTrackConnector.NowMs(), tz), _wt, tz);
        }

        private void Poll()
        {
            if (!_enabled || _polling || _conn == null) return;
            _polling = true;
            var conn = _conn;
            bool full = DateTime.Now - _lastFull > TimeSpan.FromHours(1);
            long since = _lastSeenUpdated - OverlapMs;
            string today = WtTime.Iso(Ctx().Today);
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    var rows = full ? conn.AllTasks() : conn.TasksUpdatedSince(since);
                    var settings = conn.GetDocument("settings/main");
                    var day = conn.GetDocument("days/" + today);
                    _post(() => Merge(rows, full, settings, day));
                }
                catch (Exception ex)
                {
                    _log("[Work Track] Couldn't refresh reminders: " + ex.Message);
                    _post(() => _polling = false);
                }
            });
        }

        private void Merge(List<Dictionary<string, object>> rows, bool full, Dictionary<string, object> settings, Dictionary<string, object> day)
        {
            _polling = false;
            if (full) { _tasks.Clear(); _lastFull = DateTime.Now; }
            foreach (var r in rows)
            {
                string id = (string)r["id"];
                r.Remove("id");
                var t = WtTask.FromPlain(id, r);
                _tasks[id] = t;
                if (t.UpdatedAt > _lastSeenUpdated) _lastSeenUpdated = t.UpdatedAt;
            }
            if (settings != null) _wt = WtSettings.FromPlain(settings);
            _day = new WtDay
            {
                MorningPlanDone = day != null && day.ContainsKey("morningPlanDone") && Equals(day["morningPlanDone"], true),
                EveningReviewDone = day != null && day.ContainsKey("eveningReviewDone") && Equals(day["eveningReviewDone"], true)
            };
            bool first = !_loaded;
            _loaded = true;
            UpdateBadge();
            if (first) Tick();
        }

        // ---------- firing ----------

        private void Tick()
        {
            if (!_enabled || !_loaded) return;
            long now = WorkTrackConnector.NowMs();
            long from = Math.Max(_state.LastFired + 1, now - CatchUpMs);
            long to = now + 60 * 1000;
            _state.LastFired = to;
            SaveState();
            UpdateBadge();
            if (from > to) return;

            var ctx = Ctx();
            if (_state.PausedDate == WtTime.Iso(ctx.Today)) return;
            var tasks = _tasks.Values.ToList();
            var slots = WtSchedule.Slots(from, to, _wt, ctx.Tz, tasks, true)
                .Concat(WtSchedule.MeetingEndSlots(_calendar, _wt, ctx.Tz, from, to))
                .OrderBy(s => s.At);
            foreach (var slot in slots)
            {
                if (slot.Kind == "nag" && now < _state.SnoozeUntil) continue;
                var content = WtSchedule.Content(slot, tasks, ctx, _day);
                if (content != null) _pending.Enqueue(content);
            }
        }

        // One notification at a time, a few seconds apart, so none are lost.
        private void ShowNext()
        {
            if (_tray == null || _pending.Count == 0 || DateTime.Now - _lastBalloon < TimeSpan.FromSeconds(8)) return;
            var c = _pending.Dequeue();
            _lastRoute = c.Route;
            _lastBalloon = DateTime.Now;
            string body = string.IsNullOrEmpty(c.Body) ? "Open Work Track" : c.Body;
            _tray.ShowBalloonTip(15000, c.Title, body.Length > 250 ? body.Substring(0, 247) + "…" : body, ToolTipIcon.Info);
            _log("[Work Track] Reminder: " + c.Title);
        }

        // ---------- tray ----------

        private void UpdateBadge()
        {
            if (_tray == null) return;
            var ctx = Ctx();
            var today = WtRules.OpenToday(_tasks.Values, ctx);
            string tip = today.Count == 0
                ? "Work Track · plan your day"
                : "Work Track · " + today.Count + " of " + Math.Max(today.Count, _wt.TodayCap) + " left · Now: " + today[0].Title;
            _tray.Text = tip.Length > 63 ? tip.Substring(0, 62) + "…" : tip;   // Windows limit

            var old = _badgeIcon;
            _badgeIcon = today.Count == 0 ? null : DrawBadge(today.Count);
            _tray.Icon = _badgeIcon ?? _baseIcon;
            if (old != null) { FreeIcon(old); }
        }

        // The listener's own icon with the count in a small TKE-purple dot (bottom-right).
        private Icon DrawBadge(int count)
        {
            var size = SystemInformation.SmallIconSize;
            using (var bmp = new Bitmap(size.Width, size.Height))
            {
                using (var g = Graphics.FromImage(bmp))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                    using (var icon = new Icon(_baseIcon, size)) g.DrawIcon(icon, new Rectangle(Point.Empty, size));
                    int d = (int)(size.Width * 0.75);
                    var r = new Rectangle(size.Width - d, size.Height - d, d, d);
                    using (var b = new SolidBrush(TkeTheme.Purple)) g.FillEllipse(b, r);
                    using (var f = new Font(FontFamily.GenericSansSerif, d * 0.8F, FontStyle.Bold, GraphicsUnit.Pixel))
                    using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        g.DrawString(count > 9 ? "9" : count.ToString(), f, Brushes.White, r, sf);
                }
                return Icon.FromHandle(bmp.GetHicon());
            }
        }

        [DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr handle);

        // Icon.FromHandle doesn't own its handle, so free it ourselves (Windows only).
        private static void FreeIcon(Icon icon)
        {
            try { DestroyIcon(icon.Handle); } catch (Exception) { }
            icon.Dispose();
        }

        private void BuildMenu()
        {
            var menu = _tray.ContextMenuStrip;
            menu.Items.Clear();
            menu.Font = TkeTheme.Body(9.5F);
            var ctx = Ctx();
            var today = WtRules.OpenToday(_tasks.Values, ctx);
            var top = today.FirstOrDefault();

            menu.Items.Add(new ToolStripMenuItem(top != null ? "Now: " + Short(top.Title) : "Nothing picked for today") { Enabled = false });
            if (top != null)
            {
                menu.Items.Add("Done: " + Short(top.Title), null, (s, e) => Complete(top));
                if (today.Count > 1) menu.Items.Add("Next task", null, (s, e) => Rotate(today));
            }
            else
            {
                menu.Items.Add("Plan your day…", null, (s, e) => Open("plan"));
            }
            menu.Items.Add(new ToolStripSeparator());
            long now = WorkTrackConnector.NowMs();
            if (now < _state.SnoozeUntil)
                menu.Items.Add("Overdue reminders snoozed until " + WtTime.Local(_state.SnoozeUntil, ctx.Tz).ToString("HH:mm") + " — resume", null,
                    (s, e) => { _state.SnoozeUntil = 0; SaveState(); });
            else
                menu.Items.Add("Snooze overdue reminders for 2 hours", null, (s, e) => { _state.SnoozeUntil = now + 2 * 3600 * 1000L; SaveState(); });
            bool paused = _state.PausedDate == WtTime.Iso(ctx.Today);
            menu.Items.Add(paused ? "Resume reminders today" : "Pause all reminders for today", null, (s, e) =>
            {
                _state.PausedDate = paused ? null : WtTime.Iso(ctx.Today);
                SaveState();
            });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Open Work Track", null, (s, e) => Open("today"));
        }

        private static string Short(string s)
        {
            return s.Length > 48 ? s.Substring(0, 47) + "…" : s;
        }

        private void Open(string route)
        {
            try { Process.Start(_appUrl + "#/" + (route ?? "today")); }
            catch (Exception ex) { _log("[Work Track] Couldn't open the browser: " + ex.Message); }
        }

        // ---------- actions (same writes the phone makes) ----------

        private void Complete(WtTask t)
        {
            var conn = _conn;
            var ctx = Ctx();
            long now = WorkTrackConnector.NowMs();
            t.Status = "done";
            t.CompletedAt = now;
            UpdateBadge();
            var next = WtRules.NextInstance(t, ctx.Today, _wt.WorkDays);
            Dictionary<string, object> nextFields = null;
            if (next.HasValue)
            {
                nextFields = new Dictionary<string, object>(t.Raw);
                nextFields["seriesId"] = string.IsNullOrEmpty(t.SeriesId) ? t.Id : t.SeriesId;
                nextFields["status"] = "open";
                nextFields["dueDate"] = WtTime.Iso(next.Value.Value);
                nextFields["todayDate"] = null;
                nextFields["todayRank"] = null;
                nextFields["snoozedUntil"] = null;
                nextFields["completedAt"] = null;
                nextFields["deleted"] = false;
                nextFields["createdAt"] = now;
                nextFields["updatedAt"] = now;
                nextFields["lastTouchedAt"] = now;
            }
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    conn.UpdateTask(t.Id, new Dictionary<string, object> { { "status", "done" }, { "completedAt", now }, { "updatedAt", now }, { "lastTouchedAt", now } });
                    if (nextFields != null) conn.SetDocument("tasks/" + next.Value.Key, nextFields);
                    _log("[Work Track] Done from the tray: " + t.Title);
                    _post(Poll);
                }
                catch (Exception ex) { _log("[Work Track] Couldn't mark done: " + ex.Message); }
            });
        }

        private void Rotate(List<WtTask> today)
        {
            var conn = _conn;
            var order = today.Skip(1).Concat(today.Take(1)).ToList();
            string date = WtTime.Iso(Ctx().Today);
            long now = WorkTrackConnector.NowMs();
            for (int i = 0; i < order.Count; i++) order[i].TodayRank = i + 1;
            UpdateBadge();
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    for (int i = 0; i < order.Count; i++)
                        conn.UpdateTask(order[i].Id, new Dictionary<string, object> { { "todayDate", date }, { "todayRank", i + 1 }, { "updatedAt", now } });
                }
                catch (Exception ex) { _log("[Work Track] Couldn't reorder: " + ex.Message); }
            });
        }

        // ---------- state ----------

        private static State LoadState()
        {
            try
            {
                if (File.Exists(StatePath)) return new JavaScriptSerializer().Deserialize<State>(File.ReadAllText(StatePath)) ?? new State();
            }
            catch { }
            return new State();
        }

        private void SaveState()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(StatePath));
                File.WriteAllText(StatePath, new JavaScriptSerializer().Serialize(_state));
            }
            catch { }
        }
    }
}
