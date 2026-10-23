using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace OutlookListener
{
    // Dashboard dialog that links the listener to Work Track. Three plain steps:
    // 1) paste the connection code from Work Track, 2) choose what to send,
    // 3) say whose mail always matters. Saving writes worktrack.json, which the running
    // listener picks up on its own (no restart).
    public class WorkTrackForm : Form
    {
        private readonly WorkTrackSettings _settings;
        private ConnectorCode _code;

        private Label _lblStatus;
        private TextBox _txtCode;
        private Label _lblTestResult;
        private CheckBox _chkFlags, _chkAsks, _chkCalendar, _chkPaused;
        private TextBox _txtBoss, _txtBossPerson, _txtVips, _txtPhrases, _txtIgnore;
        private Panel _pnlAdvanced;
        private Button _btnAdvanced;

        private const int Wd = 640;

        public WorkTrackForm()
        {
            _settings = WorkTrackSettings.Load();
            _code = _settings.GetCode();

            Text = "Work Track";
            Width = Wd;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = TkeTheme.Gray5;
            ForeColor = TkeTheme.Gray1;
            Font = TkeTheme.Body(10F);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            BuildUi();
            ShowStatus();
        }

        private Label Heading(string text, int y)
        {
            var l = new Label { Text = text, Left = 14, Top = y, Width = Wd - 40, Height = 24, Font = TkeTheme.Bold(11F), ForeColor = TkeTheme.Black };
            Controls.Add(l);
            return l;
        }

        private Label Help(Control parent, string text, int left, int y, int width, int height)
        {
            var l = new Label { Text = text, Left = left, Top = y, Width = width, Height = height, ForeColor = TkeTheme.Gray1, Font = TkeTheme.Body(9F) };
            parent.Controls.Add(l);
            return l;
        }

        private CheckBox Option(string text, string help, int y)
        {
            var c = new CheckBox { Text = text, Left = 14, Top = y, Width = Wd - 40, Height = 24, Font = TkeTheme.Medium(10F), ForeColor = TkeTheme.Black, BackColor = TkeTheme.Gray5 };
            Controls.Add(c);
            Help(this, help, 32, y + 24, Wd - 60, 34);
            return c;
        }

        private void BuildUi()
        {
            var header = new Panel { Dock = DockStyle.Top, Height = 72, BackColor = TkeTheme.Black };
            header.Paint += (s, e) => TkeTheme.PaintHeader(e.Graphics, header.ClientRectangle,
                "Work Track", 16F, 18, "Send flagged mail, asks and meetings to your Work Track app", 9.5F, 46);
            Controls.Add(header);

            int y = 84;
            _lblStatus = new Label { Left = 14, Top = y, Width = Wd - 40, Height = 24, Font = TkeTheme.Bold(10.5F) };
            Controls.Add(_lblStatus);
            y += 32;

            // Step 1
            Heading("1.  Paste your connection code", y);
            y += 26;
            Help(this, "In Work Track open Settings → Connect laptop (Outlook) → Copy code, then paste it here. "
                + "Treat it like a password: it lets this computer add tasks to your Work Track.", 14, y, Wd - 40, 36);
            y += 40;
            _txtCode = new TextBox { Left = 14, Top = y, Width = 360, BackColor = Color.White, ForeColor = TkeTheme.Gray1, UseSystemPasswordChar = true };
            Controls.Add(_txtCode);
            var btnPaste = TkeTheme.CreateSecondaryButton("Paste", 384, y - 6, 92, 34);
            btnPaste.Click += (s, e) => { if (Clipboard.ContainsText()) _txtCode.Text = Clipboard.GetText().Trim(); };
            Controls.Add(btnPaste);
            var btnTest = TkeTheme.CreatePrimaryButton("Test", 486, y - 6, 120, 34);
            btnTest.Click += (s, e) => TestConnection();
            Controls.Add(btnTest);
            y += 32;
            _lblTestResult = Help(this, "", 14, y, Wd - 40, 22);
            y += 30;

            // Step 2
            Heading("2.  Choose what to send", y);
            y += 30;
            _chkFlags = Option("Flagged mail becomes a task",
                "Flag a mail (Follow up) and it appears in Work Track. Finishing it in either place finishes it in both.", y);
            y += 62;
            _chkAsks = Option("Mail that asks me to do something goes to my Work Track inbox",
                "Only mail where you are on To (not CC), e.g. \"can you…\", \"please review\", \"by EOD\". You decide later whether to keep it.", y);
            y += 70;
            _chkCalendar = Option("Send today's meetings to my phone",
                "After each meeting your phone asks \"any action items?\". Only titles and times are sent.", y);
            y += 66;

            // Step 3
            Heading("3.  Whose mail always counts?", y);
            y += 30;
            Help(this, "Your boss (name or email as shown in Outlook):", 14, y + 3, 290, 20);
            _txtBoss = new TextBox { Left = 310, Top = y, Width = 296, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            Controls.Add(_txtBoss);
            y += 32;
            Help(this, "Other people (separate with ;):", 14, y + 3, 290, 20);
            _txtVips = new TextBox { Left = 310, Top = y, Width = 296, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            Controls.Add(_txtVips);
            y += 40;

            // Advanced (collapsed)
            _btnAdvanced = TkeTheme.CreateSecondaryButton("Advanced ▸", 14, y, 130, 32);
            _btnAdvanced.Click += (s, e) => ToggleAdvanced();
            Controls.Add(_btnAdvanced);
            _chkPaused = new CheckBox { Text = "Pause (send nothing for now)", Left = 160, Top = y + 5, Width = 300, ForeColor = TkeTheme.Gray1, BackColor = TkeTheme.Gray5 };
            Controls.Add(_chkPaused);
            y += 42;

            _pnlAdvanced = new Panel { Left = 0, Top = y, Width = Wd - 16, Height = 250, Visible = false, BackColor = TkeTheme.Gray5 };
            Help(_pnlAdvanced, "Your boss's @name in Work Track (as in the 1:1 meeting):", 14, 3, 430, 20);
            _txtBossPerson = new TextBox { Left = 450, Top = 0, Width = 156, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            _pnlAdvanced.Controls.Add(_txtBossPerson);
            Help(_pnlAdvanced, "Words that mean \"this is an ask\" (one per line):", 14, 34, 290, 20);
            _txtPhrases = new TextBox { Left = 14, Top = 56, Width = 290, Height = 150, Multiline = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            _pnlAdvanced.Controls.Add(_txtPhrases);
            Help(_pnlAdvanced, "…but ignore these (one per line):", 316, 34, 290, 20);
            _txtIgnore = new TextBox { Left = 316, Top = 56, Width = 290, Height = 150, Multiline = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            _pnlAdvanced.Controls.Add(_txtIgnore);
            var btnDefaults = TkeTheme.CreateSecondaryButton("Reset words", 14, 212, 130, 32);
            btnDefaults.Click += (s, e) =>
            {
                _txtPhrases.Text = string.Join(Environment.NewLine, AskDetector.DefaultPhrases);
                _txtIgnore.Text = string.Join(Environment.NewLine, AskDetector.DefaultIgnorePhrases);
            };
            _pnlAdvanced.Controls.Add(btnDefaults);
            var btnTestTask = TkeTheme.CreateSecondaryButton("Send a test task", 156, 212, 170, 32);
            btnTestTask.Click += (s, e) => SendTestTask();
            _pnlAdvanced.Controls.Add(btnTestTask);
            var btnRemove = TkeTheme.CreateSecondaryButton("Remove connection", 438, 212, 168, 32);
            btnRemove.Click += (s, e) => RemoveConnection();
            _pnlAdvanced.Controls.Add(btnRemove);
            Controls.Add(_pnlAdvanced);

            var btnSave = TkeTheme.CreatePrimaryButton("Save", 0, 0, 120, 38);
            var btnCancel = TkeTheme.CreateSecondaryButton("Cancel", 0, 0, 110, 38);
            btnSave.Click += (s, e) => Save();
            btnCancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            btnSave.Tag = "save";
            btnCancel.Tag = "cancel";
            Controls.Add(btnSave);
            Controls.Add(btnCancel);
            AcceptButton = btnSave;
            CancelButton = btnCancel;

            // Load values
            _chkFlags.Checked = _settings.FlagSync;
            _chkAsks.Checked = _settings.AskDetection;
            _chkCalendar.Checked = _settings.CalendarPush;
            _chkPaused.Checked = _code != null && !_settings.Enabled;
            _txtBoss.Text = _settings.BossNames;
            _txtVips.Text = _settings.VipNames;
            _txtBossPerson.Text = _settings.BossPerson;
            _txtPhrases.Text = string.Join(Environment.NewLine, _settings.AskPhrases.ToArray());
            _txtIgnore.Text = string.Join(Environment.NewLine, _settings.IgnorePhrases.ToArray());

            ArrangeButtons(false);
        }

        private int _baseHeight;

        private void ArrangeButtons(bool advanced)
        {
            if (_baseHeight == 0) _baseHeight = _pnlAdvanced.Top;
            int bottom = advanced ? _pnlAdvanced.Bottom + 8 : _baseHeight;
            foreach (Control c in Controls)
            {
                var b = c as Button;
                if (b == null || b.Tag == null) continue;
                b.Top = bottom + 6;
                b.Left = (string)b.Tag == "save" ? Wd - 156 : Wd - 276;
            }
            ClientSize = new Size(Wd - 16, bottom + 58);
        }

        private void ToggleAdvanced()
        {
            _pnlAdvanced.Visible = !_pnlAdvanced.Visible;
            _btnAdvanced.Text = _pnlAdvanced.Visible ? "Advanced ▾" : "Advanced ▸";
            ArrangeButtons(_pnlAdvanced.Visible);
        }

        private void ShowStatus()
        {
            if (_code == null)
            {
                _lblStatus.Text = "Not connected yet — start with step 1.";
                _lblStatus.ForeColor = TkeTheme.Gray1;
            }
            else
            {
                _lblStatus.Text = _settings.Enabled ? "✓ Connected to Work Track" : "✓ Connected (paused)";
                _lblStatus.ForeColor = TkeTheme.Purple;
            }
        }

        // The code typed in, or the saved one when the box is left empty.
        private ConnectorCode CodeInBox(bool quiet)
        {
            string text = _txtCode.Text.Trim();
            if (text.Length == 0) return _code;
            ConnectorCode parsed = null;
            try { parsed = ConnectorCode.Parse(text); } catch { }
            if (parsed == null && !quiet)
                MessageBox.Show(this, "That doesn't look like a Work Track connection code. It starts with \"WT1:\" — copy it again from Work Track → Settings.",
                    "Work Track", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return parsed;
        }

        private void RunInBackground(string working, Func<string> work)
        {
            _lblTestResult.Text = working;
            _lblTestResult.ForeColor = TkeTheme.Gray1;
            Cursor = Cursors.WaitCursor;
            ThreadPool.QueueUserWorkItem(_ =>
            {
                string result;
                bool ok;
                try { result = work(); ok = true; }
                catch (Exception ex) { result = ex.Message; ok = false; }
                BeginInvoke((MethodInvoker)(() =>
                {
                    Cursor = Cursors.Default;
                    _lblTestResult.Text = (ok ? "✓ " : "✕ ") + result;
                    _lblTestResult.ForeColor = ok ? TkeTheme.Purple : TkeTheme.Black;
                }));
            });
        }

        private void TestConnection()
        {
            var code = CodeInBox(false);
            if (code == null)
            {
                if (_txtCode.Text.Trim().Length == 0) _lblTestResult.Text = "Paste your connection code first.";
                return;
            }
            RunInBackground("Checking…", () =>
            {
                new WorkTrackConnector(code).Test();
                return "Works. Click Save to turn it on.";
            });
        }

        private void SendTestTask()
        {
            var code = CodeInBox(false);
            if (code == null) { _lblTestResult.Text = "Connect first (step 1)."; return; }
            RunInBackground("Sending…", () =>
            {
                string id = "ol-test-" + DateTime.Now.ToString("yyyyMMddHHmmss");
                new WorkTrackConnector(code).CreateTask(id, WorkTrackConnector.NewTask(
                    "Test from Outlook — you can drop this", "Sent from the Rules Dashboard.", null, "inbox", null, null, false));
                return "Sent. Look in your Work Track inbox.";
            });
        }

        private void RemoveConnection()
        {
            if (MessageBox.Show(this, "Remove the connection? Nothing more is sent to Work Track until you paste a code again.",
                "Work Track", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            _code = null;
            _txtCode.Text = "";
            _settings.SetCode(null);
            _settings.Enabled = false;
            _settings.Save();
            ShowStatus();
            _lblTestResult.Text = "Connection removed.";
        }

        private static List<string> Lines(string text)
        {
            var list = new List<string>();
            foreach (string l in (text ?? "").Split('\n'))
            {
                string t = l.Trim();
                if (t.Length > 0) list.Add(t);
            }
            return list;
        }

        private void Save()
        {
            var code = CodeInBox(false);
            if (_txtCode.Text.Trim().Length > 0 && code == null) return;
            _code = code;
            _settings.SetCode(code);
            _settings.Enabled = code != null && !_chkPaused.Checked;
            _settings.FlagSync = _chkFlags.Checked;
            _settings.AskDetection = _chkAsks.Checked;
            _settings.CalendarPush = _chkCalendar.Checked;
            _settings.BossNames = _txtBoss.Text.Trim();
            _settings.VipNames = _txtVips.Text.Trim();
            _settings.BossPerson = _txtBossPerson.Text.Trim().Length > 0 ? _txtBossPerson.Text.Trim() : "Boss";
            _settings.AskPhrases = Lines(_txtPhrases.Text);
            _settings.IgnorePhrases = Lines(_txtIgnore.Text);
            _settings.Save();
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
