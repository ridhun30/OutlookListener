using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookListener
{
    // Console app that attaches to Outlook and hooks Inbox.Items.ItemAdd, which fires
    // whenever a new mail lands in the Inbox. Runs a WinForms message loop so the STA
    // thread keeps pumping messages - required for the COM event callback to fire.
    internal static class Program
    {
        private static Outlook.Application _outlookApp;
        private static Outlook.NameSpace _namespace;
        private static Outlook.MAPIFolder _inbox;
        private static Outlook.Items _items;
        private static RuleSet _rules;
        private static FileSystemWatcher _rulesWatcher;
        private static WorkTrackSync _workTrack;
        private static readonly string LogFile = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookListener", "listener.log");

        [STAThread]
        private static void Main()
        {
            Console.WriteLine("Outlook Listener - press Ctrl+C to stop.");
            Console.WriteLine("Log file: " + LogFile);

            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                Console.WriteLine("Stopping...");
                Cleanup();
                Application.Exit();
            };

            try
            {
                LoadRules();
                SetupRulesWatcher();
                AttachToOutlook();
                StartWorkTrack();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + Unwrap(ex));
                return;
            }

            Application.Run(new ApplicationContext());
        }

        private static void LoadRules()
        {
            _rules = RulesStore.Load();
            Console.WriteLine("Rules file: " + RulesStore.RulesPath);
            Console.WriteLine("Loaded " + _rules.Rules.Count + " rule(s), "
                + CountEnabled(_rules) + " enabled.");
        }

        private static int CountEnabled(RuleSet rules)
        {
            int n = 0;
            foreach (var r in rules.Rules) if (r.Enabled) n++;
            return n;
        }

        private static void SetupRulesWatcher()
        {
            string dir = Path.GetDirectoryName(RulesStore.RulesPath);
            _rulesWatcher = new FileSystemWatcher(dir, Path.GetFileName(RulesStore.RulesPath));
            _rulesWatcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size;
            _rulesWatcher.Changed += (s, e) => ReloadRulesSafely();
            _rulesWatcher.Created += (s, e) => ReloadRulesSafely();
            _rulesWatcher.EnableRaisingEvents = true;
        }

        private static void ReloadRulesSafely()
        {
            // The dashboard may still be mid-write when this fires; retry briefly
            // rather than failing on a transient file lock.
            for (int attempt = 0; attempt < 5; attempt++)
            {
                try
                {
                    _rules = RulesStore.Load();
                    string line = "Rules reloaded: " + _rules.Rules.Count + " rule(s), "
                        + CountEnabled(_rules) + " enabled.";
                    Console.WriteLine(line);
                    AppendToLogFile(line);
                    return;
                }
                catch
                {
                    System.Threading.Thread.Sleep(150);
                }
            }
        }

        private static void AttachToOutlook()
        {
            try
            {
                _outlookApp = (Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                Console.WriteLine("Attached to running Outlook instance.");
            }
            catch (COMException)
            {
                _outlookApp = new Outlook.Application();
                Console.WriteLine("Started a new Outlook instance.");
            }

            _namespace = _outlookApp.GetNamespace("MAPI");
            _inbox = _namespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Console.WriteLine("Watching folder: " + _inbox.Name);

            _items = _inbox.Items;
            _items.ItemAdd += Items_ItemAdd;

            Console.WriteLine("Listening for new mail...");
        }

        // Optional link to Work Track (flagged mail, asks, calendar) - see WorkTrackSync.cs.
        // A failure here never stops the rules from running.
        private static void StartWorkTrack()
        {
            try
            {
                _workTrack = new WorkTrackSync(_namespace, _items, line =>
                {
                    Console.WriteLine(line);
                    AppendToLogFile(line);
                });
                _workTrack.Start();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Work Track link not started: " + Unwrap(ex));
                _workTrack = null;
            }
        }

        private static void Items_ItemAdd(object item)
        {
            try
            {
                var mail = item as Outlook.MailItem;
                string subject = mail != null ? mail.Subject : "(non-mail item)";
                string sender = mail != null ? mail.SenderName : "(unknown sender)";
                string line = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] New mail from {1}: {2}",
                    DateTime.Now, sender, subject);

                Console.WriteLine(line);
                AppendToLogFile(line);

                if (mail != null) ApplyRules(mail);
                if (mail != null && _workTrack != null)
                {
                    try { _workTrack.OnNewMail(mail); }
                    catch (Exception ex) { Console.WriteLine("[Work Track] " + Unwrap(ex)); }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR handling new item: " + Unwrap(ex));
            }
        }

        private static void ApplyRules(Outlook.MailItem mail)
        {
            foreach (var rule in _rules.Rules)
            {
                if (!rule.Enabled) continue;

                bool matched;
                try
                {
                    matched = RulesEngine.Matches(rule, mail);
                }
                catch (Exception ex)
                {
                    string err = "  [ERROR evaluating rule '" + rule.Name + "']: " + Unwrap(ex);
                    Console.WriteLine(err);
                    AppendToLogFile(err);
                    continue;
                }

                if (!matched) continue;

                string matchLine = "  -> rule matched: " + rule.Name + " (" + rule.Action.Summarize() + ")";
                Console.WriteLine(matchLine);
                AppendToLogFile(matchLine);

                try
                {
                    RulesEngine.Execute(rule, mail, msg =>
                    {
                        Console.WriteLine(msg);
                        AppendToLogFile(msg);
                    });
                }
                catch (Exception ex)
                {
                    string err = "  [ERROR running rule '" + rule.Name + "']: " + Unwrap(ex);
                    Console.WriteLine(err);
                    AppendToLogFile(err);
                }
            }
        }

        private static void AppendToLogFile(string line)
        {
            try
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(LogFile));
                System.IO.File.AppendAllText(LogFile, line + Environment.NewLine);
            }
            catch
            {
                // Logging is best-effort; never let a log write failure crash the listener.
            }
        }

        private static void Cleanup()
        {
            try
            {
                if (_rulesWatcher != null) { _rulesWatcher.EnableRaisingEvents = false; _rulesWatcher.Dispose(); }
                if (_workTrack != null) _workTrack.Stop();
                if (_items != null) { _items.ItemAdd -= Items_ItemAdd; Marshal.ReleaseComObject(_items); }
                if (_inbox != null) Marshal.ReleaseComObject(_inbox);
                if (_namespace != null) Marshal.ReleaseComObject(_namespace);
                // Deliberately not releasing/quitting _outlookApp: if we attached to the
                // user's already-running Outlook, we must not close their session.
            }
            catch
            {
            }
        }

        private static string Unwrap(Exception ex)
        {
            Exception cur = ex;
            while (cur.InnerException != null) cur = cur.InnerException;
            return cur.Message;
        }
    }
}
