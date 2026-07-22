using System;
using System.Diagnostics;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookListener
{
    // Evaluates a rule's conditions against an arrived MailItem, and executes a
    // matched rule's action. Only referenced by OutlookListener.exe (the listener) -
    // the dashboard only edits RuleSet/Rule data, it never touches live mail.
    public static class RulesEngine
    {
        public static bool Matches(Rule rule, Outlook.MailItem mail)
        {
            var c = rule.Conditions;

            if (!string.IsNullOrEmpty(c.SenderContains))
            {
                string sender = SafeGet(() => mail.SenderName) + " " + SafeGet(() => mail.SenderEmailAddress);
                if (sender.IndexOf(c.SenderContains, StringComparison.OrdinalIgnoreCase) < 0) return false;
            }

            if (!string.IsNullOrEmpty(c.SubjectContains))
            {
                string subject = SafeGet(() => mail.Subject);
                if (subject.IndexOf(c.SubjectContains, StringComparison.OrdinalIgnoreCase) < 0) return false;
            }

            if (!string.IsNullOrEmpty(c.BodyContains))
            {
                string body = SafeGet(() => mail.Body);
                if (body.IndexOf(c.BodyContains, StringComparison.OrdinalIgnoreCase) < 0) return false;
            }

            if (c.RequiresTable)
            {
                string html = SafeGet(() => mail.HTMLBody);
                if (html.IndexOf("<table", StringComparison.OrdinalIgnoreCase) < 0) return false;
            }

            if (c.HasAttachment.HasValue)
            {
                bool has = mail.Attachments != null && mail.Attachments.Count > 0;
                if (has != c.HasAttachment.Value) return false;
            }

            return true;
        }

        public static void Execute(Rule rule, Outlook.MailItem mail, Action<string> log)
        {
            string tempMsgPath = null;
            try
            {
                if (rule.Action.Type == RuleAction.ExtractToExcel)
                {
                    tempMsgPath = SaveTempMsg(mail);
                    RunExtractToExcel(rule.Action, tempMsgPath, log);
                }
                else if (rule.Action.Type == RuleAction.RunCommand)
                {
                    if (UsesMsgPathPlaceholder(rule.Action)) tempMsgPath = SaveTempMsg(mail);
                    RunCustomCommand(rule.Action, mail, tempMsgPath, log);
                }
                else
                {
                    log("  [ERROR] unknown action type: " + rule.Action.Type);
                }
            }
            finally
            {
                if (tempMsgPath != null)
                {
                    try { File.Delete(tempMsgPath); } catch { }
                }
            }
        }

        private static bool UsesMsgPathPlaceholder(RuleAction action)
        {
            return (action.Command ?? "").IndexOf("{msgPath}", StringComparison.OrdinalIgnoreCase) >= 0
                || (action.Arguments ?? "").IndexOf("{msgPath}", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string SaveTempMsg(Outlook.MailItem mail)
        {
            string dir = Path.Combine(Path.GetTempPath(), "OutlookListener");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, Guid.NewGuid().ToString("N") + ".msg");
            mail.SaveAs(path, Outlook.OlSaveAsType.olMSG);
            return path;
        }

        private static void RunExtractToExcel(RuleAction action, string msgPath, Action<string> log)
        {
            if (!File.Exists(action.ScriptPath))
            {
                log("  [ERROR] Extract-MsgTable.ps1 not found at: " + action.ScriptPath);
                return;
            }

            string args = string.Format(
                "-NoProfile -ExecutionPolicy Bypass -File \"{0}\" -MsgPath \"{1}\" -ExcelPath \"{2}\" -SheetName \"{3}\"",
                action.ScriptPath, msgPath, action.ExcelPath, action.SheetName);
            RunProcess("powershell.exe", args, log);
        }

        private static void RunCustomCommand(RuleAction action, Outlook.MailItem mail, string msgPath, Action<string> log)
        {
            string command = Substitute(action.Command, mail, msgPath);
            string args = Substitute(action.Arguments, mail, msgPath);
            RunProcess(command, args, log);
        }

        private static string Substitute(string template, Outlook.MailItem mail, string msgPath)
        {
            string s = template ?? "";
            s = s.Replace("{subject}", SafeGet(() => mail.Subject));
            s = s.Replace("{sender}", SafeGet(() => mail.SenderName));
            s = s.Replace("{senderEmail}", SafeGet(() => mail.SenderEmailAddress));
            s = s.Replace("{receivedTime}", SafeGet(() => mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")));
            s = s.Replace("{msgPath}", msgPath ?? "");
            return s;
        }

        private static void RunProcess(string exe, string args, Action<string> log)
        {
            var psi = new ProcessStartInfo(exe, args)
            {
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (var p = Process.Start(psi))
            {
                string stdout = p.StandardOutput.ReadToEnd();
                string stderr = p.StandardError.ReadToEnd();
                bool exited = p.WaitForExit(30000);
                if (!exited)
                {
                    log("  [WARNING] action did not finish within 30s, letting it continue in the background.");
                    return;
                }
                if (!string.IsNullOrWhiteSpace(stdout)) log("  [output] " + stdout.Trim());
                if (!string.IsNullOrWhiteSpace(stderr)) log("  [stderr] " + stderr.Trim());
                if (p.ExitCode != 0) log("  [WARNING] action exited with code " + p.ExitCode);
            }
        }

        private static string SafeGet(Func<string> getter)
        {
            try { return getter() ?? ""; } catch { return ""; }
        }
    }
}
