using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace OutlookListener
{
    // Settings for the Work Track connector (https://github.com/ridhun30/work-track), shared
    // by OutlookListener.exe (which acts on them) and RulesDashboard.exe (which edits them).
    // Stored next to rules.json; the connector code is a credential, so it is encrypted
    // with DPAPI for the current Windows user and never written in plain text.
    public class WorkTrackSettings
    {
        public bool Enabled { get; set; }
        public bool FlagSync { get; set; }
        public bool AskDetection { get; set; }
        public bool CalendarPush { get; set; }

        // Show Work Track's reminders (morning plan, nags, follow-ups, …) as Windows notifications.
        public bool DesktopReminders { get; set; }

        // DPAPI-protected, base64. Use GetCode()/SetCode().
        public string ProtectedCode { get; set; }

        // One phrase per entry, matched case-insensitively in the subject and new part of the body.
        public List<string> AskPhrases { get; set; }
        public List<string> IgnorePhrases { get; set; }

        // Semicolon-separated names or addresses. Mail from these always counts as an ask.
        public string BossNames { get; set; }
        public string VipNames { get; set; }

        // The @person Work Track uses for the boss (must match the 1:1 meeting's person).
        public string BossPerson { get; set; }

        public WorkTrackSettings()
        {
            Enabled = false;
            FlagSync = true;
            AskDetection = true;
            CalendarPush = true;
            DesktopReminders = true;
            ProtectedCode = "";
            AskPhrases = new List<string>(AskDetector.DefaultPhrases);
            IgnorePhrases = new List<string>(AskDetector.DefaultIgnorePhrases);
            BossNames = "";
            VipNames = "";
            BossPerson = "Boss";
        }

        public static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookListener", "worktrack.json");

        public static WorkTrackSettings Load()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return new WorkTrackSettings();
                var loaded = new JavaScriptSerializer().Deserialize<WorkTrackSettings>(File.ReadAllText(SettingsPath));
                return loaded ?? new WorkTrackSettings();
            }
            catch
            {
                return new WorkTrackSettings();
            }
        }

        public void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath));
            File.WriteAllText(SettingsPath, new JavaScriptSerializer().Serialize(this));
        }

        public ConnectorCode GetCode()
        {
            if (string.IsNullOrEmpty(ProtectedCode)) return null;
            try
            {
                byte[] plain = ProtectedData.Unprotect(Convert.FromBase64String(ProtectedCode), null, DataProtectionScope.CurrentUser);
                return ConnectorCode.Parse(Encoding.UTF8.GetString(plain));
            }
            catch
            {
                return null;
            }
        }

        public void SetCode(ConnectorCode code)
        {
            if (code == null) { ProtectedCode = ""; return; }
            byte[] plain = Encoding.UTF8.GetBytes(code.ToText());
            ProtectedCode = Convert.ToBase64String(ProtectedData.Protect(plain, null, DataProtectionScope.CurrentUser));
        }
    }

    // What the Work Track web app's "Connect laptop (Outlook)" button produces:
    // "WT1:" + base64(JSON {v, apiKey, projectId, uid, refreshToken}).
    public class ConnectorCode
    {
        public string ApiKey { get; set; }
        public string ProjectId { get; set; }
        public string Uid { get; set; }
        public string RefreshToken { get; set; }

        // Address of the Work Track web app (opened when a reminder is clicked).
        public string AppUrl { get; set; }

        public const string DefaultAppUrl = "https://ridhun30.github.io/work-track/";

        public static ConnectorCode Parse(string text)
        {
            if (text == null) return null;
            text = text.Trim();
            if (!text.StartsWith("WT1:", StringComparison.Ordinal)) return null;
            string json = Encoding.UTF8.GetString(Convert.FromBase64String(text.Substring(4)));
            var d = new JavaScriptSerializer().Deserialize<Dictionary<string, object>>(json);
            var code = new ConnectorCode
            {
                ApiKey = Str(d, "apiKey"),
                ProjectId = Str(d, "projectId"),
                Uid = Str(d, "uid"),
                RefreshToken = Str(d, "refreshToken"),
                AppUrl = Str(d, "appUrl")
            };
            if (string.IsNullOrEmpty(code.AppUrl)) code.AppUrl = DefaultAppUrl;
            if (string.IsNullOrEmpty(code.ApiKey) || string.IsNullOrEmpty(code.ProjectId)
                || string.IsNullOrEmpty(code.Uid) || string.IsNullOrEmpty(code.RefreshToken)) return null;
            return code;
        }

        public string ToText()
        {
            var d = new Dictionary<string, object>
            {
                { "v", 1 }, { "apiKey", ApiKey }, { "projectId", ProjectId }, { "uid", Uid }, { "refreshToken", RefreshToken },
                { "appUrl", AppUrl ?? DefaultAppUrl }
            };
            return "WT1:" + Convert.ToBase64String(Encoding.UTF8.GetBytes(new JavaScriptSerializer().Serialize(d)));
        }

        private static string Str(Dictionary<string, object> d, string key)
        {
            object v;
            return d.TryGetValue(key, out v) && v != null ? v.ToString() : "";
        }
    }
}
