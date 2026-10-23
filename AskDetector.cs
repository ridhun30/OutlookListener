using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OutlookListener
{
    // Decides whether a mail "looks like an ask" (something someone wants you to do).
    // Pure text logic with no Outlook references, so it is unit-tested on its own
    // (tests\AskDetectorTests.cs, run by the GitHub Actions build).
    public static class AskDetector
    {
        public static readonly string[] DefaultPhrases =
        {
            "can you", "could you", "would you", "can u", "please", "pls", "kindly",
            "action required", "action needed", "need you to", "need your", "your input",
            "your approval", "your feedback", "please review", "please confirm", "please share",
            "by eod", "by end of day", "by eow", "by end of week", "by tomorrow", "asap",
            "deadline", "due by", "awaiting your", "waiting for your", "?"
        };

        // Boilerplate that contains an ask phrase but isn't an ask.
        public static readonly string[] DefaultIgnorePhrases =
        {
            "please find attached", "please see attached", "please find below", "please find the attached",
            "please do not reply", "please don't reply", "please consider the environment",
            "unsubscribe", "newsletter", "automatic reply", "out of office"
        };

        private const int BodyChars = 800;

        // Lines that start the quoted history of a reply/forward.
        private static readonly Regex QuoteStart = new Regex(
            @"^\s*(-{3,}\s*Original Message\s*-{3,}|_{10,}|From:\s.+|Sent:\s.+|On .+ wrote:)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Multiline);

        private static readonly Regex SubjectPrefix = new Regex(
            @"^\s*((re|fw|fwd|aw|wg)\s*:\s*|\[external\]\s*|external\s*:\s*)+", RegexOptions.IgnoreCase);

        /// <summary>The new part of a body: everything before the quoted reply history.</summary>
        public static string StripQuoted(string body)
        {
            if (string.IsNullOrEmpty(body)) return "";
            var m = QuoteStart.Match(body);
            string fresh = m.Success ? body.Substring(0, m.Index) : body;
            return fresh.Length > BodyChars ? fresh.Substring(0, BodyChars) : fresh;
        }

        public static string CleanSubject(string subject)
        {
            string s = SubjectPrefix.Replace(subject ?? "", "").Trim();
            return s.Length == 0 ? "(no subject)" : s;
        }

        /// <summary>
        /// True when the subject or new body contains an ask phrase (and no ignore phrase
        /// explains it away). <paramref name="matched"/> is the phrase that triggered it.
        /// </summary>
        public static bool LooksLikeAsk(string subject, string body, IList<string> phrases, IList<string> ignore, out string matched)
        {
            matched = null;
            string text = ((subject ?? "") + "\n" + StripQuoted(body)).ToLowerInvariant();
            foreach (string ig in ignore ?? new string[0])
            {
                if (string.IsNullOrEmpty(ig)) continue;
                text = text.Replace(ig.ToLowerInvariant(), " ");
            }
            foreach (string p in phrases ?? new string[0])
            {
                if (string.IsNullOrEmpty(p)) continue;
                string phrase = p.Trim().ToLowerInvariant();
                if (phrase.Length == 0) continue;
                bool hit = phrase == "?"
                    ? text.IndexOf('?') >= 0
                    : Regex.IsMatch(text, @"(^|[^a-z0-9])" + Regex.Escape(phrase) + @"($|[^a-z0-9])");
                if (hit)
                {
                    matched = p.Trim();
                    return true;
                }
            }
            return false;
        }

        /// <summary>True when "Name <address>" matches any entry of a ;-or-,-separated list.</summary>
        public static bool NameMatches(string senderNameAndEmail, string list)
        {
            if (string.IsNullOrEmpty(senderNameAndEmail) || string.IsNullOrEmpty(list)) return false;
            foreach (string raw in list.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string n = raw.Trim();
                if (n.Length > 1 && senderNameAndEmail.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }

        /// <summary>Teams' "you have missed messages" notification emails.</summary>
        public static bool IsTeamsDigest(string senderNameAndEmail, string subject)
        {
            string s = senderNameAndEmail ?? "";
            bool fromTeams = s.IndexOf("teams.microsoft.com", StringComparison.OrdinalIgnoreCase) >= 0
                || s.IndexOf("Microsoft Teams", StringComparison.OrdinalIgnoreCase) >= 0;
            return fromTeams && Regex.IsMatch(subject ?? "", @"missed|new message|sent you a message|mentioned you", RegexOptions.IgnoreCase);
        }

        /// <summary>A short title for a Teams digest: the first sentence-like line of its body.</summary>
        public static string TeamsTitle(string body, string fallback)
        {
            string firstShort = null;
            foreach (string line in (body ?? "").Split('\n'))
            {
                string l = line.Trim();
                if (l.Length < 8 || l.Length > 200) continue;
                if (Regex.IsMatch(l, @"^(you have|go to teams|reply in teams|microsoft teams|unsubscribe|install|download|privacy|\W+$)", RegexOptions.IgnoreCase)) continue;
                if (l.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length >= 4) return "Teams: " + l;
                if (firstShort == null) firstShort = l;
            }
            return "Teams: " + (firstShort ?? CleanSubject(fallback));
        }
    }
}
