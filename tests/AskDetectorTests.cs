using System;
using OutlookListener;

// Unit tests for AskDetector (no Outlook needed). Built and run by .github/workflows/build.yml:
//   csc /out:AskDetectorTests.exe tests\AskDetectorTests.cs AskDetector.cs && AskDetectorTests.exe
internal static class AskDetectorTests
{
    private static int _failed;

    private static void Check(string name, bool actual, bool expected)
    {
        if (actual == expected) { Console.WriteLine("ok   " + name); return; }
        _failed++;
        Console.WriteLine("FAIL " + name + " (expected " + expected + ")");
    }

    private static void Equal(string name, string actual, string expected)
    {
        if (actual == expected) { Console.WriteLine("ok   " + name); return; }
        _failed++;
        Console.WriteLine("FAIL " + name + ": got \"" + actual + "\", expected \"" + expected + "\"");
    }

    private static bool Ask(string subject, string body)
    {
        string matched;
        return AskDetector.LooksLikeAsk(subject, body, AskDetector.DefaultPhrases, AskDetector.DefaultIgnorePhrases, out matched);
    }

    private static int Main()
    {
        Check("direct request", Ask("Budget numbers", "Hi, can you send the Q3 budget by EOD?"), true);
        Check("action required in subject", Ask("Action required: approve PO 4411", ""), true);
        Check("question", Ask("Quick question", "Is the deck ready"  + "?"), true);
        Check("can u", Ask("Report", "can u check the totals"), true);
        Check("attachment boilerplate is not an ask", Ask("Invoice", "Please find attached the invoice for October."), false);
        Check("quoted history is ignored", Ask("RE: Weekly report", "Thanks!\r\n\r\nFrom: Me\r\nSent: Monday\r\nCan you update the numbers?"), false);
        Check("newsletter", Ask("October newsletter", "Read our latest news. Unsubscribe here."), false);
        Check("plain info", Ask("Lunch", "See you at 1"), false);
        Check("'pleased' is not 'please'", Ask("News", "We are pleased to announce the new office."), false);
        Check("by tomorrow", Ask("Slides", "Send me the slides by tomorrow."), true);

        Equal("clean subject", AskDetector.CleanSubject("RE: FW: [EXTERNAL] Budget"), "Budget");
        Equal("empty subject", AskDetector.CleanSubject("  "), "(no subject)");

        Check("boss by address", AskDetector.NameMatches("Ahmed Ali ahmed.ali@corp.com", "Sara; ahmed.ali@corp.com"), true);
        Check("boss by name", AskDetector.NameMatches("Ahmed Ali ahmed.ali@corp.com", "ahmed ali"), true);
        Check("no list", AskDetector.NameMatches("Ahmed Ali", ""), false);
        Check("not on list", AskDetector.NameMatches("Omar Z omar@corp.com", "Sara; Ahmed"), false);

        Check("teams digest", AskDetector.IsTeamsDigest("Microsoft Teams noreply@email.teams.microsoft.com", "You have 3 missed messages"), true);
        Check("not teams", AskDetector.IsTeamsDigest("Sara sara@corp.com", "You have 3 missed messages"), false);
        Equal("teams title", AskDetector.TeamsTitle("You have 3 missed messages\nSara Khan\nCan you send me the file?\nReply in Teams", "x"), "Teams: Can you send me the file?");

        Check("strip keeps new text", AskDetector.StripQuoted("New text\n-----Original Message-----\nold").Trim() == "New text", true);

        Console.WriteLine(_failed == 0 ? "All AskDetector tests passed." : _failed + " test(s) failed.");
        return _failed == 0 ? 0 : 1;
    }
}
