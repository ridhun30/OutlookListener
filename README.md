# OutlookListener

Two pieces that share a rules file:

- **`OutlookListener.exe`** - a console app that keeps running in the background,
  watches your Outlook Inbox, and runs your rules against every new mail.
- **`RulesDashboard.exe`** - a GUI for adding/editing/deleting the rules. Changes
  take effect in the listener immediately, without restarting it.

## Usage

1. Build both: `.\build.ps1` (compiles via `csc.exe` - ships with .NET Framework,
   no separate install; also stops any already-running instances first since a
   running exe can't be overwritten).
2. Run `RulesDashboard.exe` to set up your rules (see below).
3. Run `OutlookListener.exe` and leave the window open - it needs to keep running
   to receive new-mail events. Press `Ctrl+C` to stop it.

Every new mail in the Inbox logs a line regardless of rules, e.g.:

```
[2026-07-22 15:22:59] New mail from Jane Doe: Quarterly report
  -> rule matched: Invoices to Excel (Extract to Excel: C:\...\TEST.xlsx)
  [output] Appended 3 row(s) to 'C:\...\TEST.xlsx' (sheet 'Sheet1').
```

Log file: `%LOCALAPPDATA%\OutlookListener\listener.log`
Rules file: `%LOCALAPPDATA%\OutlookListener\rules.json`

## Rules

A rule has **conditions** (the "context") and one **action**. All filled-in
conditions must match (AND) for the action to run; a blank/unset condition is
ignored.

Conditions:
- **Sender contains** - substring match against sender display name + email address
- **Subject contains** - substring match
- **Body contains** - substring match against the plain-text body
- **Body must contain an HTML table** - useful for gating table-extraction rules
- **Attachments** - any / must have / must not have

Actions:
- **Extract to Excel** - runs MsgToExcel's `Extract-MsgTable.ps1` against the
  mail (saved to a temp `.msg` file first, deleted after) with the script path,
  target workbook, and sheet name you choose.
- **Run a command/script** - runs any exe/script with arguments you write,
  supporting placeholders: `{subject}` `{sender}` `{senderEmail}`
  `{receivedTime}` `{msgPath}` (a temp `.msg` copy of the mail, only saved if
  you actually reference `{msgPath}`).

Use "View Log" / "Open rules.json" in the dashboard to inspect either file
directly (opens in Notepad).

## How it works

- Uses `Microsoft.Office.Interop.Outlook` (already present in this machine's GAC,
  installed alongside Office) to attach to Outlook and hook `Inbox.Items.ItemAdd`.
- Runs a WinForms message loop (`Application.Run`) on the listener's main STA
  thread - required for COM to actually deliver the event callback.
- Rules are stored as JSON (`System.Web.Script.Serialization.JavaScriptSerializer`
  - built into .NET Framework, no NuGet) and shared by both exes via `RulesModel.cs`.
  The listener watches `rules.json` with a `FileSystemWatcher` and reloads on any
  change, so edits made in the dashboard apply live.
- Deliberately does **not** call `Quit()` on the Outlook application object on
  exit, since it may have attached to your already-running Outlook session
  rather than starting its own.
- Rule actions run via `Process.Start` with output captured and appended to the
  same log the listener writes to, so you can see what a triggered action
  actually did/printed.

## Limitations

- Only watches the default Inbox folder. To watch a different folder, change
  the `Outlook.OlDefaultFolders.olFolderInbox` argument in `AttachToOutlook()`
  in `OutlookListener.cs`.
- Requires the console window to stay open (or the process to keep running
  some other way, e.g. as a scheduled task at logon) - closing it stops the
  listener.
- Attaches to whatever Outlook profile/mailbox is already signed in; doesn't
  support watching a shared mailbox differently from the primary one.
- `{sender}`/`{subject}` values are substituted into command-line arguments as
  raw text - if a mail's subject/sender contains a `"`, it can break the quoting
  of whatever command you configured. Not sanitized in v1.
- A rule with every condition left blank matches **every** mail - the dashboard
  doesn't warn about this currently.
