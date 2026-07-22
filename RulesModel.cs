using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;

namespace OutlookListener
{
    // Data model + on-disk storage for rules, shared by OutlookListener.exe (which
    // evaluates/executes rules against arriving mail) and RulesDashboard.exe (which
    // lets the user add/edit/delete them). Both read/write the same rules.json, so
    // edits made in the dashboard take effect in the running listener without a
    // restart (the listener watches the file for changes).

    public class RuleSet
    {
        public int Version { get; set; }
        public List<Rule> Rules { get; set; }

        public RuleSet()
        {
            Version = 1;
            Rules = new List<Rule>();
        }
    }

    public class Rule
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public bool Enabled { get; set; }
        public RuleConditions Conditions { get; set; }
        public RuleAction Action { get; set; }

        public Rule()
        {
            Id = Guid.NewGuid().ToString("N");
            Name = "New rule";
            Enabled = true;
            Conditions = new RuleConditions();
            Action = new RuleAction();
        }
    }

    // A condition left blank/null means "don't care" - all non-blank conditions
    // on a rule must match (AND) for the rule to fire.
    public class RuleConditions
    {
        public string SenderContains { get; set; }
        public string SubjectContains { get; set; }
        public string BodyContains { get; set; }
        public bool RequiresTable { get; set; }
        public bool? HasAttachment { get; set; }

        public RuleConditions()
        {
            SenderContains = "";
            SubjectContains = "";
            BodyContains = "";
            RequiresTable = false;
            HasAttachment = null;
        }

        public string Summarize()
        {
            var parts = new List<string>();
            if (!string.IsNullOrEmpty(SenderContains)) parts.Add("sender contains \"" + SenderContains + "\"");
            if (!string.IsNullOrEmpty(SubjectContains)) parts.Add("subject contains \"" + SubjectContains + "\"");
            if (!string.IsNullOrEmpty(BodyContains)) parts.Add("body contains \"" + BodyContains + "\"");
            if (RequiresTable) parts.Add("has a table in the body");
            if (HasAttachment.HasValue) parts.Add(HasAttachment.Value ? "has attachment(s)" : "no attachments");
            return parts.Count == 0 ? "(matches every mail)" : string.Join("; ", parts.ToArray());
        }
    }

    public class RuleAction
    {
        public const string ExtractToExcel = "ExtractToExcel";
        public const string RunCommand = "RunCommand";

        public string Type { get; set; }

        // ExtractToExcel fields
        public string ScriptPath { get; set; }
        public string ExcelPath { get; set; }
        public string SheetName { get; set; }

        // RunCommand fields
        public string Command { get; set; }
        public string Arguments { get; set; }

        public RuleAction()
        {
            Type = RunCommand;
            ScriptPath = @"C:\Users\NizamuddinRidhwan\OneDrive - TK Elevator\R1\Visual Studio\Claude\Projects\MsgToExcel\Extract-MsgTable.ps1";
            ExcelPath = @"C:\Users\NizamuddinRidhwan\OneDrive - TK Elevator\Desktop\TEST.xlsx";
            SheetName = "Sheet1";
            Command = "";
            Arguments = "";
        }

        public string Summarize()
        {
            if (Type == ExtractToExcel)
            {
                return "Extract to Excel: " + ExcelPath;
            }
            return "Run: " + Command + " " + Arguments;
        }
    }

    public static class RulesStore
    {
        public static readonly string RulesPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "OutlookListener", "rules.json");

        public static RuleSet Load()
        {
            string dir = Path.GetDirectoryName(RulesPath);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            if (!File.Exists(RulesPath))
            {
                var fresh = new RuleSet();
                Save(fresh);
                return fresh;
            }

            string json = File.ReadAllText(RulesPath);
            if (string.IsNullOrWhiteSpace(json)) return new RuleSet();

            var serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            var loaded = serializer.Deserialize<RuleSet>(json);
            return loaded ?? new RuleSet();
        }

        public static void Save(RuleSet rules)
        {
            string dir = Path.GetDirectoryName(RulesPath);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            var serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = int.MaxValue;
            string json = serializer.Serialize(rules);
            File.WriteAllText(RulesPath, json);
        }
    }
}
