using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace OutlookListener
{
    internal static class DashboardProgram
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private DataGridView _grid;
        private RuleSet _ruleSet;

        public MainForm()
        {
            Text = "OutlookListener - Rules Dashboard";
            Width = 900;
            Height = 590;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = TkeTheme.Gray5;
            ForeColor = TkeTheme.Gray1;
            Font = TkeTheme.Body(10F);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Padding = new Padding(0);

            BuildUi();
            LoadRules();
        }

        private void BuildUi()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 90,
                BackColor = TkeTheme.Black
            };
            header.Paint += (s, e) =>
            {
                TkeTheme.PaintHeader(e.Graphics, header.ClientRectangle,
                    "OutlookListener", 18F, 18,
                    "Rules Dashboard", 9.5F, 50);
            };
            Controls.Add(header);

            _grid = new DataGridView
            {
                Left = 12,
                Top = 104,
                Width = 866,
                Height = 372,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                BackgroundColor = TkeTheme.Gray5,
                BorderStyle = BorderStyle.None,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(218, 218, 218),
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                RowTemplate = { Height = 30 }
            };

            // Table Header (per the Kit's Tables page, p43): plain black text with a thin
            // underline, no fill by default - purple fill is reserved for a hovered or
            // currently-sorted column, neither of which applies to this static grid.
            _grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.White,
                ForeColor = TkeTheme.Black,
                Font = TkeTheme.BoldUnderline(9.5F),
                Alignment = DataGridViewContentAlignment.MiddleLeft,
                Padding = new Padding(2, 0, 0, 0)
            };
            _grid.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = Color.White,
                ForeColor = TkeTheme.Gray1,
                SelectionBackColor = TkeTheme.Purple,
                SelectionForeColor = Color.White,
                Font = TkeTheme.Body(9.5F),
                Padding = new Padding(4, 0, 0, 0)
            };

            var colEnabled = new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "On", Width = 44 };
            var colName = new DataGridViewTextBoxColumn { Name = "Name", HeaderText = "Rule", Width = 160, ReadOnly = true };
            var colConditions = new DataGridViewTextBoxColumn { Name = "Conditions", HeaderText = "When", Width = 310, ReadOnly = true };
            var colAction = new DataGridViewTextBoxColumn { Name = "Action", HeaderText = "Then", Width = 280, ReadOnly = true };
            _grid.Columns.AddRange(new DataGridViewColumn[] { colEnabled, colName, colConditions, colAction });

            _grid.CellValueChanged += Grid_CellValueChanged;
            _grid.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (_grid.IsCurrentCellDirty) _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };
            _grid.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) EditSelected(); };

            Controls.Add(_grid);

            int buttonTop = 490;
            var btnAdd = TkeTheme.CreatePrimaryButton("Add Rule", 12, buttonTop, 104, 38);
            var btnEdit = TkeTheme.CreateSecondaryButton("Edit", 128, buttonTop, 98, 38);
            var btnDelete = TkeTheme.CreateSecondaryButton("Delete", 238, buttonTop, 100, 38);
            var btnRefresh = TkeTheme.CreateSecondaryButton("Refresh", 350, buttonTop, 100, 38);
            var btnOpenLog = TkeTheme.CreateSecondaryButton("View Log", 660, buttonTop, 110, 38);
            var btnOpenRules = TkeTheme.CreateSecondaryButton("Open rules.json", 472, buttonTop, 176, 38);

            foreach (var btn in new[] { btnAdd, btnEdit, btnDelete, btnRefresh, btnOpenLog, btnOpenRules })
            {
                btn.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            }

            btnAdd.Click += (s, e) => AddRule();
            btnEdit.Click += (s, e) => EditSelected();
            btnDelete.Click += (s, e) => DeleteSelected();
            btnRefresh.Click += (s, e) => LoadRules();
            btnOpenLog.Click += (s, e) => OpenLog();
            btnOpenRules.Click += (s, e) => OpenRulesFile();

            Controls.Add(btnAdd);
            Controls.Add(btnEdit);
            Controls.Add(btnDelete);
            Controls.Add(btnRefresh);
            Controls.Add(btnOpenLog);
            Controls.Add(btnOpenRules);
        }

        private void LoadRules()
        {
            _ruleSet = RulesStore.Load();
            RefreshGrid();
        }

        private void RefreshGrid()
        {
            _grid.CellValueChanged -= Grid_CellValueChanged;
            _grid.Rows.Clear();
            foreach (var rule in _ruleSet.Rules)
            {
                int idx = _grid.Rows.Add(rule.Enabled, rule.Name, rule.Conditions.Summarize(), rule.Action.Summarize());
                _grid.Rows[idx].Tag = rule;
            }
            _grid.CellValueChanged += Grid_CellValueChanged;
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _grid.Rows.Count) return;
            if (_grid.Columns[e.ColumnIndex].Name != "Enabled") return;

            var rule = _grid.Rows[e.RowIndex].Tag as Rule;
            if (rule == null) return;

            rule.Enabled = (bool)_grid.Rows[e.RowIndex].Cells["Enabled"].Value;
            RulesStore.Save(_ruleSet);
        }

        private Rule GetSelectedRule()
        {
            if (_grid.CurrentRow == null) return null;
            return _grid.CurrentRow.Tag as Rule;
        }

        private void AddRule()
        {
            var rule = new Rule();
            using (var editor = new RuleEditorForm(rule))
            {
                if (editor.ShowDialog(this) == DialogResult.OK)
                {
                    _ruleSet.Rules.Add(rule);
                    RulesStore.Save(_ruleSet);
                    RefreshGrid();
                }
            }
        }

        private void EditSelected()
        {
            var rule = GetSelectedRule();
            if (rule == null)
            {
                MessageBox.Show(this, "Select a rule first.", "Rules Dashboard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var editor = new RuleEditorForm(rule))
            {
                if (editor.ShowDialog(this) == DialogResult.OK)
                {
                    RulesStore.Save(_ruleSet);
                    RefreshGrid();
                }
            }
        }

        private void DeleteSelected()
        {
            var rule = GetSelectedRule();
            if (rule == null)
            {
                MessageBox.Show(this, "Select a rule first.", "Rules Dashboard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirm = MessageBox.Show(this, "Delete rule \"" + rule.Name + "\"?", "Rules Dashboard",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            _ruleSet.Rules.Remove(rule);
            RulesStore.Save(_ruleSet);
            RefreshGrid();
        }

        private void OpenLog()
        {
            string logFile = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookListener", "listener.log");

            if (!File.Exists(logFile))
            {
                MessageBox.Show(this, "No log file yet - the listener writes one once it's caught its first mail.\n\n" + logFile,
                    "Rules Dashboard", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Process.Start("notepad.exe", "\"" + logFile + "\"");
        }

        private void OpenRulesFile()
        {
            Process.Start("notepad.exe", "\"" + RulesStore.RulesPath + "\"");
        }
    }
}
