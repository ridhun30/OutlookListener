using System;
using System.Drawing;
using System.Windows.Forms;

namespace OutlookListener
{
    // Modal add/edit dialog for a single rule. Writes changes directly back into
    // the Rule instance passed in when the user clicks OK; caller is responsible
    // for adding it to the RuleSet (new rules) and persisting via RulesStore.Save.
    public class RuleEditorForm : Form
    {
        private readonly Rule _rule;

        private TextBox _txtName;
        private CheckBox _chkEnabled;

        private TextBox _txtSender;
        private TextBox _txtSubject;
        private TextBox _txtBody;
        private CheckBox _chkRequiresTable;
        private ComboBox _cmbAttachment;

        private RadioButton _rbExtract;
        private RadioButton _rbCommand;
        private Panel _pnlExtract;
        private Panel _pnlCommand;

        private TextBox _txtScriptPath;
        private TextBox _txtExcelPath;
        private TextBox _txtSheetName;

        private TextBox _txtCommand;
        private TextBox _txtArguments;

        public RuleEditorForm(Rule rule)
        {
            _rule = rule;

            Text = "Rule";
            Width = 620;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = TkeTheme.Gray5;
            ForeColor = TkeTheme.Gray1;
            Font = TkeTheme.Body(10F);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            BuildUi();
            LoadFromRule();
        }

        private void BuildUi()
        {
            var header = new Panel
            {
                Dock = DockStyle.Top,
                Height = 72,
                BackColor = TkeTheme.Black
            };
            header.Paint += (s, e) =>
            {
                TkeTheme.PaintHeader(e.Graphics, header.ClientRectangle,
                    "Rule", 16F, 18,
                    "Edit the outbound rule definition", 9.5F, 46);
            };
            Controls.Add(header);

            int y = 88;

            Controls.Add(new Label { Text = "Name:", Left = 12, Top = y + 3, Width = 60, ForeColor = TkeTheme.Gray1 });
            _txtName = new TextBox { Left = 82, Top = y, Width = 364, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            Controls.Add(_txtName);
            _chkEnabled = new CheckBox { Text = "Enabled", Left = 462, Top = y, Width = 90, ForeColor = TkeTheme.Gray1, BackColor = TkeTheme.Gray5 };
            Controls.Add(_chkEnabled);
            y += 36;

            var grpConditions = new GroupBox { Text = "Conditions (blank = don't care; all filled-in ones must match)", Left = 10, Top = y, Width = 580, Height = 182, ForeColor = TkeTheme.Gray1, BackColor = TkeTheme.Gray5 };
            int cy = 25;
            grpConditions.Controls.Add(new Label { Text = "Sender contains:", Left = 12, Top = cy + 3, Width = 120, ForeColor = TkeTheme.Gray1 });
            _txtSender = new TextBox { Left = 142, Top = cy, Width = 400, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            grpConditions.Controls.Add(_txtSender);
            cy += 30;

            grpConditions.Controls.Add(new Label { Text = "Subject contains:", Left = 12, Top = cy + 3, Width = 120, ForeColor = TkeTheme.Gray1 });
            _txtSubject = new TextBox { Left = 142, Top = cy, Width = 400, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            grpConditions.Controls.Add(_txtSubject);
            cy += 30;

            grpConditions.Controls.Add(new Label { Text = "Body contains:", Left = 12, Top = cy + 3, Width = 120, ForeColor = TkeTheme.Gray1 });
            _txtBody = new TextBox { Left = 142, Top = cy, Width = 400, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            grpConditions.Controls.Add(_txtBody);
            cy += 30;

            _chkRequiresTable = new CheckBox { Text = "Body must contain an HTML table", Left = 142, Top = cy, Width = 300, ForeColor = TkeTheme.Gray1, BackColor = TkeTheme.Gray5 };
            grpConditions.Controls.Add(_chkRequiresTable);
            cy += 32;

            grpConditions.Controls.Add(new Label { Text = "Attachments:", Left = 12, Top = cy + 3, Width = 120, ForeColor = TkeTheme.Gray1 });
            _cmbAttachment = new ComboBox { Left = 142, Top = cy, Width = 250, DropDownStyle = ComboBoxStyle.DropDownList, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            _cmbAttachment.Items.Add("Any (don't care)");
            _cmbAttachment.Items.Add("Must have attachment(s)");
            _cmbAttachment.Items.Add("Must NOT have attachment(s)");
            grpConditions.Controls.Add(_cmbAttachment);

            Controls.Add(grpConditions);
            y += grpConditions.Height + 10;

            var grpAction = new GroupBox { Text = "Action", Left = 10, Top = y, Width = 580, Height = 330, ForeColor = TkeTheme.Gray1, BackColor = TkeTheme.Gray5 };
            int ay = 25;

            _rbExtract = new RadioButton { Text = "Extract to Excel (MsgToExcel)", Left = 10, Top = ay, Width = 300, ForeColor = TkeTheme.Gray1, BackColor = TkeTheme.Gray5 };
            _rbCommand = new RadioButton { Text = "Run a command / script", Left = 10, Top = ay + 24, Width = 300, ForeColor = TkeTheme.Gray1, BackColor = TkeTheme.Gray5 };
            _rbExtract.CheckedChanged += (s, e) => UpdateActionPanels();
            grpAction.Controls.Add(_rbExtract);
            grpAction.Controls.Add(_rbCommand);
            ay += 56;

            _pnlExtract = new Panel { Left = 10, Top = ay, Width = 555, Height = 110, BackColor = TkeTheme.Gray5 };
            int ey = 0;
            _pnlExtract.Controls.Add(new Label { Text = "Extract-MsgTable.ps1:", Left = 0, Top = ey + 3, Width = 140, ForeColor = TkeTheme.Gray1 });
            _txtScriptPath = new TextBox { Left = 145, Top = ey, Width = 300, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            var btnBrowseScript = TkeTheme.CreateSecondaryButton("...", 30, 24);
            btnBrowseScript.Left = 450;
            btnBrowseScript.Top = ey - 1;
            btnBrowseScript.Click += (s, e) => BrowseFile(_txtScriptPath, "PowerShell script (*.ps1)|*.ps1|All files (*.*)|*.*");
            _pnlExtract.Controls.Add(_txtScriptPath);
            _pnlExtract.Controls.Add(btnBrowseScript);
            ey += 30;

            _pnlExtract.Controls.Add(new Label { Text = "Target Excel file:", Left = 0, Top = ey + 3, Width = 140, ForeColor = TkeTheme.Gray1 });
            _txtExcelPath = new TextBox { Left = 145, Top = ey, Width = 300, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            var btnBrowseExcel = TkeTheme.CreateSecondaryButton("...", 30, 24);
            btnBrowseExcel.Left = 450;
            btnBrowseExcel.Top = ey - 1;
            btnBrowseExcel.Click += (s, e) => BrowseFile(_txtExcelPath, "Excel workbook (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*");
            _pnlExtract.Controls.Add(_txtExcelPath);
            _pnlExtract.Controls.Add(btnBrowseExcel);
            ey += 30;

            _pnlExtract.Controls.Add(new Label { Text = "Sheet name:", Left = 0, Top = ey + 3, Width = 140, ForeColor = TkeTheme.Gray1 });
            _txtSheetName = new TextBox { Left = 145, Top = ey, Width = 150, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            _pnlExtract.Controls.Add(_txtSheetName);

            grpAction.Controls.Add(_pnlExtract);

            _pnlCommand = new Panel { Left = 10, Top = ay, Width = 555, Height = 150, BackColor = TkeTheme.Gray5 };
            int cy2 = 0;
            _pnlCommand.Controls.Add(new Label { Text = "Command / exe:", Left = 0, Top = cy2 + 3, Width = 140, ForeColor = TkeTheme.Gray1 });
            _txtCommand = new TextBox { Left = 145, Top = cy2, Width = 300, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            var btnBrowseCommand = TkeTheme.CreateSecondaryButton("...", 30, 24);
            btnBrowseCommand.Left = 450;
            btnBrowseCommand.Top = cy2 - 1;
            btnBrowseCommand.Click += (s, e) => BrowseFile(_txtCommand, "Executable/script (*.exe;*.ps1;*.bat;*.cmd)|*.exe;*.ps1;*.bat;*.cmd|All files (*.*)|*.*");
            _pnlCommand.Controls.Add(_txtCommand);
            _pnlCommand.Controls.Add(btnBrowseCommand);
            cy2 += 30;

            _pnlCommand.Controls.Add(new Label { Text = "Arguments:", Left = 0, Top = cy2 + 3, Width = 140, ForeColor = TkeTheme.Gray1 });
            _txtArguments = new TextBox { Left = 145, Top = cy2, Width = 340, Height = 50, Multiline = true, BackColor = Color.White, ForeColor = TkeTheme.Gray1 };
            _pnlCommand.Controls.Add(_txtArguments);
            cy2 += 58;

            var lblPlaceholders = new Label
            {
                Text = "Placeholders: {subject} {sender} {senderEmail} {receivedTime} {msgPath}\r\n({msgPath} saves the mail to a temp .msg file only if referenced)",
                Left = 0,
                Top = cy2,
                Width = 490,
                Height = 40,
                ForeColor = TkeTheme.Gray1
            };
            _pnlCommand.Controls.Add(lblPlaceholders);

            grpAction.Controls.Add(_pnlCommand);
            Controls.Add(grpAction);
            y += grpAction.Height + 10;

            var btnOk = TkeTheme.CreatePrimaryButton("OK", 320, y, 100, 36);
            var btnCancel = TkeTheme.CreateSecondaryButton("Cancel", 430, y, 100, 36);
            // A Button only auto-closes a modal dialog on click if its own DialogResult is
            // set (AcceptButton/CancelButton below only wire up Enter/Escape, not the click
            // itself) - without this, clicking OK/Cancel did nothing but raise Click.
            btnOk.DialogResult = DialogResult.OK;
            btnCancel.DialogResult = DialogResult.Cancel;
            btnOk.Click += BtnOk_Click;
            Controls.Add(btnOk);
            Controls.Add(btnCancel);

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            // Use ClientSize, not Height: Height includes the OS title bar/borders (which
            // vary by Windows version/DPI), so a fixed Height can silently clip content that
            // was laid out against the client area - this is what left OK/Cancel partly or
            // fully below the visible window before.
            ClientSize = new Size(ClientSize.Width, y + 36 + 20);
        }

        private void BrowseFile(TextBox target, string filter)
        {
            using (var dlg = new OpenFileDialog { Filter = filter })
            {
                if (!string.IsNullOrEmpty(target.Text) && System.IO.File.Exists(target.Text))
                {
                    dlg.InitialDirectory = System.IO.Path.GetDirectoryName(target.Text);
                }
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    target.Text = dlg.FileName;
                }
            }
        }

        private void UpdateActionPanels()
        {
            _pnlExtract.Visible = _rbExtract.Checked;
            _pnlCommand.Visible = _rbCommand.Checked;
        }

        private void LoadFromRule()
        {
            _txtName.Text = _rule.Name;
            _chkEnabled.Checked = _rule.Enabled;

            _txtSender.Text = _rule.Conditions.SenderContains;
            _txtSubject.Text = _rule.Conditions.SubjectContains;
            _txtBody.Text = _rule.Conditions.BodyContains;
            _chkRequiresTable.Checked = _rule.Conditions.RequiresTable;

            if (!_rule.Conditions.HasAttachment.HasValue) _cmbAttachment.SelectedIndex = 0;
            else _cmbAttachment.SelectedIndex = _rule.Conditions.HasAttachment.Value ? 1 : 2;

            _txtScriptPath.Text = _rule.Action.ScriptPath;
            _txtExcelPath.Text = _rule.Action.ExcelPath;
            _txtSheetName.Text = _rule.Action.SheetName;
            _txtCommand.Text = _rule.Action.Command;
            _txtArguments.Text = _rule.Action.Arguments;

            if (_rule.Action.Type == RuleAction.ExtractToExcel) _rbExtract.Checked = true;
            else _rbCommand.Checked = true;

            UpdateActionPanels();
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_txtName.Text))
            {
                MessageBox.Show(this, "Give the rule a name.", "Rule", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (_rbExtract.Checked)
            {
                if (string.IsNullOrWhiteSpace(_txtScriptPath.Text) || string.IsNullOrWhiteSpace(_txtExcelPath.Text))
                {
                    MessageBox.Show(this, "Fill in both the script path and the target Excel file.", "Rule", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }
            }
            else
            {
                if (string.IsNullOrWhiteSpace(_txtCommand.Text))
                {
                    MessageBox.Show(this, "Fill in a command/exe to run.", "Rule", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }
            }

            _rule.Name = _txtName.Text.Trim();
            _rule.Enabled = _chkEnabled.Checked;

            _rule.Conditions.SenderContains = _txtSender.Text.Trim();
            _rule.Conditions.SubjectContains = _txtSubject.Text.Trim();
            _rule.Conditions.BodyContains = _txtBody.Text.Trim();
            _rule.Conditions.RequiresTable = _chkRequiresTable.Checked;
            _rule.Conditions.HasAttachment = _cmbAttachment.SelectedIndex == 0 ? (bool?)null : _cmbAttachment.SelectedIndex == 1;

            if (_rbExtract.Checked)
            {
                _rule.Action.Type = RuleAction.ExtractToExcel;
                _rule.Action.ScriptPath = _txtScriptPath.Text.Trim();
                _rule.Action.ExcelPath = _txtExcelPath.Text.Trim();
                _rule.Action.SheetName = string.IsNullOrWhiteSpace(_txtSheetName.Text) ? "Sheet1" : _txtSheetName.Text.Trim();
            }
            else
            {
                _rule.Action.Type = RuleAction.RunCommand;
                _rule.Action.Command = _txtCommand.Text.Trim();
                _rule.Action.Arguments = _txtArguments.Text.Trim();
            }
        }
    }
}
