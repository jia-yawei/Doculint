using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class TableFormattingOptions
    {
        public string HeaderFontName { get; set; }
        public float HeaderFontSizePoints { get; set; }
        public string BodyFontName { get; set; }
        public float BodyFontSizePoints { get; set; }
        public float TableWidthCentimeters { get; set; }
        public float OuterBorderWidthPoints { get; set; }
        public float InnerBorderWidthPoints { get; set; }

        public static TableFormattingOptions CreateDefault()
        {
            return new TableFormattingOptions
            {
                HeaderFontName = "黑体",
                HeaderFontSizePoints = 12f,
                BodyFontName = "宋体",
                BodyFontSizePoints = 10.5f,
                TableWidthCentimeters = 17.4f,
                OuterBorderWidthPoints = 1.5f,
                InnerBorderWidthPoints = 0.5f
            };
        }

        public TableFormattingOptions Clone()
        {
            return new TableFormattingOptions
            {
                HeaderFontName = HeaderFontName,
                HeaderFontSizePoints = HeaderFontSizePoints,
                BodyFontName = BodyFontName,
                BodyFontSizePoints = BodyFontSizePoints,
                TableWidthCentimeters = TableWidthCentimeters,
                OuterBorderWidthPoints = OuterBorderWidthPoints,
                InnerBorderWidthPoints = InnerBorderWidthPoints
            };
        }
    }

    internal sealed class TableFormattingSettingsForm : Form
    {
        private readonly TextBox headerFontTextBox;
        private readonly NumericUpDown headerFontSizeNumeric;
        private readonly TextBox bodyFontTextBox;
        private readonly NumericUpDown bodyFontSizeNumeric;
        private readonly NumericUpDown tableWidthNumeric;
        private readonly NumericUpDown outerBorderNumeric;
        private readonly NumericUpDown innerBorderNumeric;
        private readonly Button cancelButton;
        private readonly Button editButton;
        private readonly Button confirmButton;

        public TableFormattingOptions Options { get; private set; }

        public TableFormattingSettingsForm(TableFormattingOptions options)
        {
            Options = (options ?? TableFormattingOptions.CreateDefault()).Clone();

            Font = SystemFonts.MessageBoxFont;
            Text = "一键规范表格";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;
            ClientSize = new Size(620, 410);

            Label titleLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 50,
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = Color.Black,
                Text = "一键规范表格参数",
                Padding = new Padding(22, 14, 0, 0)
            };

            Label scopeLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 70,
                ForeColor = Color.Black,
                Text = "处理范围：默认处理目录之后的所有表格；如果文档没有目录，则处理全部表格。\r\n表头识别：手动设置的表头优先，否则按自动识别结果处理。",
                Padding = new Padding(22, 2, 22, 0)
            };

            TableLayoutPanel grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 238,
                ColumnCount = 2,
                RowCount = 7,
                Padding = new Padding(22, 8, 22, 0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 180));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 7; i++)
            {
                grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));
            }

            headerFontTextBox = CreateTextBox(Options.HeaderFontName);
            headerFontSizeNumeric = CreateDecimalNumeric(8, 30, 0.5m, (decimal)Options.HeaderFontSizePoints);
            bodyFontTextBox = CreateTextBox(Options.BodyFontName);
            bodyFontSizeNumeric = CreateDecimalNumeric(8, 30, 0.5m, (decimal)Options.BodyFontSizePoints);
            tableWidthNumeric = CreateDecimalNumeric(1, 40, 0.1m, (decimal)Options.TableWidthCentimeters);
            outerBorderNumeric = CreateDecimalNumeric(0.25m, 6m, 0.25m, (decimal)Options.OuterBorderWidthPoints);
            innerBorderNumeric = CreateDecimalNumeric(0.25m, 6m, 0.25m, (decimal)Options.InnerBorderWidthPoints);

            AddField(grid, 0, "表头字体", headerFontTextBox);
            AddField(grid, 1, "表头字号（磅）", headerFontSizeNumeric);
            AddField(grid, 2, "正文字体", bodyFontTextBox);
            AddField(grid, 3, "正文字号（磅）", bodyFontSizeNumeric);
            AddField(grid, 4, "表格宽度（cm）", tableWidthNumeric);
            AddField(grid, 5, "外框线（磅）", outerBorderNumeric);
            AddField(grid, 6, "其他线宽（磅）", innerBorderNumeric);

            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 72,
                Padding = new Padding(0, 12, 22, 16),
                BackColor = Color.White
            };

            confirmButton = new Button
            {
                Text = "确认",
                Size = new Size(84, 30),
                Dock = DockStyle.Right
            };
            confirmButton.Click += ConfirmButton_Click;

            editButton = new Button
            {
                Text = "修改",
                Size = new Size(84, 30),
                Dock = DockStyle.Right
            };
            editButton.Click += EditButton_Click;

            cancelButton = new Button
            {
                Text = "取消",
                Size = new Size(84, 30),
                Dock = DockStyle.Right
            };
            cancelButton.Click += (sender, args) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };

            buttonPanel.Controls.Add(confirmButton);
            buttonPanel.Controls.Add(editButton);
            buttonPanel.Controls.Add(cancelButton);

            Controls.Add(buttonPanel);
            Controls.Add(grid);
            Controls.Add(scopeLabel);
            Controls.Add(titleLabel);

            SetEditingEnabled(false);
        }

        private void EditButton_Click(object sender, EventArgs e)
        {
            SetEditingEnabled(true);
            headerFontTextBox.Focus();
            editButton.Enabled = false;
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            Options = new TableFormattingOptions
            {
                HeaderFontName = string.IsNullOrWhiteSpace(headerFontTextBox.Text) ? "黑体" : headerFontTextBox.Text.Trim(),
                HeaderFontSizePoints = (float)headerFontSizeNumeric.Value,
                BodyFontName = string.IsNullOrWhiteSpace(bodyFontTextBox.Text) ? "宋体" : bodyFontTextBox.Text.Trim(),
                BodyFontSizePoints = (float)bodyFontSizeNumeric.Value,
                TableWidthCentimeters = (float)tableWidthNumeric.Value,
                OuterBorderWidthPoints = (float)outerBorderNumeric.Value,
                InnerBorderWidthPoints = (float)innerBorderNumeric.Value
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        private void SetEditingEnabled(bool enabled)
        {
            headerFontTextBox.ReadOnly = !enabled;
            bodyFontTextBox.ReadOnly = !enabled;
            headerFontSizeNumeric.Enabled = enabled;
            bodyFontSizeNumeric.Enabled = enabled;
            tableWidthNumeric.Enabled = enabled;
            outerBorderNumeric.Enabled = enabled;
            innerBorderNumeric.Enabled = enabled;
        }

        private static TextBox CreateTextBox(string text)
        {
            return new TextBox
            {
                Text = text ?? string.Empty,
                Dock = DockStyle.Fill,
                Height = 28
            };
        }

        private static NumericUpDown CreateDecimalNumeric(decimal min, decimal max, decimal increment, decimal value)
        {
            return new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                Increment = increment,
                DecimalPlaces = increment < 1m ? 1 : 0,
                Value = Math.Max(min, Math.Min(max, value)),
                Dock = DockStyle.Left,
                Width = 140,
                Height = 28
            };
        }

        private static void AddField(TableLayoutPanel grid, int rowIndex, string labelText, Control editor)
        {
            Label label = new Label
            {
                Text = labelText,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 6, 12, 6)
            };

            editor.Margin = new Padding(0, 4, 0, 4);
            grid.Controls.Add(label, 0, rowIndex);
            grid.Controls.Add(editor, 1, rowIndex);
        }
    }
}
