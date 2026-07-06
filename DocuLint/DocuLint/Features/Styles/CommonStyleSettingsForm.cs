using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class CommonTextStyleOption
    {
        public string Label { get; set; }
        public string FontName { get; set; }
        public float FontSizePoints { get; set; }
        public int OutlineLevel { get; set; }

        public CommonTextStyleOption Clone()
        {
            return new CommonTextStyleOption
            {
                Label = Label,
                FontName = FontName,
                FontSizePoints = FontSizePoints,
                OutlineLevel = OutlineLevel
            };
        }
    }

    internal sealed class CommonStyleSettings
    {
        public List<CommonTextStyleOption> Styles { get; set; }
        public OutlineNumberPattern NumberPattern { get; set; }

        public static CommonStyleSettings CreateDefault()
        {
            List<CommonTextStyleOption> styles = new List<CommonTextStyleOption>();
            for (int i = 1; i <= 6; i++)
            {
                styles.Add(new CommonTextStyleOption
                {
                    Label = $"{i}级标题",
                    FontName = i == 1 ? "黑体" : "宋体",
                    FontSizePoints = 12f,
                    OutlineLevel = i
                });
            }

            styles.Add(new CommonTextStyleOption
            {
                Label = "正文",
                FontName = "宋体",
                FontSizePoints = 12f,
                OutlineLevel = 10
            });

            return new CommonStyleSettings
            {
                Styles = styles,
                NumberPattern = OutlineNumberPattern.Decimal
            };
        }

        public CommonStyleSettings Clone()
        {
            return new CommonStyleSettings
            {
                Styles = (Styles ?? new List<CommonTextStyleOption>()).Select(item => item.Clone()).ToList(),
                NumberPattern = NumberPattern
            };
        }
    }

    internal sealed class CommonStyleSettingsForm : Form
    {
        private readonly List<TextBox> fontTextBoxes = new List<TextBox>();
        private readonly List<NumericUpDown> sizeInputs = new List<NumericUpDown>();
        private readonly ComboBox numberPatternComboBox;
        private readonly CommonStyleSettings workingSettings;

        public CommonStyleSettings Settings { get; private set; }

        public CommonStyleSettingsForm(CommonStyleSettings settings)
        {
            workingSettings = (settings ?? CommonStyleSettings.CreateDefault()).Clone();

            Font = SystemFonts.MessageBoxFont;
            Text = "常用样式库设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.White;
            ClientSize = new Size(560, 390);

            Label title = new Label
            {
                Dock = DockStyle.Top,
                Height = 42,
                Padding = new Padding(18, 12, 0, 0),
                Font = new Font(Font, FontStyle.Bold),
                Text = "设置标题/正文格式"
            };

            TableLayoutPanel grid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 252,
                ColumnCount = 3,
                RowCount = 8,
                Padding = new Padding(18, 6, 18, 0)
            };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            for (int i = 0; i < 8; i++)
            {
                grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            }

            AddHeader(grid);
            for (int i = 0; i < workingSettings.Styles.Count; i++)
            {
                AddStyleRow(grid, i + 1, workingSettings.Styles[i]);
            }

            Label numberLabel = new Label
            {
                Text = "自动章节号类型",
                Location = new Point(18, 306),
                Size = new Size(120, 26),
                TextAlign = ContentAlignment.MiddleLeft
            };

            numberPatternComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Location = new Point(145, 306),
                Size = new Size(170, 26)
            };
            numberPatternComboBox.Items.Add("1, 1.1, 1.1.1");
            numberPatternComboBox.Items.Add("(1), (1.1), (1.1.1)");
            numberPatternComboBox.SelectedIndex = workingSettings.NumberPattern == OutlineNumberPattern.Parenthesized ? 1 : 0;

            Button okButton = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Location = new Point(372, 344),
                Size = new Size(78, 28)
            };
            okButton.Click += OkButton_Click;

            Button cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Location = new Point(462, 344),
                Size = new Size(78, 28)
            };

            Controls.Add(cancelButton);
            Controls.Add(okButton);
            Controls.Add(numberPatternComboBox);
            Controls.Add(numberLabel);
            Controls.Add(grid);
            Controls.Add(title);

            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        private static void AddHeader(TableLayoutPanel grid)
        {
            grid.Controls.Add(CreateHeaderLabel("项目"), 0, 0);
            grid.Controls.Add(CreateHeaderLabel("字体"), 1, 0);
            grid.Controls.Add(CreateHeaderLabel("字号"), 2, 0);
        }

        private void AddStyleRow(TableLayoutPanel grid, int row, CommonTextStyleOption option)
        {
            grid.Controls.Add(new Label
            {
                Text = option.Label,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            }, 0, row);

            TextBox fontBox = new TextBox
            {
                Text = option.FontName,
                Dock = DockStyle.Fill
            };
            fontTextBoxes.Add(fontBox);
            grid.Controls.Add(fontBox, 1, row);

            NumericUpDown sizeInput = new NumericUpDown
            {
                Minimum = 8,
                Maximum = 30,
                Increment = 0.5m,
                DecimalPlaces = 1,
                Value = (decimal)option.FontSizePoints,
                Dock = DockStyle.Left,
                Width = 82
            };
            sizeInputs.Add(sizeInput);
            grid.Controls.Add(sizeInput, 2, row);
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < workingSettings.Styles.Count; i++)
            {
                string fontName = fontTextBoxes[i].Text.Trim();
                workingSettings.Styles[i].FontName = string.IsNullOrWhiteSpace(fontName)
                    ? (i == 0 ? "黑体" : "宋体")
                    : fontName;
                workingSettings.Styles[i].FontSizePoints = (float)sizeInputs[i].Value;
            }

            workingSettings.NumberPattern = numberPatternComboBox.SelectedIndex == 1
                ? OutlineNumberPattern.Parenthesized
                : OutlineNumberPattern.Decimal;

            Settings = workingSettings.Clone();
            DialogResult = DialogResult.OK;
            Close();
        }

        private static Label CreateHeaderLabel(string text)
        {
            return new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft
            };
        }
    }
}
