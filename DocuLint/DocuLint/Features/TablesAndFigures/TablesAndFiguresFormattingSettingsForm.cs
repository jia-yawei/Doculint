using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class TablesAndFiguresFormattingSettings
    {
        public TableFormattingOptions TableOptions { get; set; }
        public ImageFormattingOptions ImageOptions { get; set; }

        public static TablesAndFiguresFormattingSettings CreateDefault()
        {
            return new TablesAndFiguresFormattingSettings
            {
                TableOptions = TableFormattingOptions.CreateDefault(),
                ImageOptions = ImageFormattingOptions.CreateDefault()
            };
        }

        public TablesAndFiguresFormattingSettings Clone()
        {
            return new TablesAndFiguresFormattingSettings
            {
                TableOptions = (TableOptions ?? TableFormattingOptions.CreateDefault()).Clone(),
                ImageOptions = (ImageOptions ?? ImageFormattingOptions.CreateDefault()).Clone()
            };
        }
    }

    internal sealed class TablesAndFiguresFormattingSettingsForm : Form
    {
        private readonly TextBox headerFontTextBox;
        private readonly NumericUpDown headerFontSizeNumeric;
        private readonly TextBox bodyFontTextBox;
        private readonly NumericUpDown bodyFontSizeNumeric;
        private readonly NumericUpDown tableWidthNumeric;
        private readonly NumericUpDown outerBorderNumeric;
        private readonly NumericUpDown innerBorderNumeric;
        private readonly RadioButton scale50Radio;
        private readonly RadioButton scale75Radio;
        private readonly RadioButton scale100Radio;

        public TablesAndFiguresFormattingSettings Settings { get; private set; }

        public TablesAndFiguresFormattingSettingsForm(TablesAndFiguresFormattingSettings settings)
        {
            Settings = (settings ?? TablesAndFiguresFormattingSettings.CreateDefault()).Clone();

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            Text = "图表规范参数";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;
            ClientSize = new Size(600, 700);
            MinimumSize = new Size(600, 700);

            TableFormattingOptions tableOptions = Settings.TableOptions ?? TableFormattingOptions.CreateDefault();
            ImageFormattingOptions imageOptions = Settings.ImageOptions ?? ImageFormattingOptions.CreateDefault();

            TableLayoutPanel rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = Color.White
            };
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 72f));

            Panel contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            TableLayoutPanel contentLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(22, 18, 22, 18),
                BackColor = Color.White
            };
            contentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            contentLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            contentLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            GroupBox tableGroup = new GroupBox
            {
                Text = "表格规范",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(16, 18, 16, 14),
                Margin = new Padding(0, 0, 0, 16)
            };

            TableLayoutPanel tableGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 7
            };
            tableGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            tableGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            for (int i = 0; i < 7; i++)
            {
                tableGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            headerFontTextBox = CreateTextBox(tableOptions.HeaderFontName);
            headerFontSizeNumeric = CreateDecimalNumeric(8, 30, 0.5m, (decimal)tableOptions.HeaderFontSizePoints);
            bodyFontTextBox = CreateTextBox(tableOptions.BodyFontName);
            bodyFontSizeNumeric = CreateDecimalNumeric(8, 30, 0.5m, (decimal)tableOptions.BodyFontSizePoints);
            tableWidthNumeric = CreateDecimalNumeric(1, 40, 0.1m, (decimal)tableOptions.TableWidthCentimeters);
            outerBorderNumeric = CreateDecimalNumeric(0.25m, 6m, 0.25m, (decimal)tableOptions.OuterBorderWidthPoints);
            innerBorderNumeric = CreateDecimalNumeric(0.25m, 6m, 0.25m, (decimal)tableOptions.InnerBorderWidthPoints);

            AddField(tableGrid, 0, "表头字体", headerFontTextBox);
            AddField(tableGrid, 1, "表头字号（磅）", headerFontSizeNumeric);
            AddField(tableGrid, 2, "正文字体", bodyFontTextBox);
            AddField(tableGrid, 3, "正文字号（磅）", bodyFontSizeNumeric);
            AddField(tableGrid, 4, "表格宽度（cm）", tableWidthNumeric);
            AddField(tableGrid, 5, "外框线（磅）", outerBorderNumeric);
            AddField(tableGrid, 6, "其他线宽（磅）", innerBorderNumeric);
            tableGroup.Controls.Add(tableGrid);

            GroupBox imageGroup = new GroupBox
            {
                Text = "图片规范",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(16, 18, 16, 14),
                Margin = new Padding(0)
            };

            TableLayoutPanel imagePanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 2
            };
            imagePanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            imagePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            imagePanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label imageHintLabel = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                MaximumSize = new Size(620, 0),
                Text = "图片大小：请选择执行“规范全部图片”时使用的缩放比例。",
                Margin = new Padding(0, 0, 0, 10)
            };

            FlowLayoutPanel radioPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Margin = new Padding(0)
            };

            scale50Radio = new RadioButton
            {
                Text = "50%",
                AutoSize = true,
                Margin = new Padding(0, 2, 24, 0)
            };
            scale75Radio = new RadioButton
            {
                Text = "75%",
                AutoSize = true,
                Margin = new Padding(0, 2, 24, 0)
            };
            scale100Radio = new RadioButton
            {
                Text = "100%",
                AutoSize = true,
                Margin = new Padding(0, 2, 24, 0)
            };

            scale50Radio.Checked = imageOptions.ScalePercent == 50;
            scale75Radio.Checked = imageOptions.ScalePercent == 75;
            scale100Radio.Checked = !scale50Radio.Checked && !scale75Radio.Checked;

            radioPanel.Controls.Add(scale50Radio);
            radioPanel.Controls.Add(scale75Radio);
            radioPanel.Controls.Add(scale100Radio);

            imagePanel.Controls.Add(imageHintLabel, 0, 0);
            imagePanel.Controls.Add(radioPanel, 0, 1);
            imageGroup.Controls.Add(imagePanel);

            contentLayout.Controls.Add(tableGroup, 0, 0);
            contentLayout.Controls.Add(imageGroup, 0, 1);
            contentPanel.Controls.Add(contentLayout);

            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 12, 22, 16),
                BackColor = Color.White
            };

            Button confirmButton = new Button
            {
                Text = "确认",
                Size = new Size(84, 30),
                Dock = DockStyle.Right
            };
            confirmButton.Click += ConfirmButton_Click;

            Button cancelButton = new Button
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
            buttonPanel.Controls.Add(cancelButton);

            rootLayout.Controls.Add(contentPanel, 0, 0);
            rootLayout.Controls.Add(buttonPanel, 0, 1);
            Controls.Add(rootLayout);
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            Settings = new TablesAndFiguresFormattingSettings
            {
                TableOptions = new TableFormattingOptions
                {
                    HeaderFontName = string.IsNullOrWhiteSpace(headerFontTextBox.Text) ? "黑体" : headerFontTextBox.Text.Trim(),
                    HeaderFontSizePoints = (float)headerFontSizeNumeric.Value,
                    BodyFontName = string.IsNullOrWhiteSpace(bodyFontTextBox.Text) ? "宋体" : bodyFontTextBox.Text.Trim(),
                    BodyFontSizePoints = (float)bodyFontSizeNumeric.Value,
                    TableWidthCentimeters = (float)tableWidthNumeric.Value,
                    OuterBorderWidthPoints = (float)outerBorderNumeric.Value,
                    InnerBorderWidthPoints = (float)innerBorderNumeric.Value
                },
                ImageOptions = new ImageFormattingOptions
                {
                    ScalePercent = scale50Radio.Checked ? 50 : (scale75Radio.Checked ? 75 : 100)
                }
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        private static TextBox CreateTextBox(string text)
        {
            return new TextBox
            {
                Text = text ?? string.Empty,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Width = 360,
                Margin = new Padding(0, 2, 0, 2)
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
                Anchor = AnchorStyles.Left,
                Width = 140,
                Margin = new Padding(0, 2, 0, 2)
            };
        }

        private static void AddField(TableLayoutPanel grid, int rowIndex, string labelText, Control editor)
        {
            Label label = new Label
            {
                Text = labelText,
                TextAlign = ContentAlignment.MiddleLeft,
                Dock = DockStyle.Fill,
                AutoSize = true,
                Margin = new Padding(0, 8, 12, 8)
            };

            grid.Controls.Add(label, 0, rowIndex);
            grid.Controls.Add(editor, 1, rowIndex);
        }
    }
}
