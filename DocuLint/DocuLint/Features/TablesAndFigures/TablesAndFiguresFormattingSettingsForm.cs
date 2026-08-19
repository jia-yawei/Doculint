using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class TablesAndFiguresFormattingSettings
    {
        public TableFormattingOptions TableOptions { get; set; }

        public static TablesAndFiguresFormattingSettings CreateDefault()
        {
            return new TablesAndFiguresFormattingSettings
            {
                TableOptions = TableFormattingOptions.CreateDefault()
            };
        }

        public TablesAndFiguresFormattingSettings Clone()
        {
            return new TablesAndFiguresFormattingSettings
            {
                TableOptions = (TableOptions ?? TableFormattingOptions.CreateDefault()).Clone()
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
        private readonly TextBox phraseLibraryPathTextBox;
        private readonly ListBox phrasePreviewListBox;
        private string commonPhraseLibraryPath;

        public TablesAndFiguresFormattingSettings Settings { get; private set; }

        internal string CommonPhraseLibraryPath => commonPhraseLibraryPath;

        public TablesAndFiguresFormattingSettingsForm(
            TablesAndFiguresFormattingSettings settings,
            string configuredCommonPhraseLibraryPath)
        {
            Settings = (settings ?? TablesAndFiguresFormattingSettings.CreateDefault()).Clone();
            commonPhraseLibraryPath = configuredCommonPhraseLibraryPath ?? string.Empty;

            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            Text = "快速工具设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;
            ClientSize = new Size(680, 520);
            MinimumSize = new Size(680, 520);

            TableFormattingOptions tableOptions = Settings.TableOptions ?? TableFormattingOptions.CreateDefault();
            headerFontTextBox = CreateTextBox(tableOptions.HeaderFontName);
            headerFontSizeNumeric = CreateDecimalNumeric(8, 30, 0.5m, (decimal)tableOptions.HeaderFontSizePoints);
            bodyFontTextBox = CreateTextBox(tableOptions.BodyFontName);
            bodyFontSizeNumeric = CreateDecimalNumeric(8, 30, 0.5m, (decimal)tableOptions.BodyFontSizePoints);
            tableWidthNumeric = CreateDecimalNumeric(1, 40, 0.1m, (decimal)tableOptions.TableWidthCentimeters);
            outerBorderNumeric = CreateDecimalNumeric(0.25m, 6m, 0.25m, (decimal)tableOptions.OuterBorderWidthPoints);
            innerBorderNumeric = CreateDecimalNumeric(0.25m, 6m, 0.25m, (decimal)tableOptions.InnerBorderWidthPoints);

            phraseLibraryPathTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White
            };
            phrasePreviewListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                HorizontalScrollbar = true
            };

            TableLayoutPanel rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = Color.White
            };
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 66f));

            TabControl tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Point(12, 5),
                Margin = new Padding(14, 14, 14, 0)
            };
            tabs.TabPages.Add(CreateTableSettingsTab());
            tabs.TabPages.Add(CreateCommonPhrasesTab());

            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 12, 18, 16),
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
            cancelButton.Click += (_, __) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };
            buttonPanel.Controls.Add(confirmButton);
            buttonPanel.Controls.Add(cancelButton);
            rootLayout.Controls.Add(tabs, 0, 0);
            rootLayout.Controls.Add(buttonPanel, 0, 1);
            Controls.Add(rootLayout);

            RefreshPhraseLibraryPreview();
        }

        private TabPage CreateTableSettingsTab()
        {
            TabPage page = new TabPage("快速表格样式") { BackColor = Color.White };
            TableLayoutPanel tableGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 7,
                Padding = new Padding(20)
            };
            tableGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200));
            tableGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            for (int i = 0; i < 7; i++)
            {
                tableGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            AddField(tableGrid, 0, "表头字体", headerFontTextBox);
            AddField(tableGrid, 1, "表头字号（磅）", headerFontSizeNumeric);
            AddField(tableGrid, 2, "正文字体", bodyFontTextBox);
            AddField(tableGrid, 3, "正文字号（磅）", bodyFontSizeNumeric);
            AddField(tableGrid, 4, "表格宽度（cm）", tableWidthNumeric);
            AddField(tableGrid, 5, "外框线（磅）", outerBorderNumeric);
            AddField(tableGrid, 6, "其他线宽（磅）", innerBorderNumeric);
            page.Controls.Add(tableGrid);
            return page;
        }

        private TabPage CreateCommonPhrasesTab()
        {
            TabPage page = new TabPage("常用语设置") { BackColor = Color.White };
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(20)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Label hint = new Label
            {
                AutoSize = true,
                Text = "加载独立 JSON 常用语库。当前仅支持纯文本字符串数组。",
                Margin = new Padding(0, 0, 0, 12)
            };
            TableLayoutPanel pathRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                Height = 32,
                Margin = new Padding(0, 0, 0, 12)
            };
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 124F));
            Button loadButton = new Button
            {
                Text = "加载常用语库",
                Dock = DockStyle.Fill,
                Margin = new Padding(8, 0, 0, 0)
            };
            loadButton.Click += (_, __) => SelectPhraseLibrary();
            pathRow.Controls.Add(phraseLibraryPathTextBox, 0, 0);
            pathRow.Controls.Add(loadButton, 1, 0);
            Label previewLabel = new Label
            {
                Text = "常用语预览",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 6)
            };

            layout.Controls.Add(hint, 0, 0);
            layout.Controls.Add(pathRow, 0, 1);
            layout.Controls.Add(previewLabel, 0, 2);
            layout.Controls.Add(phrasePreviewListBox, 0, 3);
            page.Controls.Add(layout);
            return page;
        }

        private void SelectPhraseLibrary()
        {
            string defaultLibraryPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "resource",
                "common-phrases.default.json");
            using (OpenFileDialog dialog = new OpenFileDialog
            {
                Title = "加载常用语库",
                Filter = "JSON 文件 (*.json)|*.json|所有文件 (*.*)|*.*",
                CheckFileExists = true,
                Multiselect = false,
                InitialDirectory = File.Exists(defaultLibraryPath)
                    ? Path.GetDirectoryName(defaultLibraryPath)
                    : string.Empty,
                FileName = File.Exists(defaultLibraryPath)
                    ? Path.GetFileName(defaultLibraryPath)
                    : string.Empty
            })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                if (!CommonPhraseLibrary.TryLoad(dialog.FileName, out List<string> phrases, out string error))
                {
                    MessageBox.Show(this, error, "常用语设置", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                commonPhraseLibraryPath = dialog.FileName;
                SetPhrasePreview(phrases);
            }
        }

        private void RefreshPhraseLibraryPreview()
        {
            phraseLibraryPathTextBox.Text = commonPhraseLibraryPath;
            if (CommonPhraseLibrary.TryLoad(commonPhraseLibraryPath, out List<string> phrases, out string _))
            {
                SetPhrasePreview(phrases);
            }
            else
            {
                SetPhrasePreview(Enumerable.Empty<string>());
            }
        }

        private void SetPhrasePreview(IEnumerable<string> phrases)
        {
            List<string> values = (phrases ?? Enumerable.Empty<string>()).ToList();
            phraseLibraryPathTextBox.Text = commonPhraseLibraryPath;
            phrasePreviewListBox.BeginUpdate();
            try
            {
                phrasePreviewListBox.Items.Clear();
                foreach (string phrase in values)
                {
                    phrasePreviewListBox.Items.Add(phrase);
                }
            }
            finally
            {
                phrasePreviewListBox.EndUpdate();
            }
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
