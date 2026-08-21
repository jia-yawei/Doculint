using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class ChapterNumberRepairFormattingSettings
    {
        internal bool ApplyFormatting { get; set; } = true;
        internal string LevelOneFontName { get; set; } = "黑体";
        internal decimal LevelOneFontSize { get; set; } = 12M;
        internal string OtherLevelsFontName { get; set; } = "宋体";
        internal decimal OtherLevelsFontSize { get; set; } = 12M;
        internal bool Bold { get; set; }
        internal int Alignment { get; set; } = 0;
        internal int LineSpacingRule { get; set; } = 0;
        internal decimal SpaceBefore { get; set; }
        internal decimal SpaceAfter { get; set; }
    }

    internal sealed class CommonStyleSettingsForm : Form
    {
        private const int MaxCommonStyles = 9;
        private readonly Func<Action<int, int>, IEnumerable<string>> styleLibraryLoader;
        private readonly ComboBox availableStylesComboBox;
        private readonly Button addButton;
        private readonly Button loadStylesButton;
        private readonly ProgressBar loadProgressBar;
        private readonly Label loadStatusLabel;
        private readonly ListBox selectedStylesListBox;
        private readonly Label countLabel;
        private readonly List<string> selectedStyleNames;
        private readonly ChapterNumberRepairFormattingSettings formattingSettings;
        private bool loadingStyleLibrary;
        private bool styleLibraryLoaded;

        internal CommonStyleSettingsForm(
            Func<Action<int, int>, IEnumerable<string>> styleLibraryLoader,
            IEnumerable<string> configuredStyleNames,
            IEnumerable<string> loadedStyleNames = null,
            ChapterNumberRepairFormattingSettings configuredFormatting = null)
        {
            this.styleLibraryLoader = styleLibraryLoader;
            formattingSettings = configuredFormatting ?? new ChapterNumberRepairFormattingSettings();
            selectedStyleNames = (configuredStyleNames ?? Enumerable.Empty<string>())
                .Select(name => (name ?? string.Empty).Trim())
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(MaxCommonStyles)
                .ToList();

            Text = "样式管理设置";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(570, 450);
            Font = new Font("Microsoft YaHei UI", 9F);

            TabControl tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Point(12, 5)
            };
            TabPage commonStylesTab = new TabPage("常用样式设置")
            {
                Padding = new Padding(3)
            };
            TabPage formattingTab = new TabPage("格式设置")
            {
                Padding = new Padding(3)
            };
            tabs.TabPages.Add(commonStylesTab);
            tabs.TabPages.Add(formattingTab);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(16),
                AutoSize = false
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label titleLabel = new Label
            {
                AutoSize = true,
                Text = "添加常用样式",
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 8)
            };

            TableLayoutPanel loadRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 3,
                RowCount = 1,
                Height = 34,
                Margin = new Padding(0, 0, 0, 8)
            };
            loadRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108F));
            loadRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            loadRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120F));

            loadStylesButton = new Button
            {
                Text = "加载样式库",
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.System,
                Margin = new Padding(0, 0, 8, 0)
            };
            loadStylesButton.Click += (_, __) => LoadStyleLibrary();
            loadStatusLabel = new Label
            {
                Text = "未加载样式库",
                AutoEllipsis = true,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(82, 91, 104),
                Margin = new Padding(0)
            };
            loadProgressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Minimum = 0,
                Maximum = 1,
                Value = 0,
                Margin = new Padding(8, 5, 0, 5)
            };
            loadRow.Controls.Add(loadStylesButton, 0, 0);
            loadRow.Controls.Add(loadStatusLabel, 1, 0);
            loadRow.Controls.Add(loadProgressBar, 2, 0);

            TableLayoutPanel addRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                Height = 34,
                Margin = new Padding(0, 0, 0, 12)
            };
            addRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            addRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 78F));

            availableStylesComboBox = new ComboBox
            {
                Dock = DockStyle.Fill,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 0, 8, 0),
                Enabled = false
            };

            addButton = new Button
            {
                Text = "添加",
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.System,
                Enabled = false
            };
            addButton.Click += (_, __) => AddSelectedStyle();
            addRow.Controls.Add(availableStylesComboBox, 0, 0);
            addRow.Controls.Add(addButton, 1, 0);

            TableLayoutPanel listSection = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Margin = Padding.Empty
            };
            listSection.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            listSection.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 78F));
            listSection.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            listSection.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            countLabel = new Label
            {
                AutoSize = true,
                ForeColor = Color.FromArgb(82, 91, 104),
                Margin = new Padding(0, 0, 0, 6)
            };
            selectedStylesListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                SelectionMode = SelectionMode.One
            };
            Button removeButton = new Button
            {
                Text = "删除",
                Dock = DockStyle.Top,
                Height = 30,
                FlatStyle = FlatStyle.System,
                Margin = new Padding(8, 0, 0, 0)
            };
            removeButton.Click += (_, __) => RemoveSelectedStyle();
            listSection.Controls.Add(countLabel, 0, 0);
            listSection.SetColumnSpan(countLabel, 2);
            listSection.Controls.Add(selectedStylesListBox, 0, 1);
            listSection.Controls.Add(removeButton, 1, 1);

            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                Margin = new Padding(0, 12, 0, 0)
            };
            Button saveButton = new Button
            {
                Text = "保存",
                DialogResult = DialogResult.OK,
                AutoSize = true,
                MinimumSize = new Size(76, 30),
                FlatStyle = FlatStyle.System,
                Margin = new Padding(6, 0, 0, 0)
            };
            Button cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                AutoSize = true,
                MinimumSize = new Size(76, 30),
                FlatStyle = FlatStyle.System
            };
            buttons.Controls.Add(saveButton);
            buttons.Controls.Add(cancelButton);

            layout.Controls.Add(titleLabel, 0, 0);
            layout.Controls.Add(loadRow, 0, 1);
            layout.Controls.Add(addRow, 0, 2);
            layout.Controls.Add(listSection, 0, 3);
            layout.Controls.Add(buttons, 0, 4);
            commonStylesTab.Controls.Add(layout);
            formattingTab.Controls.Add(CreateFormattingPanel());
            Controls.Add(tabs);

            AcceptButton = saveButton;
            CancelButton = cancelButton;
            FormClosing += (_, e) => e.Cancel = loadingStyleLibrary;
            RefreshSelectedStyleList();

            List<string> cachedStyles = NormalizeStyleNames(loadedStyleNames);
            if (cachedStyles.Count > 0)
            {
                ApplyStyleLibrary(cachedStyles);
                loadStylesButton.Enabled = false;
                loadStatusLabel.Text = "已加载 " + cachedStyles.Count + " 个段落样式";
            }
        }

        internal IReadOnlyList<string> SelectedStyleNames => selectedStyleNames;

        internal ChapterNumberRepairFormattingSettings FormattingSettings => formattingSettings;

        private Control CreateFormattingPanel()
        {
            TableLayoutPanel panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 9,
                Padding = new Padding(14),
                AutoScroll = true
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 190F));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            for (int row = 0; row < 9; row++)
            {
                panel.RowStyles.Add(new RowStyle(SizeType.Absolute, row == 0 ? 38F : 36F));
            }

            CheckBox applyFormatting = new CheckBox
            {
                Text = "修复章节号后同步应用标题格式",
                Checked = formattingSettings.ApplyFormatting,
                Dock = DockStyle.Fill,
                AutoSize = true
            };
            applyFormatting.CheckedChanged += (_, __) => formattingSettings.ApplyFormatting = applyFormatting.Checked;
            panel.Controls.Add(applyFormatting, 0, 0);
            panel.SetColumnSpan(applyFormatting, 2);

            ComboBox firstFont = CreateFontSelector(formattingSettings.LevelOneFontName);
            firstFont.TextChanged += (_, __) => formattingSettings.LevelOneFontName = firstFont.Text.Trim();
            NumericUpDown firstSize = CreateFontSizeSelector(formattingSettings.LevelOneFontSize);
            firstSize.ValueChanged += (_, __) => formattingSettings.LevelOneFontSize = firstSize.Value;
            ComboBox otherFont = CreateFontSelector(formattingSettings.OtherLevelsFontName);
            otherFont.TextChanged += (_, __) => formattingSettings.OtherLevelsFontName = otherFont.Text.Trim();
            NumericUpDown otherSize = CreateFontSizeSelector(formattingSettings.OtherLevelsFontSize);
            otherSize.ValueChanged += (_, __) => formattingSettings.OtherLevelsFontSize = otherSize.Value;
            CheckBox bold = new CheckBox { Text = "加粗", Checked = formattingSettings.Bold, AutoSize = true, Anchor = AnchorStyles.Left };
            bold.CheckedChanged += (_, __) => formattingSettings.Bold = bold.Checked;
            ComboBox alignment = new ComboBox { Dock = DockStyle.Left, Width = 170, DropDownStyle = ComboBoxStyle.DropDownList };
            alignment.Items.AddRange(new object[] { "左对齐", "居中", "右对齐", "两端对齐" });
            alignment.SelectedIndex = Math.Max(0, Math.Min(3, formattingSettings.Alignment));
            alignment.SelectedIndexChanged += (_, __) => formattingSettings.Alignment = alignment.SelectedIndex;
            ComboBox lineSpacing = new ComboBox { Dock = DockStyle.Left, Width = 170, DropDownStyle = ComboBoxStyle.DropDownList };
            lineSpacing.Items.AddRange(new object[] { "单倍行距", "1.5 倍行距", "2 倍行距" });
            lineSpacing.SelectedIndex = Math.Max(0, Math.Min(2, formattingSettings.LineSpacingRule));
            lineSpacing.SelectedIndexChanged += (_, __) => formattingSettings.LineSpacingRule = lineSpacing.SelectedIndex;
            NumericUpDown before = CreateSpacingSelector(formattingSettings.SpaceBefore);
            before.ValueChanged += (_, __) => formattingSettings.SpaceBefore = before.Value;
            NumericUpDown after = CreateSpacingSelector(formattingSettings.SpaceAfter);
            after.ValueChanged += (_, __) => formattingSettings.SpaceAfter = after.Value;

            AddFormatRow(panel, 1, "一级标题字体", firstFont);
            AddFormatRow(panel, 2, "一级标题字号（磅）", firstSize);
            AddFormatRow(panel, 3, "其他标题字体", otherFont);
            AddFormatRow(panel, 4, "其他标题字号（磅）", otherSize);
            AddFormatRow(panel, 5, "标题字形", bold);
            AddFormatRow(panel, 6, "段落对齐", alignment);
            AddFormatRow(panel, 7, "行距", lineSpacing);
            TableLayoutPanel spacingRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3 };
            spacingRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            spacingRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105F));
            spacingRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 105F));
            spacingRow.Controls.Add(new Label { Text = "段前 / 段后（磅）", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            spacingRow.Controls.Add(before, 1, 0);
            spacingRow.Controls.Add(after, 2, 0);
            AddFormatRow(panel, 8, "段落间距", spacingRow);
            return panel;
        }

        private static void AddFormatRow(TableLayoutPanel panel, int row, string label, Control control)
        {
            panel.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 10, 0) }, 0, row);
            control.Dock = control.Dock == DockStyle.None ? DockStyle.Fill : control.Dock;
            control.Margin = new Padding(0, 4, 0, 4);
            panel.Controls.Add(control, 1, row);
        }

        private static ComboBox CreateFontSelector(string value)
        {
            ComboBox selector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Dock = DockStyle.Fill };
            selector.Items.AddRange(new object[] { "黑体", "宋体", "仿宋", "楷体", "微软雅黑" });
            selector.Text = string.IsNullOrWhiteSpace(value) ? "宋体" : value;
            return selector;
        }

        private static NumericUpDown CreateFontSizeSelector(decimal value)
        {
            return new NumericUpDown { Minimum = 5M, Maximum = 72M, DecimalPlaces = 1, Increment = 0.5M, Value = Math.Max(5M, Math.Min(72M, value)), Width = 170, Anchor = AnchorStyles.Left };
        }

        private static NumericUpDown CreateSpacingSelector(decimal value)
        {
            return new NumericUpDown { Minimum = 0M, Maximum = 120M, DecimalPlaces = 1, Increment = 0.5M, Value = Math.Max(0M, Math.Min(120M, value)), Width = 96, Anchor = AnchorStyles.Left };
        }

        private void LoadStyleLibrary()
        {
            if (styleLibraryLoader == null)
            {
                return;
            }

            loadStylesButton.Enabled = false;
            loadStatusLabel.Text = "正在读取样式库...";
            loadProgressBar.Maximum = 1;
            loadProgressBar.Value = 0;
            UseWaitCursor = true;
            loadingStyleLibrary = true;

            try
            {
                IEnumerable<string> loadedStyles = styleLibraryLoader(ReportStyleLibraryProgress);
                List<string> styleNames = (loadedStyles ?? Enumerable.Empty<string>())
                    .Select(name => (name ?? string.Empty).Trim())
                    .Where(name => name.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                ApplyStyleLibrary(styleNames);
                if (styleNames.Count > 0)
                {
                    loadStatusLabel.Text = "已加载 " + styleNames.Count + " 个段落样式";
                }
                else
                {
                    loadStatusLabel.Text = "未找到可用的段落样式";
                }
            }
            catch (Exception ex)
            {
                loadStatusLabel.Text = "样式库加载失败";
                MessageBox.Show(this, "加载样式库失败：\r\n" + ex.Message, "常用样式");
            }
            finally
            {
                loadStylesButton.Enabled = !styleLibraryLoaded;
                UseWaitCursor = false;
                loadingStyleLibrary = false;
            }
        }

        private static List<string> NormalizeStyleNames(IEnumerable<string> styleNames)
        {
            return (styleNames ?? Enumerable.Empty<string>())
                .Select(name => (name ?? string.Empty).Trim())
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private void ApplyStyleLibrary(IEnumerable<string> styleNames)
        {
            List<string> names = NormalizeStyleNames(styleNames);
            availableStylesComboBox.BeginUpdate();
            try
            {
                availableStylesComboBox.Items.Clear();
                foreach (string styleName in names)
                {
                    availableStylesComboBox.Items.Add(styleName);
                }
            }
            finally
            {
                availableStylesComboBox.EndUpdate();
            }

            availableStylesComboBox.Enabled = names.Count > 0;
            addButton.Enabled = names.Count > 0;
            styleLibraryLoaded = names.Count > 0;
            if (names.Count > 0)
            {
                availableStylesComboBox.SelectedIndex = 0;
            }
        }

        private void ReportStyleLibraryProgress(int current, int total)
        {
            int safeTotal = Math.Max(1, total);
            int safeCurrent = Math.Max(0, Math.Min(current, safeTotal));
            loadProgressBar.Maximum = safeTotal;
            loadProgressBar.Value = safeCurrent;
            loadStatusLabel.Text = "正在加载样式库：" + safeCurrent + " / " + total;
            loadProgressBar.Refresh();
            loadStatusLabel.Refresh();
            Application.DoEvents();
        }

        private void AddSelectedStyle()
        {
            string styleName = availableStylesComboBox.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(styleName))
            {
                MessageBox.Show(this, "请选择一种样式。", "常用样式");
                return;
            }

            if (selectedStyleNames.Any(name => string.Equals(name, styleName, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            if (selectedStyleNames.Count >= MaxCommonStyles)
            {
                MessageBox.Show(this, "最多只能添加 9 个常用样式。", "常用样式");
                return;
            }

            selectedStyleNames.Add(styleName);
            RefreshSelectedStyleList();
        }

        private void RemoveSelectedStyle()
        {
            int index = selectedStylesListBox.SelectedIndex;
            if (index < 0 || index >= selectedStyleNames.Count)
            {
                return;
            }

            selectedStyleNames.RemoveAt(index);
            RefreshSelectedStyleList();
        }

        private void RefreshSelectedStyleList()
        {
            selectedStylesListBox.BeginUpdate();
            try
            {
                selectedStylesListBox.Items.Clear();
                foreach (string styleName in selectedStyleNames)
                {
                    selectedStylesListBox.Items.Add(styleName);
                }
            }
            finally
            {
                selectedStylesListBox.EndUpdate();
            }

            countLabel.Text = $"已添加 {selectedStyleNames.Count} / {MaxCommonStyles} 个常用样式";
        }
    }
}
