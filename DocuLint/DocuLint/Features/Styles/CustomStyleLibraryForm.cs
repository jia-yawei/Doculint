using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class CustomStyleLibraryForm : Form
    {
        private static readonly string[] FontNames = { "宋体", "黑体", "仿宋", "楷体", "微软雅黑", "Times New Roman", "Arial" };
        private static readonly string[] FontSizes = { "小四", "五号", "四号", "小三", "三号", "12", "10.5", "14", "16" };
        private readonly Dictionary<int, StyleDefinitionRequest> definitionsByLevel;
        private readonly ListBox templateListBox;
        private readonly ComboBox outlineLevelComboBox;
        private readonly ComboBox fontComboBox;
        private readonly ComboBox fontSizeComboBox;
        private readonly ComboBox alignmentComboBox;
        private readonly CheckBox boldCheckBox;
        private readonly ComboBox lineSpacingComboBox;
        private readonly Label previewLabel;
        private readonly CheckedListBox createStyleCheckedListBox;
        private bool loading;
        private int currentLevel = 1;

        public List<StyleDefinitionRequest> StyleDefinitions { get; } = new List<StyleDefinitionRequest>();
        public CustomStyleLibraryForm(IEnumerable<StyleDefinitionRequest> currentDefinitions)
        {
            definitionsByLevel = (currentDefinitions ?? StyleDefinitionRequest.CreateDefaultSet())
                .Where(item => item != null)
                .Select(CloneDefinition)
                .GroupBy(item => item.Level)
                .ToDictionary(group => group.Key, group => group.First());
            foreach (StyleDefinitionRequest definition in StyleDefinitionRequest.CreateDefaultSet())
            {
                if (!definitionsByLevel.ContainsKey(definition.Level))
                    definitionsByLevel[definition.Level] = CloneDefinition(definition);
            }

            Text = "创建自定义样式";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(860, 650);
            MinimumSize = new Size(800, 610);
            Font = new Font("Microsoft YaHei UI", 9F);
            Padding = new Padding(12);

            TableLayoutPanel root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52f));
            Controls.Add(root);

            TabControl tabs = new TabControl { Dock = DockStyle.Fill };
            TabPage createPage = new TabPage("样式创建") { Padding = new Padding(10), BackColor = Color.White };
            TabPage editPage = new TabPage("样式编辑") { Padding = new Padding(10), BackColor = Color.White };
            tabs.TabPages.Add(createPage);
            tabs.TabPages.Add(editPage);
            root.Controls.Add(tabs, 0, 0);

            templateListBox = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, BorderStyle = BorderStyle.FixedSingle };
            templateListBox.SelectedIndexChanged += templateListBox_SelectedIndexChanged;
            outlineLevelComboBox = CreateComboBox(160, false);
            for (int i = 1; i <= 9; i++) outlineLevelComboBox.Items.Add(i + "级");
            outlineLevelComboBox.Items.Add("正文");
            fontComboBox = CreateComboBox(200, true);
            fontComboBox.Items.AddRange(FontNames);
            fontSizeComboBox = CreateComboBox(120, true);
            fontSizeComboBox.Items.AddRange(FontSizes);
            alignmentComboBox = CreateComboBox(160, false);
            alignmentComboBox.Items.AddRange(new object[] { "左对齐", "居中", "右对齐", "两端对齐" });
            boldCheckBox = new CheckBox { Text = "加粗", Dock = DockStyle.Left, AutoSize = true };
            lineSpacingComboBox = CreateComboBox(160, false);
            lineSpacingComboBox.Items.AddRange(new object[] { "20磅", "单倍行距" });
            previewLabel = new Label { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, TextAlign = ContentAlignment.MiddleCenter, BackColor = Color.FromArgb(250, 251, 253) };
            BuildEditPage(editPage);

            createStyleCheckedListBox = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, BorderStyle = BorderStyle.FixedSingle, IntegralHeight = false };
            BuildCreatePage(createPage);

            FlowLayoutPanel buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 8, 0, 0) };
            Button cancelButton = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 92, Height = 30 };
            Button okButton = new Button { Text = "创建", DialogResult = DialogResult.OK, Width = 92, Height = 30, BackColor = Color.FromArgb(43, 108, 176), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            okButton.FlatAppearance.BorderSize = 0;
            okButton.Click += okButton_Click;
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(okButton);
            root.Controls.Add(buttons, 0, 1);
            AcceptButton = okButton;
            CancelButton = cancelButton;

            RefreshTemplateList();
            loading = true;
            templateListBox.SelectedIndex = 0;
            loading = false;
            LoadDefinition(1);
        }

        private void BuildEditPage(TabPage page)
        {
            TableLayoutPanel body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 225f));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            page.Controls.Add(body);
            body.Controls.Add(BuildGroup("默认自定义样式", templateListBox), 0, 0);

            TableLayoutPanel fields = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 9, Padding = new Padding(14, 12, 8, 8) };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100f));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            for (int i = 0; i < 8; i++) fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 40f));
            fields.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            AddField(fields, 0, "说明", new Label { Text = "模板可编辑，不能删除", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(96, 96, 96), TextAlign = ContentAlignment.MiddleLeft });
            AddField(fields, 1, "大纲级别", outlineLevelComboBox);
            AddField(fields, 2, "字体", fontComboBox);
            AddField(fields, 3, "字号", fontSizeComboBox);
            AddField(fields, 4, "对齐方式", alignmentComboBox);
            AddField(fields, 5, "字形", boldCheckBox);
            AddField(fields, 6, "行距", lineSpacingComboBox);
            AddField(fields, 7, "预览", previewLabel);
            AddField(fields, 8, "格式", BuildFormatButton());
            body.Controls.Add(BuildGroup("样式编辑", fields), 1, 0);
        }

        private void BuildCreatePage(TabPage page)
        {
            TableLayoutPanel layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 1 };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            page.Controls.Add(layout);

            TableLayoutPanel createPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(8) };
            createPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 26f));
            createPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            createPanel.Controls.Add(new Label { Text = "勾选需要创建到当前文档的自定义样式，可多选", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(96, 96, 96), TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            createPanel.Controls.Add(createStyleCheckedListBox, 0, 1);
            layout.Controls.Add(BuildGroup("选择自定义样式", createPanel), 0, 0);
        }

        private void RefreshTemplateList()
        {
            templateListBox.BeginUpdate();
            try
            {
                templateListBox.Items.Clear();
                createStyleCheckedListBox.Items.Clear();
                foreach (StyleDefinitionRequest definition in definitionsByLevel.Values.Where(item => item.Level <= 10).OrderBy(item => item.Level))
                {
                    templateListBox.Items.Add(new StyleItem(definition.StyleName, definition.Level, false));
                    createStyleCheckedListBox.Items.Add(new StyleItem(definition.StyleName, definition.Level, false), definition.ShouldCreate);
                }
            }
            finally { templateListBox.EndUpdate(); }
        }

        private void templateListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loading) return;
            SaveCurrentDefinition();
            LoadDefinition(templateListBox.SelectedItem is StyleItem item ? item.Level : 1);
        }

        private void LoadDefinition(int level)
        {
            loading = true;
            try
            {
                currentLevel = level;
                StyleDefinitionRequest definition = GetDefinition(level);
                int outlineLevel = definition.OutlineLevel > 0 ? definition.OutlineLevel : definition.Level;
                outlineLevelComboBox.SelectedIndex = outlineLevel == 10 ? 9 : Math.Max(0, Math.Min(8, outlineLevel - 1));
                fontComboBox.Text = definition.FontName;
                fontSizeComboBox.Text = FormatFontSize(definition.FontSize);
                alignmentComboBox.SelectedIndex = Math.Max(0, Math.Min(3, definition.Alignment));
                boldCheckBox.Checked = definition.Bold;
                lineSpacingComboBox.SelectedIndex = Math.Abs(definition.LineSpacing - 1f) < 0.01f ? 1 : 0;
                UpdatePreview(definition);
            }
            finally { loading = false; }
        }

        private void SaveCurrentDefinition()
        {
            if (loading || currentLevel < 1 || currentLevel > 10) return;
            StyleDefinitionRequest definition = GetDefinition(currentLevel);
            definition.OutlineLevel = outlineLevelComboBox.SelectedIndex == 9 ? 10 : outlineLevelComboBox.SelectedIndex + 1;
            definition.FontName = string.IsNullOrWhiteSpace(fontComboBox.Text) ? "宋体" : fontComboBox.Text.Trim();
            definition.FontSize = ParseFontSize(fontSizeComboBox.Text);
            definition.ListFontName = definition.FontName;
            definition.ListFontSize = definition.FontSize;
            definition.Alignment = alignmentComboBox.SelectedIndex < 0 ? 0 : alignmentComboBox.SelectedIndex;
            definition.Bold = boldCheckBox.Checked;
            definition.LineSpacing = lineSpacingComboBox.SelectedIndex == 1 ? 1f : 20f;
            UpdatePreview(definition);
        }

        private Control BuildFormatButton()
        {
            Button button = new Button { Text = "选择字体...", Dock = DockStyle.Left, Width = 110, Height = 28 };
            button.Click += (_, __) => ShowFontDialog();
            return button;
        }

        private void ShowFontDialog()
        {
            using (FontDialog dialog = new FontDialog())
            {
                dialog.Font = new Font(fontComboBox.Text, ParseFontSize(fontSizeComboBox.Text), boldCheckBox.Checked ? FontStyle.Bold : FontStyle.Regular);
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                fontComboBox.Text = dialog.Font.Name;
                fontSizeComboBox.Text = FormatFontSize(dialog.Font.Size);
                boldCheckBox.Checked = dialog.Font.Bold;
                SaveCurrentDefinition();
            }
        }

        private void UpdatePreview(StyleDefinitionRequest definition)
        {
            previewLabel.Text = definition.StyleName + " 示例";
            previewLabel.TextAlign = ToContentAlignment(definition.Alignment);
            try { previewLabel.Font = new Font(definition.FontName, definition.FontSize, definition.Bold ? FontStyle.Bold : FontStyle.Regular); } catch { }
        }

        private StyleDefinitionRequest GetDefinition(int level) => definitionsByLevel[level];

        private void okButton_Click(object sender, EventArgs e)
        {
            SaveCurrentDefinition();
            for (int index = 0; index < createStyleCheckedListBox.Items.Count; index++)
            {
                if (createStyleCheckedListBox.Items[index] is StyleItem item)
                {
                    GetDefinition(item.Level).ShouldCreate = createStyleCheckedListBox.GetItemChecked(index);
                }
            }
            StyleDefinitions.Clear();
            StyleDefinitions.AddRange(definitionsByLevel.Values.OrderBy(item => item.Level).Select(CloneDefinition));
        }

        private static GroupBox BuildGroup(string title, Control content)
        {
            GroupBox group = new GroupBox { Text = title, Dock = DockStyle.Fill, Padding = new Padding(8) };
            group.Controls.Add(content);
            return group;
        }

        private static ComboBox CreateComboBox(int width, bool editable) => new ComboBox { Dock = DockStyle.Left, Width = width, DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList };
        private static void AddField(TableLayoutPanel table, int row, string label, Control editor)
        {
            table.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
            table.Controls.Add(editor, 1, row);
        }
        private static StyleDefinitionRequest CloneDefinition(StyleDefinitionRequest definition) => new StyleDefinitionRequest { Level = definition.Level, OutlineLevel = definition.OutlineLevel, ShouldCreate = definition.ShouldCreate, StyleName = definition.StyleName, FontName = definition.FontName, FontSize = definition.FontSize, ListFontName = definition.ListFontName, ListFontSize = definition.ListFontSize, Bold = definition.Bold, Alignment = definition.Alignment, LineSpacing = definition.LineSpacing };
        private static ContentAlignment ToContentAlignment(int alignment) => alignment == 1 ? ContentAlignment.MiddleCenter : alignment == 2 ? ContentAlignment.MiddleRight : ContentAlignment.MiddleLeft;
        private static string FormatFontSize(float size) => Math.Abs(size - 12f) < 0.01f ? "小四" : Math.Abs(size - 10.5f) < 0.01f ? "五号" : Math.Abs(size - 14f) < 0.01f ? "四号" : size.ToString("0.#");
        private static float ParseFontSize(string text) => (text ?? string.Empty).Trim() == "五号" ? 10.5f : (text ?? string.Empty).Trim() == "小四" ? 12f : (text ?? string.Empty).Trim() == "四号" ? 14f : (text ?? string.Empty).Trim() == "小三" ? 15f : (text ?? string.Empty).Trim() == "三号" ? 16f : float.TryParse(text, out float value) && value > 0 ? value : 12f;

        private sealed class StyleItem
        {
            public StyleItem(string text, int level, bool isDocumentStyle, string styleName = null) { Text = text; Level = level; IsDocumentStyle = isDocumentStyle; StyleName = styleName ?? text; }
            public string Text { get; }
            public int Level { get; }
            public bool IsDocumentStyle { get; }
            public string StyleName { get; }
            public override string ToString() => Text;
        }
    }
}
