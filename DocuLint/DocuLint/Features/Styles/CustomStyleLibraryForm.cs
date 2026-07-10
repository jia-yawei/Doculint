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
        private readonly ListBox styleListBox;
        private readonly CheckBox shouldCreateCheckBox;
        private readonly TextBox styleNameTextBox;
        private readonly ComboBox outlineLevelComboBox;
        private readonly ComboBox fontComboBox;
        private readonly ComboBox fontSizeComboBox;
        private readonly ComboBox alignmentComboBox;
        private readonly CheckBox boldCheckBox;
        private readonly ComboBox lineSpacingComboBox;
        private readonly Label previewLabel;
        private bool loading;
        private int currentLevel = 1;

        public List<StyleDefinitionRequest> StyleDefinitions { get; } = new List<StyleDefinitionRequest>();

        public CustomStyleLibraryForm(IEnumerable<StyleDefinitionRequest> currentDefinitions)
        {
            SuspendLayout();
            definitionsByLevel = (currentDefinitions ?? StyleDefinitionRequest.CreateDefaultSet())
                .Select(CloneDefinition)
                .GroupBy(item => item.Level)
                .ToDictionary(group => group.Key, group => group.First());

            foreach (StyleDefinitionRequest definition in StyleDefinitionRequest.CreateDefaultSet())
            {
                if (!definitionsByLevel.ContainsKey(definition.Level))
                {
                    definitionsByLevel[definition.Level] = CloneDefinition(definition);
                }
            }

            Text = "创建自定义样式";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(760, 520);
            MinimizeBox = false;
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Font = new Font("Microsoft YaHei UI", 10F);
            Padding = new Padding(12);

            TableLayoutPanel root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            root.SuspendLayout();
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            Controls.Add(root);

            TableLayoutPanel body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
            body.SuspendLayout();
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 210));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            root.Controls.Add(body, 0, 0);

            styleListBox = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false };
            styleListBox.BeginUpdate();
            try
            {
                foreach (StyleDefinitionRequest definition in definitionsByLevel.Values.OrderBy(item => item.Level))
                {
                    styleListBox.Items.Add(new LevelItem(definition.StyleName, definition.Level));
                }
            }
            finally
            {
                styleListBox.EndUpdate();
            }
            styleListBox.SelectedIndexChanged += styleListBox_SelectedIndexChanged;
            body.Controls.Add(BuildGroup("样式", styleListBox), 0, 0);

            shouldCreateCheckBox = new CheckBox { Text = "创建此样式", Dock = DockStyle.Left, AutoSize = true };
            styleNameTextBox = new TextBox { Dock = DockStyle.Left, Width = 260 };
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
            previewLabel = new Label { Dock = DockStyle.Fill, Height = 72, BorderStyle = BorderStyle.FixedSingle, TextAlign = ContentAlignment.MiddleCenter };

            TableLayoutPanel fields = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 10, Padding = new Padding(18, 20, 12, 8) };
            fields.SuspendLayout();
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 9; i++) fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            fields.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            AddField(fields, 0, "是否创建", shouldCreateCheckBox);
            AddField(fields, 1, "名称", styleNameTextBox);
            AddField(fields, 2, "大纲级别", outlineLevelComboBox);
            AddField(fields, 3, "字体", fontComboBox);
            AddField(fields, 4, "字号", fontSizeComboBox);
            AddField(fields, 5, "对齐方式", alignmentComboBox);
            AddField(fields, 6, "加粗", boldCheckBox);
            AddField(fields, 7, "行距", lineSpacingComboBox);
            AddField(fields, 8, "预览", previewLabel);
            AddField(fields, 9, "格式", BuildFormatButton());
            fields.ResumeLayout(false);
            body.Controls.Add(BuildGroup("属性", fields), 1, 0);

            FlowLayoutPanel buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 10, 0, 0) };
            Button cancelButton = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 92, Height = 32 };
            Button okButton = new Button { Text = "确定", DialogResult = DialogResult.OK, Width = 92, Height = 32 };
            okButton.Click += okButton_Click;
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(okButton);
            root.Controls.Add(buttons, 0, 1);
            AcceptButton = okButton;
            CancelButton = cancelButton;

            body.ResumeLayout(false);
            root.ResumeLayout(false);
            loading = true;
            styleListBox.SelectedIndex = 0;
            loading = false;
            LoadDefinition(1);
            ResumeLayout(false);
        }

        private Control BuildFormatButton()
        {
            Button button = new Button { Text = "格式(&O) ▼", Dock = DockStyle.Left, Width = 110, Height = 32 };
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("字体(&F)...", null, (_, __) => ShowFontDialog());
            menu.Items.Add("段落(&P)...", null, (_, __) => MessageBox.Show("段落属性请直接在右侧设置：大纲级别、对齐方式、行距。", "创建自定义样式"));
            button.Click += (_, __) => menu.Show(button, 0, button.Height);
            return button;
        }

        private static GroupBox BuildGroup(string title, Control content)
        {
            GroupBox group = new GroupBox { Text = title, Dock = DockStyle.Fill, Padding = new Padding(8) };
            group.Controls.Add(content);
            return group;
        }

        private static ComboBox CreateComboBox(int width, bool editable)
        {
            return new ComboBox
            {
                Dock = DockStyle.Left,
                Width = width,
                DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList
            };
        }

        private static void AddField(TableLayoutPanel table, int row, string label, Control editor)
        {
            table.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, row);
            table.Controls.Add(editor, 1, row);
        }

        private void styleListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (loading)
            {
                return;
            }

            SaveCurrentDefinition();
            LoadDefinition(styleListBox.SelectedItem is LevelItem item ? item.Level : 1);
        }

        private void LoadDefinition(int level)
        {
            loading = true;
            try
            {
                currentLevel = level;
                StyleDefinitionRequest definition = GetDefinition(level);
                shouldCreateCheckBox.Checked = definition.ShouldCreate;
                styleNameTextBox.Text = definition.StyleName;
                int outlineLevel = definition.OutlineLevel > 0 ? definition.OutlineLevel : definition.Level;
                outlineLevelComboBox.SelectedIndex = outlineLevel == 10 ? 9 : Math.Max(0, Math.Min(8, outlineLevel - 1));
                fontComboBox.Text = definition.FontName;
                fontSizeComboBox.Text = FormatFontSize(definition.FontSize);
                alignmentComboBox.SelectedIndex = Math.Max(0, Math.Min(3, definition.Alignment));
                boldCheckBox.Checked = definition.Bold;
                lineSpacingComboBox.SelectedIndex = Math.Abs(definition.LineSpacing - 1f) < 0.01f ? 1 : 0;
                UpdatePreview(definition);
            }
            finally
            {
                loading = false;
            }
        }

        private void SaveCurrentDefinition()
        {
            if (loading || currentLevel == 0)
            {
                return;
            }

            StyleDefinitionRequest definition = GetDefinition(currentLevel);
            definition.ShouldCreate = shouldCreateCheckBox.Checked;
            definition.StyleName = string.IsNullOrWhiteSpace(styleNameTextBox.Text)
                ? StyleDefinitionRequest.GetDefaultStyleName(currentLevel)
                : styleNameTextBox.Text.Trim();
            definition.OutlineLevel = outlineLevelComboBox.SelectedIndex == 9 ? 10 : outlineLevelComboBox.SelectedIndex + 1;
            definition.FontName = string.IsNullOrWhiteSpace(fontComboBox.Text) ? "宋体" : fontComboBox.Text.Trim();
            definition.FontSize = ParseFontSize(fontSizeComboBox.Text);
            definition.ListFontName = definition.FontName;
            definition.ListFontSize = definition.FontSize;
            definition.Alignment = alignmentComboBox.SelectedIndex < 0 ? 0 : alignmentComboBox.SelectedIndex;
            definition.Bold = boldCheckBox.Checked;
            definition.LineSpacing = lineSpacingComboBox.SelectedIndex == 1 ? 1f : 20f;

            if (styleListBox.SelectedItem is LevelItem item)
            {
                item.Text = definition.StyleName;
            }
        }

        private void ShowFontDialog()
        {
            StyleDefinitionRequest definition = GetDefinition(currentLevel);
            using (FontDialog dialog = new FontDialog())
            {
                dialog.Font = new Font(
                    string.IsNullOrWhiteSpace(fontComboBox.Text) ? definition.FontName : fontComboBox.Text,
                    ParseFontSize(fontSizeComboBox.Text),
                    boldCheckBox.Checked ? FontStyle.Bold : FontStyle.Regular);
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                fontComboBox.Text = dialog.Font.Name;
                fontSizeComboBox.Text = FormatFontSize(dialog.Font.Size);
                boldCheckBox.Checked = dialog.Font.Bold;
                SaveCurrentDefinition();
            }
        }

        private void UpdatePreview(StyleDefinitionRequest definition)
        {
            previewLabel.Text = definition.StyleName;
            previewLabel.TextAlign = ToContentAlignment(definition.Alignment);
            try
            {
                previewLabel.Font = new Font(definition.FontName, definition.FontSize, definition.Bold ? FontStyle.Bold : FontStyle.Regular);
            }
            catch
            {
            }
        }

        private StyleDefinitionRequest GetDefinition(int level)
        {
            if (!definitionsByLevel.TryGetValue(level, out StyleDefinitionRequest definition))
            {
                definition = StyleDefinitionRequest.CreateDefaultSet().First(item => item.Level == level);
                definitionsByLevel[level] = definition;
            }
            return definition;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            SaveCurrentDefinition();
            StyleDefinitions.Clear();
            StyleDefinitions.AddRange(definitionsByLevel.Values.OrderBy(item => item.Level).Select(CloneDefinition));
        }

        private static StyleDefinitionRequest CloneDefinition(StyleDefinitionRequest definition)
        {
            return new StyleDefinitionRequest
            {
                Level = definition.Level,
                OutlineLevel = definition.OutlineLevel,
                ShouldCreate = definition.ShouldCreate,
                StyleName = definition.StyleName,
                FontName = definition.FontName,
                FontSize = definition.FontSize,
                ListFontName = definition.ListFontName,
                ListFontSize = definition.ListFontSize,
                Bold = definition.Bold,
                Alignment = definition.Alignment,
                LineSpacing = definition.LineSpacing
            };
        }

        private static ContentAlignment ToContentAlignment(int alignment)
        {
            switch (alignment)
            {
                case 1: return ContentAlignment.MiddleCenter;
                case 2: return ContentAlignment.MiddleRight;
                default: return ContentAlignment.MiddleLeft;
            }
        }

        private static string FormatFontSize(float size) => Math.Abs(size - 12f) < 0.01f ? "小四" : Math.Abs(size - 10.5f) < 0.01f ? "五号" : Math.Abs(size - 14f) < 0.01f ? "四号" : size.ToString("0.#");

        private static float ParseFontSize(string text)
        {
            switch ((text ?? string.Empty).Trim())
            {
                case "五号": return 10.5f;
                case "小四": return 12f;
                case "四号": return 14f;
                case "小三": return 15f;
                case "三号": return 16f;
                default: return float.TryParse(text, out float value) && value > 0 ? value : 12f;
            }
        }

        private sealed class LevelItem
        {
            public LevelItem(string text, int level) { Text = text; Level = level; }
            public string Text { get; set; }
            public int Level { get; }
            public override string ToString() => Text;
        }
    }
}
