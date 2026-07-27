using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class StyleLinkSettingsForm : Form
    {
        private readonly List<string> systemStyleNames;
        private readonly Func<IEnumerable<string>> styleLoader;
        private readonly Dictionary<string, string> boundStyleDisplayNames =
            new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
        private readonly Dictionary<int, string> styleLinks = new Dictionary<int, string>();
        private readonly Dictionary<int, StyleDefinitionRequest> definitionsByLevel = new Dictionary<int, StyleDefinitionRequest>();
        private readonly ComboBox boundStyleComboBox;
        private readonly ComboBox numberPatternComboBox;
        private readonly ComboBox numberSpacingComboBox;
        private readonly Label listFontLabel;
        private readonly Label previewLabel;
        private readonly FlowLayoutPanel outlineLevelPanel;
        private readonly Button okButton;
        private readonly Button cancelButton;
        private readonly List<Button> outlineLevelButtons = new List<Button>();
        private bool loadingBinding;
        private bool documentStylesLoaded;
        private int currentBindingLevel = 1;

        public Dictionary<int, string> StyleLinks { get; } = new Dictionary<int, string>();

        public OutlineNumberPattern NumberPattern { get; private set; }

        public int NumberTextSpacing { get; private set; }

        public List<StyleDefinitionRequest> StyleDefinitions { get; } = new List<StyleDefinitionRequest>();

        public StyleLinkSettingsForm(
            IDictionary<int, string> currentLinks,
            IEnumerable<string> documentStyleNames,
            Func<IEnumerable<string>> lazyStyleLoader,
            IEnumerable<string> customStyles,
            OutlineNumberPattern currentNumberPattern,
            int currentNumberTextSpacing,
            IEnumerable<StyleDefinitionRequest> currentDefinitions)
        {
            foreach (KeyValuePair<int, string> item in currentLinks ?? new Dictionary<int, string>())
            {
                if (!string.IsNullOrWhiteSpace(item.Value))
                {
                    styleLinks[item.Key] = item.Value;
                }
            }

            foreach (StyleDefinitionRequest definition in currentDefinitions ?? StyleDefinitionRequest.CreateDefaultSet())
            {
                if (definition != null)
                {
                    definitionsByLevel[definition.Level] = CloneDefinition(definition);
                }
            }

            foreach (StyleDefinitionRequest definition in StyleDefinitionRequest.CreateDefaultSet())
            {
                if (!definitionsByLevel.ContainsKey(definition.Level))
                {
                    definitionsByLevel[definition.Level] = CloneDefinition(definition);
                }
            }

            styleLoader = lazyStyleLoader;
            systemStyleNames = new List<string>();
            AddStyleNames(documentStyleNames);
            AddStyleNames(customStyles);

            NumberPattern = currentNumberPattern;
            NumberTextSpacing = Math.Max(0, Math.Min(2, currentNumberTextSpacing));
            Text = "样式绑定";
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(800, 600);
            MinimizeBox = false;
            MaximizeBox = false;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            Font = new Font("Microsoft YaHei UI", 10.5F);
            BackColor = SystemColors.Window;
            Padding = new Padding(14);
            SuspendLayout();

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = SystemColors.Window
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
            Controls.Add(root);

            Panel bindingPage = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            root.Controls.Add(bindingPage, 0, 0);

            outlineLevelPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 56,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(8, 11, 8, 0),
                BackColor = SystemColors.Window
            };

            boundStyleComboBox = CreateComboBox(360, true);
            boundStyleComboBox.DropDown += boundStyleComboBox_DropDown;
            RefreshBoundStyleItems(string.Empty);

            numberPatternComboBox = CreateComboBox(270, false);
            numberPatternComboBox.Items.Add(new NumberPatternItem("1 / 1.1 / 1.1.1", OutlineNumberPattern.Decimal));
            numberPatternComboBox.Items.Add(new NumberPatternItem("1. / 1.1. / 1.1.1.", OutlineNumberPattern.Dotted));
            numberPatternComboBox.Items.Add(new NumberPatternItem("(1) / (1.1)", OutlineNumberPattern.Parenthesized));
            numberPatternComboBox.SelectedIndexChanged += (_, __) =>
            {
                NumberPatternItem item = numberPatternComboBox.SelectedItem as NumberPatternItem;
                if (item != null)
                {
                    NumberPattern = item.Pattern;
                    UpdatePreview();
                }
            };
            SelectNumberPattern(currentNumberPattern);

            numberSpacingComboBox = CreateComboBox(150, false);
            numberSpacingComboBox.Items.AddRange(new object[] { "无", "一个空格", "两个空格" });
            numberSpacingComboBox.SelectedIndex = NumberTextSpacing;
            numberSpacingComboBox.SelectedIndexChanged += (_, __) =>
            {
                NumberTextSpacing = Math.Max(0, numberSpacingComboBox.SelectedIndex);
                UpdatePreview();
            };

            listFontLabel = CreateValueLabel();
            previewLabel = new Label
            {
                Dock = DockStyle.Left,
                Width = 410,
                Height = 34,
                TextAlign = ContentAlignment.MiddleLeft
            };
            Button listFontButton = new Button { Text = "字体...", Width = 88, Height = 32 };
            listFontButton.Click += listFontButton_Click;
            bindingPage.Controls.Add(BuildBindingPage(listFontButton));
            SelectOutlineLevel(1);

            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 8, 0, 0),
                BackColor = SystemColors.Window
            };

            cancelButton = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 90, Height = 32 };
            okButton = new Button { Text = "确定", DialogResult = DialogResult.OK, Width = 90, Height = 32 };
            okButton.Click += okButton_Click;
            buttonPanel.Controls.Add(cancelButton);
            buttonPanel.Controls.Add(okButton);
            root.Controls.Add(buttonPanel, 0, 1);

            AcceptButton = okButton;
            CancelButton = cancelButton;
            ResumeLayout(false);
        }

        private Control BuildBindingPage(Button listFontButton)
        {
            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = SystemColors.Window,
                Padding = new Padding(14, 8, 14, 8)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 58));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            root.Controls.Add(outlineLevelPanel, 0, 0);
            outlineLevelPanel.Controls.Add(new Label
            {
                Text = "大纲级别",
                AutoSize = false,
                Width = 96,
                Height = 34,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 8, 0)
            });
            for (int level = 1; level <= 9; level++)
            {
                AddOutlineLevelButton(level + "级", level);
            }
            AddOutlineLevelButton("正文", 10);

            GroupBox propertyGroup = new GroupBox { Text = "绑定设置", Dock = DockStyle.Fill };
            TableLayoutPanel fields = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                BackColor = SystemColors.Window,
                Padding = new Padding(28, 18, 22, 10)
            };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 5; i++) fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            fields.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            AddField(fields, 0, "绑定标题样式", boundStyleComboBox);
            AddField(fields, 1, "多级列表样式", numberPatternComboBox);
            AddField(fields, 2, "编号后间隔", numberSpacingComboBox);
            AddField(fields, 3, "编号字体", BuildInlinePanel(listFontLabel, listFontButton));
            AddField(fields, 4, "预览", previewLabel);
            AddField(fields, 5, "说明", CreateHintLabel("编号字体字号可单独设置；正文不参与编号。"));
            propertyGroup.Controls.Add(fields);
            root.Controls.Add(propertyGroup, 0, 1);
            return root;
        }

        private void AddOutlineLevelButton(string text, int level)
        {
            Button button = new Button
            {
                Text = text,
                Width = level == 10 ? 64 : 52,
                Height = 34,
                Tag = level,
                Margin = new Padding(0, 0, 6, 0),
                FlatStyle = FlatStyle.Flat
            };
            button.FlatAppearance.BorderSize = 0;
            button.Click += (_, __) => SelectOutlineLevel((int)button.Tag);
            outlineLevelButtons.Add(button);
            outlineLevelPanel.Controls.Add(button);
        }

        private void SelectOutlineLevel(int level)
        {
            SaveCurrentBindingLevel();
            LoadBindingLevel(level);
            foreach (Button button in outlineLevelButtons)
            {
                button.BackColor = (int)button.Tag == level ? SystemColors.Highlight : SystemColors.Control;
                button.ForeColor = (int)button.Tag == level ? SystemColors.HighlightText : SystemColors.ControlText;
            }
        }

        private static ComboBox CreateComboBox(int width, bool editable)
        {
            return new ComboBox
            {
                Dock = DockStyle.Left,
                Width = width,
                Height = 32,
                DropDownStyle = editable ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList,
                AutoCompleteMode = AutoCompleteMode.None,
                AutoCompleteSource = AutoCompleteSource.None
            };
        }

        private static FlowLayoutPanel BuildInlinePanel(params Control[] controls)
        {
            FlowLayoutPanel panel = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = Padding.Empty, Padding = Padding.Empty };
            foreach (Control control in controls) panel.Controls.Add(control);
            return panel;
        }

        private static void AddField(TableLayoutPanel table, int row, string label, Control editor)
        {
            table.Controls.Add(new Label { Text = label, Dock = DockStyle.Fill, Padding = new Padding(4, 0, 8, 0), TextAlign = ContentAlignment.MiddleLeft }, 0, row);
            table.Controls.Add(editor, 1, row);
        }

        private static Label CreateValueLabel() => new Label { Dock = DockStyle.Left, Width = 200, Height = 34, TextAlign = ContentAlignment.MiddleLeft };

        private static Label CreateHintLabel(string text) => new Label { Text = text, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = SystemColors.GrayText };

        private void LoadBindingLevel(int level)
        {
            loadingBinding = true;
            try
            {
                currentBindingLevel = level;
                string styleName = styleLinks.TryGetValue(level, out string linkedStyleName) ? linkedStyleName : string.Empty;
                RefreshBoundStyleItems(styleName);
                UpdateListFontLabel(GetDefinition(level));
                UpdatePreview();
            }
            finally
            {
                loadingBinding = false;
            }
        }

        private void SaveCurrentBindingLevel()
        {
            if (loadingBinding || currentBindingLevel == 0) return;
            string styleName = GetSelectedBoundStyleName();
            if (string.IsNullOrWhiteSpace(styleName)) styleLinks.Remove(currentBindingLevel); else styleLinks[currentBindingLevel] = styleName;
        }

        private void RefreshBoundStyleItems(string keepText)
        {
            string text = keepText ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(text))
            {
                AddStyleName(text);
            }

            boundStyleComboBox.BeginUpdate();
            try
            {
                boundStyleComboBox.Items.Clear();
                boundStyleDisplayNames.Clear();
                foreach (string name in GetVisibleStyleNames())
                {
                    string displayName = GetBoundStyleDisplayName(name);
                    boundStyleDisplayNames[displayName] = name;
                    boundStyleComboBox.Items.Add(displayName);
                }

                boundStyleComboBox.Text = string.IsNullOrWhiteSpace(text) ? string.Empty : GetBoundStyleDisplayName(text);
            }
            finally
            {
                boundStyleComboBox.EndUpdate();
            }
        }

        private void boundStyleComboBox_DropDown(object sender, EventArgs e)
        {
            string keepText = boundStyleComboBox.Text;
            EnsureDocumentStylesLoaded();
            RefreshBoundStyleItems(keepText);
        }

        private void EnsureDocumentStylesLoaded()
        {
            if (styleLoader == null || documentStylesLoaded)
            {
                return;
            }

            Cursor previousCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                AddStyleNames(styleLoader());
                documentStylesLoaded = true;
            }
            catch
            {
            }
            finally
            {
                Cursor.Current = previousCursor;
            }
        }

        private IEnumerable<string> GetVisibleStyleNames()
        {
            return systemStyleNames.OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase);
        }

        private void AddStyleNames(IEnumerable<string> names)
        {
            foreach (string name in names ?? Enumerable.Empty<string>())
            {
                AddStyleName(name);
            }
        }

        private void AddStyleName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)
                || systemStyleNames.Any(item => string.Equals(item, name, StringComparison.CurrentCultureIgnoreCase)))
            {
                return;
            }

            systemStyleNames.Add(name.Trim());
        }

        private string GetSelectedBoundStyleName()
        {
            string text = (boundStyleComboBox.Text ?? string.Empty).Trim();
            return boundStyleDisplayNames.TryGetValue(text, out string styleName)
                ? styleName
                : text;
        }

        private string GetBoundStyleDisplayName(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
            {
                return string.Empty;
            }

            return styleName;
        }

        private void listFontButton_Click(object sender, EventArgs e)
        {
            StyleDefinitionRequest definition = GetDefinition(currentBindingLevel);
            using (FontDialog dialog = new FontDialog())
            {
                dialog.Font = new Font(GetListFontName(definition), GetListFontSize(definition));
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                definition.ListFontName = dialog.Font.Name;
                definition.ListFontSize = dialog.Font.Size;
                UpdateListFontLabel(definition);
                UpdatePreview();
            }
        }

        private void UpdateListFontLabel(StyleDefinitionRequest definition)
        {
            listFontLabel.Text = $"{GetListFontName(definition)} {FormatFontSize(GetListFontSize(definition))}";
        }

        private void UpdatePreview()
        {
            if (previewLabel == null)
            {
                return;
            }

            StyleDefinitionRequest definition = GetDefinition(currentBindingLevel);
            previewLabel.Text = currentBindingLevel == 10
                ? "正文不参与编号"
                : BuildNumberPreview(currentBindingLevel, NumberPattern) + new string(' ', NumberTextSpacing) + "标题示例";
            try
            {
                previewLabel.Font = new Font(GetListFontName(definition), GetListFontSize(definition));
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
                definitionsByLevel[level] = CloneDefinition(definition);
            }
            return definition;
        }

        private void SelectNumberPattern(OutlineNumberPattern pattern)
        {
            foreach (object item in numberPatternComboBox.Items)
            {
                if (item is NumberPatternItem patternItem && patternItem.Pattern == pattern)
                {
                    numberPatternComboBox.SelectedItem = item;
                    return;
                }
            }
            if (numberPatternComboBox.Items.Count > 0) numberPatternComboBox.SelectedIndex = 0;
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            SaveCurrentBindingLevel();
            StyleLinks.Clear();
            foreach (KeyValuePair<int, string> item in styleLinks)
            {
                if (!string.IsNullOrWhiteSpace(item.Value)) StyleLinks[item.Key] = item.Value.Trim();
            }
            StyleDefinitions.Clear();
            for (int level = 1; level <= 10; level++) StyleDefinitions.Add(CloneDefinition(GetDefinition(level)));
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
                ListFontName = GetListFontName(definition),
                ListFontSize = GetListFontSize(definition),
                Alignment = definition.Alignment,
                Bold = definition.Bold,
                LineSpacing = definition.LineSpacing
            };
        }

        private static string GetListFontName(StyleDefinitionRequest definition) => string.IsNullOrWhiteSpace(definition.ListFontName) ? (string.IsNullOrWhiteSpace(definition.FontName) ? "宋体" : definition.FontName) : definition.ListFontName;

        private static float GetListFontSize(StyleDefinitionRequest definition) => definition.ListFontSize > 0f ? definition.ListFontSize : definition.FontSize;

        private static string BuildNumberPreview(int level, OutlineNumberPattern pattern)
        {
            string text = string.Join(".", Enumerable.Repeat("1", level));
            if (pattern == OutlineNumberPattern.Parenthesized)
            {
                return "(" + text + ")";
            }

            return pattern == OutlineNumberPattern.Dotted ? text + "." : text;
        }

        private static string FormatFontSize(float size) => Math.Abs(size - 12f) < 0.01f ? "小四" : Math.Abs(size - 10.5f) < 0.01f ? "五号" : Math.Abs(size - 14f) < 0.01f ? "四号" : size.ToString("0.#");

        private sealed class NumberPatternItem
        {
            public NumberPatternItem(string text, OutlineNumberPattern pattern) { Text = text; Pattern = pattern; }
            public string Text { get; }
            public OutlineNumberPattern Pattern { get; }
            public override string ToString() => Text;
        }
    }
}
