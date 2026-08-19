using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class HeadingStyleRepairForm : Form
    {
        private const string DoNotModifyText = "<不修改>";
        private readonly List<string> cachedStyleNames;
        private readonly HashSet<ComboBox> populatedSelectors = new HashSet<ComboBox>();
        private readonly Dictionary<int, ComboBox> styleSelectors =
            new Dictionary<int, ComboBox>();

        internal HeadingStyleRepairForm(
            IEnumerable<int> headingLevels,
            IEnumerable<string> styleNames)
        {
            Text = "标题样式修复";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            ClientSize = new Size(640, 500);
            Font = new Font("Microsoft YaHei UI", 9F);

            List<int> levels = (headingLevels ?? Enumerable.Empty<int>())
                .Where(level => level >= 1 && level <= 9)
                .Distinct()
                .OrderBy(level => level)
                .ToList();
            cachedStyleNames = NormalizeStyleNames(styleNames);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(18)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label hint = new Label
            {
                AutoSize = true,
                Text = "为需要修复的标题级别手动选择目标样式；保持“不修改”的级别不会处理。",
                Margin = new Padding(0, 0, 0, 12)
            };
            TableLayoutPanel selectorGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                ColumnCount = 2,
                RowCount = Math.Max(1, levels.Count),
                Padding = new Padding(0, 0, 8, 0)
            };
            selectorGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130F));
            selectorGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            for (int row = 0; row < levels.Count; row++)
            {
                int level = levels[row];
                selectorGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, 40F));
                Label levelLabel = new Label
                {
                    Text = level + "级标题",
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 0, 10, 0)
                };
                ComboBox selector = new ComboBox
                {
                    Dock = DockStyle.Fill,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Margin = new Padding(0, 4, 0, 4),
                    Enabled = cachedStyleNames.Count > 0
                };
                selector.Items.Add(DoNotModifyText);
                selector.SelectedIndex = 0;
                selector.DropDown += StyleSelector_DropDown;
                styleSelectors[level] = selector;
                selectorGrid.Controls.Add(levelLabel, 0, row);
                selectorGrid.Controls.Add(selector, 1, row);
            }

            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 12, 0, 0)
            };
            Button confirmButton = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Size = new Size(92, 36),
                Margin = new Padding(8, 0, 0, 0),
                Enabled = cachedStyleNames.Count > 0
            };
            Button cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Size = new Size(92, 36),
                Margin = new Padding(8, 0, 0, 0)
            };
            buttons.Controls.Add(confirmButton);
            buttons.Controls.Add(cancelButton);

            layout.Controls.Add(hint, 0, 0);
            layout.Controls.Add(selectorGrid, 0, 1);
            layout.Controls.Add(buttons, 0, 2);
            Controls.Add(layout);

            AcceptButton = confirmButton;
            CancelButton = cancelButton;
            FormClosing += HeadingStyleRepairForm_FormClosing;
        }

        internal IReadOnlyDictionary<int, string> SelectedStyles
        {
            get
            {
                return styleSelectors
                    .Where(item => item.Value.SelectedIndex > 0)
                    .ToDictionary(item => item.Key, item => Convert.ToString(item.Value.SelectedItem));
            }
        }

        private static List<string> NormalizeStyleNames(IEnumerable<string> styleNames)
        {
            return (styleNames ?? Enumerable.Empty<string>())
                .Select(style => (style ?? string.Empty).Trim())
                .Where(style => style.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(style => style, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private void StyleSelector_DropDown(object sender, EventArgs e)
        {
            ComboBox selector = sender as ComboBox;
            if (selector == null || populatedSelectors.Contains(selector))
            {
                return;
            }

            selector.BeginUpdate();
            try
            {
                selector.Items.AddRange(cachedStyleNames.Cast<object>().ToArray());
                populatedSelectors.Add(selector);
            }
            finally
            {
                selector.EndUpdate();
            }
        }

        private void HeadingStyleRepairForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK && SelectedStyles.Count == 0)
            {
                MessageBox.Show(this, "请至少为一个标题级别选择目标样式。", "标题样式修复");
                e.Cancel = true;
            }
        }
    }
}
