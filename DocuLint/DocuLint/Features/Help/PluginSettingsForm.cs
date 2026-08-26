using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class PluginSettingsForm : Form
    {
        private readonly TextBox commonPhraseShortcutBox;
        private readonly TextBox imageCaptionShortcutBox;
        private readonly TextBox tableCaptionShortcutBox;

        internal string CommonPhraseShortcut => commonPhraseShortcutBox.Text.Trim();

        internal string InsertImageCaptionShortcut => imageCaptionShortcutBox.Text.Trim();

        internal string InsertTableCaptionShortcut => tableCaptionShortcutBox.Text.Trim();

        internal PluginSettingsForm(
            string commonPhraseShortcut,
            string imageCaptionShortcut,
            string tableCaptionShortcut)
        {
            Text = "插件配置";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(560, 245);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 5,
                Padding = new Padding(16),
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 132F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 72F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label title = new Label
            {
                Text = "插件快捷键",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 12)
            };
            layout.Controls.Add(title, 0, 0);
            layout.SetColumnSpan(title, 4);

            commonPhraseShortcutBox = AddShortcutRow(
                layout,
                1,
                "常用语补全",
                commonPhraseShortcut);
            imageCaptionShortcutBox = AddShortcutRow(layout, 2, "插入图片题注", imageCaptionShortcut);
            tableCaptionShortcutBox = AddShortcutRow(layout, 3, "插入表格题注", tableCaptionShortcut);

            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = true,
                Margin = new Padding(0, 12, 0, 0)
            };
            Button cancelButton = new Button
            {
                Text = "取消",
                AutoSize = true,
                MinimumSize = new Size(76, 30),
                DialogResult = DialogResult.Cancel
            };
            Button saveButton = new Button
            {
                Text = "保存",
                AutoSize = true,
                MinimumSize = new Size(76, 30),
                DialogResult = DialogResult.OK
            };
            saveButton.Click += (_, __) =>
            {
                commonPhraseShortcutBox.Text = PluginShortcutService.Normalize(commonPhraseShortcutBox.Text);
                imageCaptionShortcutBox.Text = PluginShortcutService.Normalize(imageCaptionShortcutBox.Text);
                tableCaptionShortcutBox.Text = PluginShortcutService.Normalize(tableCaptionShortcutBox.Text);
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(saveButton);
            layout.Controls.Add(buttons, 0, 4);
            layout.SetColumnSpan(buttons, 4);
            Controls.Add(layout);
            AcceptButton = saveButton;
            CancelButton = cancelButton;
        }

        private static TextBox AddShortcutRow(TableLayoutPanel layout, int row, string labelText, string value)
        {
            Label label = new Label
            {
                Text = labelText,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 7, 8, 7)
            };
            TextBox box = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = Color.White,
                Text = PluginShortcutService.Normalize(value),
                TabStop = true,
                Margin = new Padding(0, 4, 8, 4)
            };
            box.KeyDown += (sender, args) =>
            {
                if (args.KeyCode == Keys.Delete || args.KeyCode == Keys.Back)
                {
                    box.Clear();
                }
                else
                {
                    string shortcut = PluginShortcutService.Format(args);
                    if (!string.IsNullOrWhiteSpace(shortcut))
                    {
                        box.Text = shortcut;
                    }
                }

                args.Handled = true;
                args.SuppressKeyPress = true;
            };

            Button clearButton = new Button
            {
                Text = "清除",
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 4, 4)
            };
            clearButton.Click += (_, __) => box.Clear();
            Label hint = new Label
            {
                Text = "按键",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                ForeColor = Color.FromArgb(100, 107, 118),
                Margin = new Padding(4, 7, 0, 7)
            };

            layout.Controls.Add(label, 0, row);
            layout.Controls.Add(box, 1, row);
            layout.Controls.Add(clearButton, 2, row);
            layout.Controls.Add(hint, 3, row);
            return box;
        }
    }
}
