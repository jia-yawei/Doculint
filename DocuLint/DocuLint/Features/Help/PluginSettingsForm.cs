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
        private readonly TextBox standardLibraryDirectoryBox;
        private readonly bool requireStandardLibraryDirectory;

        internal string CommonPhraseShortcut => commonPhraseShortcutBox.Text.Trim();

        internal string InsertImageCaptionShortcut => imageCaptionShortcutBox.Text.Trim();

        internal string InsertTableCaptionShortcut => tableCaptionShortcutBox.Text.Trim();

        internal string StandardLibraryDirectory => standardLibraryDirectoryBox.Text.Trim();

        internal PluginSettingsForm(
            string commonPhraseShortcut,
            string imageCaptionShortcut,
            string tableCaptionShortcut,
            string standardLibraryDirectory)
        {
            requireStandardLibraryDirectory = string.IsNullOrWhiteSpace(standardLibraryDirectory);
            Text = "插件配置";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(640, 370);
            MinimumSize = new Size(640, 370);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 6,
                Padding = new Padding(16),
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 132F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 84F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label title = new Label
            {
                Text = "插件快捷键与标准资料",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 12)
            };
            layout.Controls.Add(title, 0, 0);
            layout.SetColumnSpan(title, 3);

            commonPhraseShortcutBox = AddShortcutRow(
                layout,
                1,
                "常用语补全",
                commonPhraseShortcut);
            imageCaptionShortcutBox = AddShortcutRow(layout, 2, "插入图片题注", imageCaptionShortcut);
            tableCaptionShortcutBox = AddShortcutRow(layout, 3, "插入表格题注", tableCaptionShortcut);
            standardLibraryDirectoryBox = AddStandardLibraryRow(layout, 4, standardLibraryDirectory);

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
            };
            saveButton.Click += (_, __) =>
            {
                commonPhraseShortcutBox.Text = PluginShortcutService.Normalize(commonPhraseShortcutBox.Text);
                imageCaptionShortcutBox.Text = PluginShortcutService.Normalize(imageCaptionShortcutBox.Text);
                tableCaptionShortcutBox.Text = PluginShortcutService.Normalize(tableCaptionShortcutBox.Text);
                if (string.IsNullOrWhiteSpace(StandardLibraryDirectory))
                {
                    MessageBox.Show(
                        requireStandardLibraryDirectory
                            ? "首次使用前请先指定标准文件夹。"
                            : "请指定标准文件夹。",
                        "插件配置",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
                if (!System.IO.Directory.Exists(StandardLibraryDirectory))
                {
                    MessageBox.Show("指定的标准文件夹不存在，请重新选择。", "插件配置", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult = DialogResult.OK;
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(saveButton);
            layout.Controls.Add(buttons, 0, 5);
            layout.SetColumnSpan(buttons, 3);
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
            layout.Controls.Add(label, 0, row);
            layout.Controls.Add(box, 1, row);
            layout.Controls.Add(clearButton, 2, row);
            return box;
        }

        private TextBox AddStandardLibraryRow(TableLayoutPanel layout, int row, string directory)
        {
            Label label = new Label
            {
                Text = "标准文件夹",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 7, 8, 7)
            };
            TextBox box = new TextBox
            {
                Dock = DockStyle.Fill,
                Text = directory ?? string.Empty,
                Margin = new Padding(0, 4, 8, 4)
            };
            Button browseButton = new Button
            {
                Text = "浏览...",
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 4, 4, 4)
            };
            browseButton.Click += (_, __) =>
            {
                using (FolderBrowserDialog dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "选择存放软件开发标准文件的文件夹";
                    dialog.ShowNewFolderButton = false;
                    if (System.IO.Directory.Exists(box.Text))
                    {
                        dialog.SelectedPath = box.Text;
                    }

                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        box.Text = dialog.SelectedPath;
                    }
                }
            };
            layout.Controls.Add(label, 0, row);
            layout.Controls.Add(box, 1, row);
            layout.Controls.Add(browseButton, 2, row);
            return box;
        }
    }
}
