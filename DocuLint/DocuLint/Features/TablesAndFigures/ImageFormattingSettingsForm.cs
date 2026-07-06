using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class ImageFormattingOptions
    {
        public int ScalePercent { get; set; }

        public static ImageFormattingOptions CreateDefault()
        {
            return new ImageFormattingOptions
            {
                ScalePercent = 100
            };
        }

        public ImageFormattingOptions Clone()
        {
            return new ImageFormattingOptions
            {
                ScalePercent = ScalePercent
            };
        }
    }

    internal sealed class ImageFormattingSettingsForm : Form
    {
        private readonly RadioButton scale50Radio;
        private readonly RadioButton scale75Radio;
        private readonly RadioButton scale100Radio;
        private readonly Button cancelButton;
        private readonly Button editButton;
        private readonly Button confirmButton;

        public ImageFormattingOptions Options { get; private set; }

        public ImageFormattingSettingsForm(ImageFormattingOptions options)
        {
            Options = (options ?? ImageFormattingOptions.CreateDefault()).Clone();

            Font = SystemFonts.MessageBoxFont;
            Text = "规范全部图片";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;
            ClientSize = new Size(560, 300);

            Label titleLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 48,
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = Color.Black,
                Text = "规范全部图片参数",
                Padding = new Padding(22, 14, 0, 0)
            };

            Label infoLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 68,
                ForeColor = Color.Black,
                Text = "默认将主文档中的全部图片统一设置为居中、单倍行距、左右缩进 0、首行缩进 0、特殊格式为无。\r\n请选择图片缩放比例，默认 100%。",
                Padding = new Padding(22, 2, 22, 0)
            };

            GroupBox scaleGroup = new GroupBox
            {
                Text = "图片大小",
                Dock = DockStyle.Top,
                Height = 92,
                Padding = new Padding(18, 14, 18, 10)
            };

            FlowLayoutPanel radioPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false
            };

            scale50Radio = new RadioButton
            {
                Text = "50%",
                AutoSize = true,
                Margin = new Padding(8, 18, 24, 0)
            };
            scale75Radio = new RadioButton
            {
                Text = "75%",
                AutoSize = true,
                Margin = new Padding(8, 18, 24, 0)
            };
            scale100Radio = new RadioButton
            {
                Text = "100%",
                AutoSize = true,
                Margin = new Padding(8, 18, 24, 0)
            };

            radioPanel.Controls.Add(scale50Radio);
            radioPanel.Controls.Add(scale75Radio);
            radioPanel.Controls.Add(scale100Radio);
            scaleGroup.Controls.Add(radioPanel);

            Panel buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 72,
                Padding = new Padding(0, 12, 22, 16),
                BackColor = Color.White
            };

            confirmButton = new Button
            {
                Text = "确认",
                Size = new Size(84, 30),
                Dock = DockStyle.Right
            };
            confirmButton.Click += ConfirmButton_Click;

            editButton = new Button
            {
                Text = "修改",
                Size = new Size(84, 30),
                Dock = DockStyle.Right
            };
            editButton.Click += EditButton_Click;

            cancelButton = new Button
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
            buttonPanel.Controls.Add(editButton);
            buttonPanel.Controls.Add(cancelButton);

            Controls.Add(buttonPanel);
            Controls.Add(scaleGroup);
            Controls.Add(infoLabel);
            Controls.Add(titleLabel);

            ApplyCurrentScaleSelection();
            SetEditingEnabled(false);
        }

        private void EditButton_Click(object sender, EventArgs e)
        {
            SetEditingEnabled(true);
            scale100Radio.Focus();
            editButton.Enabled = false;
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            Options = new ImageFormattingOptions
            {
                ScalePercent = scale50Radio.Checked ? 50 : (scale75Radio.Checked ? 75 : 100)
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        private void ApplyCurrentScaleSelection()
        {
            scale50Radio.Checked = Options.ScalePercent == 50;
            scale75Radio.Checked = Options.ScalePercent == 75;
            scale100Radio.Checked = !scale50Radio.Checked && !scale75Radio.Checked;
        }

        private void SetEditingEnabled(bool enabled)
        {
            scale50Radio.Enabled = enabled;
            scale75Radio.Enabled = enabled;
            scale100Radio.Enabled = enabled;
        }
    }
}
