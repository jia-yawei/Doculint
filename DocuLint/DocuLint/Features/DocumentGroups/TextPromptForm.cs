using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class TextPromptForm : Form
    {
        private readonly TextBox inputBox;

        public string ResultText => inputBox.Text.Trim();

        public TextPromptForm(string title, string prompt, string initialValue = "")
        {
            Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(500, 170);

            Label promptLabel = new Label
            {
                AutoSize = false,
                Text = prompt,
                Left = 16,
                Top = 16,
                Width = 468,
                Height = 30
            };

            inputBox = new TextBox
            {
                Left = 16,
                Top = 56,
                Width = 468,
                Text = initialValue ?? string.Empty
            };

            Button okButton = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Left = 310,
                Top = 116,
                Width = 84,
                Height = 32
            };

            Button cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Left = 400,
                Top = 116,
                Width = 84,
                Height = 32
            };

            Controls.Add(promptLabel);
            Controls.Add(inputBox);
            Controls.Add(okButton);
            Controls.Add(cancelButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;
        }
    }
}
