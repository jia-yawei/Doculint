using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class SequenceFieldForm : Form
    {
        private readonly TextBox identifierTextBox;

        internal SequenceFieldForm(string selectedNumber)
        {
            Text = "插入域编号";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            ClientSize = new Size(560, 230);
            Font = new Font("Microsoft YaHei UI", 9F);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(16)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label hint = new Label
            {
                AutoSize = true,
                Text = $"已选编号：{selectedNumber}。请输入 SEQ 域名，例如 需求标识：",
                Margin = new Padding(0, 0, 0, 8)
            };
            identifierTextBox = new TextBox
            {
                Dock = DockStyle.Top,
                Height = 28,
                Text = "需求标识",
                Margin = new Padding(0, 0, 0, 8)
            };
            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                MinimumSize = new Size(0, 40),
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 0)
            };
            Button confirmButton = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Size = new Size(92, 36),
                Margin = new Padding(8, 0, 0, 0)
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
            layout.Controls.Add(identifierTextBox, 0, 1);
            layout.Controls.Add(new Label(), 0, 2);
            layout.Controls.Add(buttons, 0, 3);
            Controls.Add(layout);

            AcceptButton = confirmButton;
            CancelButton = cancelButton;
        }

        internal string SequenceIdentifier => (identifierTextBox.Text ?? string.Empty).Trim();
    }
}
