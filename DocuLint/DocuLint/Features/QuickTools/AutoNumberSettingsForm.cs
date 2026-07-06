using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class AutoNumberSettingsForm : Form
    {
        private readonly RadioButton horizontalRadioButton;
        private readonly RadioButton verticalRadioButton;
        private readonly Button okButton;
        private readonly Button cancelButton;

        public AutoNumberSettingsForm(QuickToolsAutoNumberDirection selectedDirection)
        {
            Text = "自动序号设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(320, 150);

            Label tipLabel = new Label
            {
                AutoSize = false,
                Location = new Point(20, 18),
                Size = new Size(280, 36),
                Text = "选择“自动序号”按钮在表格中填充的方向。"
            };

            horizontalRadioButton = new RadioButton
            {
                AutoSize = true,
                Location = new Point(24, 62),
                Text = "水平（从当前单元格向右填充）"
            };

            verticalRadioButton = new RadioButton
            {
                AutoSize = true,
                Location = new Point(24, 88),
                Text = "垂直（从当前单元格向下填充）"
            };

            okButton = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Location = new Point(144, 112),
                Size = new Size(75, 27)
            };

            cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Location = new Point(225, 112),
                Size = new Size(75, 27)
            };

            Controls.Add(tipLabel);
            Controls.Add(horizontalRadioButton);
            Controls.Add(verticalRadioButton);
            Controls.Add(okButton);
            Controls.Add(cancelButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;

            if (selectedDirection == QuickToolsAutoNumberDirection.Vertical)
            {
                verticalRadioButton.Checked = true;
            }
            else
            {
                horizontalRadioButton.Checked = true;
            }
        }

        public QuickToolsAutoNumberDirection SelectedDirection =>
            verticalRadioButton.Checked
                ? QuickToolsAutoNumberDirection.Vertical
                : QuickToolsAutoNumberDirection.Horizontal;
    }
}
