using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint.WpsAddin
{
    internal sealed class QuickLauncherForm : Form
    {
        private readonly Label statusLabel;

        internal QuickLauncherForm()
        {
            Text = "文档不加班 WPS 入口";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            TopMost = false;
            Size = new Size(360, 180);

            Label titleLabel = new Label
            {
                AutoSize = true,
                Text = "文档不加班 快速入口",
                Font = new Font("Segoe UI", 11F, FontStyle.Bold, GraphicsUnit.Point),
                Location = new Point(16, 16)
            };

            statusLabel = new Label
            {
                AutoSize = true,
                Text = "Host status: unknown",
                Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                Location = new Point(16, 50)
            };

            Button pingButton = new Button
            {
                Text = "Test UI Entry",
                Size = new Size(120, 32),
                Location = new Point(16, 90)
            };
            pingButton.Click += PingButton_Click;

            Button closeButton = new Button
            {
                Text = "Close",
                Size = new Size(100, 32),
                Location = new Point(152, 90)
            };
            closeButton.Click += (_, __) => Close();

            Controls.Add(titleLabel);
            Controls.Add(statusLabel);
            Controls.Add(pingButton);
            Controls.Add(closeButton);
        }

        internal void SetHostStatus(string status)
        {
            statusLabel.Text = "Host status: " + (string.IsNullOrWhiteSpace(status) ? "unknown" : status);
        }

        private static void PingButton_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "文档不加班 WPS 入口已激活。",
                "文档不加班",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
