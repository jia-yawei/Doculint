using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class TableFormattingProgressForm : Form
    {
        private readonly Label titleLabel;
        private readonly Label messageLabel;
        private readonly Label percentLabel;
        private readonly Label detailLabel;
        private readonly PictureBox statusIcon;
        private readonly ProgressBar progressBar;
        private readonly Button closeButton;
        private readonly Panel buttonPanel;
        private readonly Stopwatch uiRefreshWatch = Stopwatch.StartNew();

        public bool IsFinalized { get; private set; }

        public TableFormattingProgressForm()
        {
            Font = SystemFonts.MessageBoxFont;
            Text = "一键规范表格";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            MaximizeBox = false;
            MinimizeBox = false;
            ControlBox = false;
            BackColor = Color.White;
            ClientSize = new Size(520, 260);

            Panel cardPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(26, 22, 26, 0)
            };

            buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 50,
                BackColor = Color.White,
                Padding = new Padding(0, 7, 22, 9)
            };

            Panel titlePanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 42,
                BackColor = Color.White
            };

            statusIcon = new PictureBox
            {
                Location = new Point(0, 2),
                Size = new Size(28, 28),
                SizeMode = PictureBoxSizeMode.CenterImage,
                Visible = false
            };

            titleLabel = new Label
            {
                Location = new Point(34, 0),
                Size = new Size(430, 38),
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = Color.Black,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "一键规范表格"
            };

            messageLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 46,
                ForeColor = Color.Black,
                Text = "正在准备..."
            };

            Panel progressHost = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                Padding = new Padding(0, 12, 0, 0)
            };

            progressBar = new ProgressBar
            {
                Location = new Point(0, 13),
                Size = new Size(320, 16),
                Style = ProgressBarStyle.Continuous,
                Maximum = 100
            };

            percentLabel = new Label
            {
                Location = new Point(336, 7),
                Size = new Size(130, 28),
                TextAlign = ContentAlignment.MiddleRight,
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = Color.Black,
                Text = "进度：0%"
            };

            detailLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 66,
                ForeColor = Color.Black,
                Text = "会按当前参数统一处理目录后的所有表格。"
            };

            closeButton = new Button
            {
                Dock = DockStyle.Right,
                Size = new Size(82, 28),
                Text = "确定",
                Visible = false
            };
            closeButton.Click += (sender, args) => Close();

            titlePanel.Controls.Add(statusIcon);
            titlePanel.Controls.Add(titleLabel);
            progressHost.Controls.Add(progressBar);
            progressHost.Controls.Add(percentLabel);
            buttonPanel.Controls.Add(closeButton);

            cardPanel.Controls.Add(progressHost);
            cardPanel.Controls.Add(buttonPanel);
            cardPanel.Controls.Add(detailLabel);
            cardPanel.Controls.Add(messageLabel);
            cardPanel.Controls.Add(titlePanel);

            Controls.Add(cardPanel);
        }

        public void ReportProgress(int percent, string message, string detail)
        {
            bool shouldRefresh = percent <= 5
                || percent >= 95
                || percent >= progressBar.Value + 3
                || uiRefreshWatch.ElapsedMilliseconds >= 120;
            if (!shouldRefresh)
            {
                return;
            }

            UpdateContent(percent, message, detail);
        }

        public void Complete(string message, string detail, bool success)
        {
            IsFinalized = true;
            titleLabel.Text = success ? "规范表格完成" : "规范表格失败";
            statusIcon.Image = success ? SystemIcons.Information.ToBitmap() : SystemIcons.Error.ToBitmap();
            statusIcon.Visible = true;
            closeButton.Visible = true;
            closeButton.Focus();
            ControlBox = true;
            UpdateContent(100, message, detail);
        }

        public void WaitForUserClose()
        {
            while (!IsDisposed && Visible)
            {
                Application.DoEvents();
                Thread.Sleep(15);
            }
        }

        private void UpdateContent(int percent, string message, string detail)
        {
            percent = Math.Max(0, Math.Min(100, percent));
            messageLabel.Text = string.IsNullOrWhiteSpace(message) ? "正在处理..." : message;
            detailLabel.Text = string.IsNullOrWhiteSpace(detail) ? "请稍候，正在处理当前文档。" : detail;
            percentLabel.Text = "进度：" + percent + "%";
            progressBar.Value = percent;
            uiRefreshWatch.Restart();
            Refresh();
            Application.DoEvents();
        }
    }
}
