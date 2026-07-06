using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class OutlineProgressForm : Form
    {
        private readonly Label titleLabel;
        private readonly Label messageLabel;
        private readonly Label percentLabel;
        private readonly Label detailLabel;
        private readonly ProgressBar progressBar;
        private readonly Stopwatch uiRefreshWatch = Stopwatch.StartNew();

        public OutlineProgressForm()
        {
            Font = SystemFonts.MessageBoxFont;
            Text = "自动章节号";
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
                Padding = new Padding(26, 22, 26, 22)
            };

            titleLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 40,
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = Color.Black,
                Text = "正在重建自动章节号"
            };

            messageLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 46,
                Margin = new Padding(0, 6, 0, 0),
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
                Margin = new Padding(0, 10, 0, 0),
                ForeColor = Color.Black,
                Text = "扫描文档结构后会开始处理标题编号。"
            };

            progressHost.Controls.Add(progressBar);
            progressHost.Controls.Add(percentLabel);

            cardPanel.Controls.Add(progressHost);
            cardPanel.Controls.Add(detailLabel);
            cardPanel.Controls.Add(messageLabel);
            cardPanel.Controls.Add(titleLabel);

            Controls.Add(cardPanel);
        }

        public void ReportProgress(int current, int total, string message)
        {
            bool shouldRefresh = current <= 1 || current >= total || uiRefreshWatch.ElapsedMilliseconds >= 120;
            if (!shouldRefresh)
            {
                return;
            }

            int percent = total <= 0 ? 0 : Math.Max(0, Math.Min(100, (int)Math.Round(current * 100d / total)));
            messageLabel.Text = string.IsNullOrWhiteSpace(message) ? "正在处理..." : message;
            percentLabel.Text = "进度：" + percent + "%";
            progressBar.Value = percent;
            detailLabel.Text = total <= 0
                ? "正在整理文档中的标题段落..."
                : $"已处理 {Math.Max(0, Math.Min(current, total))} / {total} 个步骤";

            uiRefreshWatch.Restart();
            Refresh();
            Application.DoEvents();
        }
    }
}
