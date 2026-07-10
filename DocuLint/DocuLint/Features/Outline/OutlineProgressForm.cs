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
            BackColor = SystemColors.Window;
            ClientSize = new Size(460, 170);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(22, 16, 22, 16),
                BackColor = SystemColors.Window
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            titleLabel = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = SystemColors.ControlText,
                Text = "正在更新章节号",
                TextAlign = ContentAlignment.MiddleLeft
            };

            messageLabel = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = SystemColors.ControlText,
                Text = "正在准备...",
                TextAlign = ContentAlignment.MiddleLeft
            };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Style = ProgressBarStyle.Blocks,
                Maximum = 100
            };

            detailLabel = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = SystemColors.GrayText,
                Text = "请稍候，Word 正在处理文档结构。",
                TextAlign = ContentAlignment.MiddleLeft
            };

            layout.Controls.Add(titleLabel, 0, 0);
            layout.Controls.Add(messageLabel, 0, 1);
            layout.Controls.Add(progressBar, 0, 2);
            layout.Controls.Add(detailLabel, 0, 3);
            Controls.Add(layout);
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
            progressBar.Value = percent;
            detailLabel.Text = total <= 0
                ? "正在整理文档中的标题段落..."
                : $"已处理 {Math.Max(0, Math.Min(current, total))} / {total}";

            uiRefreshWatch.Restart();
            Refresh();
            Application.DoEvents();
        }
    }
}
