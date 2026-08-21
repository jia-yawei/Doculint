using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class PluginUpdateForm : Form
    {
        private readonly TextBox localFolderBox;
        private readonly Label statusHeadingLabel;
        private readonly Label versionSummaryLabel;
        private readonly Label statusLabel;
        private readonly ProgressBar updateProgressBar;
        private readonly Button installButton;
        private readonly Button checkButton;
        private readonly Button cancelButton;
        private bool updateInProgress;
        private PluginUpdateManifest latestManifest;
        private string latestSource;

        internal PluginUpdateForm()
        {
            Text = "检查更新";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(640, 420);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 7,
                Padding = new Padding(16)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 148F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Label title = new Label
            {
                Text = "搞快点更新",
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 12)
            };
            layout.Controls.Add(title, 0, 0);
            layout.SetColumnSpan(title, 3);

            layout.Controls.Add(new Label { Text = "当前版本", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            layout.Controls.Add(new Label { Text = PluginUpdateService.CurrentVersionText, AutoSize = true, Anchor = AnchorStyles.Left }, 1, 1);

            layout.Controls.Add(new Label { Text = "GitHub 更新源", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
            Label githubSourceLabel = new Label
            {
                Text = "已内置（Doculint 官方仓库）",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                ForeColor = Color.FromArgb(75, 82, 95)
            };
            layout.Controls.Add(githubSourceLabel, 1, 2);

            layout.Controls.Add(new Label { Text = "指定文件夹", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
            localFolderBox = new TextBox { Dock = DockStyle.Fill, Text = GetLocalFolder() };
            layout.Controls.Add(localFolderBox, 1, 3);
            Button browseButton = new Button { Text = "选择...", AutoSize = true, Dock = DockStyle.Fill };
            browseButton.Click += (_, __) => BrowseLocalFolder();
            layout.Controls.Add(browseButton, 2, 3);

            Panel statusPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.FromArgb(247, 249, 251),
                Padding = new Padding(14, 10, 14, 10),
            };
            TableLayoutPanel statusLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                BackColor = Color.Transparent,
            };
            statusLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            statusLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            statusLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            statusLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            statusHeadingLabel = new Label
            {
                Text = "检查更新",
                AutoSize = true,
                Font = new Font(Font.FontFamily, 12F, FontStyle.Bold),
                ForeColor = Color.FromArgb(31, 41, 55),
                Margin = new Padding(0, 0, 0, 6)
            };
            versionSummaryLabel = new Label
            {
                Text = "当前版本 " + PluginUpdateService.CurrentVersionText,
                AutoSize = true,
                ForeColor = Color.FromArgb(75, 82, 95),
                Margin = new Padding(0, 0, 0, 6)
            };
            statusLabel = new Label
            {
                Text = "点击“检查更新”获取最新版本信息。",
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.TopLeft,
                ForeColor = Color.FromArgb(75, 82, 95),
                Padding = new Padding(0, 2, 0, 0)
            };
            updateProgressBar = new ProgressBar
            {
                Dock = DockStyle.Fill,
                Height = 12,
                Minimum = 0,
                Maximum = 100,
                Value = 0,
                Style = ProgressBarStyle.Continuous,
                Visible = false,
                Margin = new Padding(0, 8, 0, 0)
            };
            statusLayout.Controls.Add(statusHeadingLabel, 0, 0);
            statusLayout.Controls.Add(versionSummaryLabel, 0, 1);
            statusLayout.Controls.Add(statusLabel, 0, 2);
            statusLayout.Controls.Add(updateProgressBar, 0, 3);
            statusPanel.Controls.Add(statusLayout);
            layout.Controls.Add(statusPanel, 0, 4);
            layout.SetColumnSpan(statusPanel, 3);

            FlowLayoutPanel buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            installButton = new Button { Text = "安装更新", AutoSize = true, Enabled = false, MinimumSize = new Size(100, 30) };
            installButton.Click += (_, __) => InstallUpdate();
            checkButton = new Button { Text = "检查更新", AutoSize = true, MinimumSize = new Size(100, 30) };
            checkButton.Click += (_, __) => CheckForUpdates();
            cancelButton = new Button { Text = "关闭", AutoSize = true, MinimumSize = new Size(76, 30), DialogResult = DialogResult.Cancel };
            buttons.Controls.Add(installButton);
            buttons.Controls.Add(checkButton);
            buttons.Controls.Add(cancelButton);
            layout.Controls.Add(buttons, 0, 6);
            layout.SetColumnSpan(buttons, 3);
            Controls.Add(layout);
            AcceptButton = checkButton;
            CancelButton = cancelButton;

            installButton.Tag = "install";
            Tag = installButton;
            Shown += (_, __) =>
            {
                SaveSettings();
            };
        }

        internal bool FoundNewVersion => latestManifest != null && latestManifest.ParsedVersion > PluginUpdateService.CurrentVersion;

        private void BrowseLocalFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog { SelectedPath = localFolderBox.Text })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    localFolderBox.Text = dialog.SelectedPath;
                    SaveSettings();
                }
            }
        }

        private void CheckForUpdates()
        {
            SaveSettings();
            latestManifest = null;
            latestSource = null;
            updateProgressBar.Visible = false;
            statusHeadingLabel.Text = "正在检查更新";
            versionSummaryLabel.Text = "当前版本 " + PluginUpdateService.CurrentVersionText;
            statusLabel.Text = string.IsNullOrWhiteSpace(localFolderBox.Text)
                ? "正在连接 GitHub，请稍候..."
                : "正在读取指定文件夹，请稍候...";
            installButton.Enabled = false;
            Application.DoEvents();

            string localError = string.Empty;
            PluginUpdateManifest selected = null;
            if (string.IsNullOrWhiteSpace(localFolderBox.Text))
            {
                selected = PluginUpdateService.LoadFromGitHub(
                    PluginUpdateService.DefaultGitHubManifestUrl,
                    out localError);
                latestSource = "GitHub";
            }
            else
            {
                selected = PluginUpdateService.LoadFromFolder(localFolderBox.Text.Trim(), out localError);
                latestSource = "指定文件夹";
            }

            if (selected == null)
            {
                statusHeadingLabel.Text = "检查更新失败";
                versionSummaryLabel.Text = "当前版本 " + PluginUpdateService.CurrentVersionText;
                statusLabel.Text = "未找到可用更新源。" + (string.IsNullOrWhiteSpace(localError) ? string.Empty : "\r\n" + localError);
                return;
            }

            latestManifest = selected;
            if (latestManifest == null || latestManifest.ParsedVersion <= PluginUpdateService.CurrentVersion)
            {
                statusHeadingLabel.Text = "无需更新";
                versionSummaryLabel.Text = "当前版本 " + PluginUpdateService.CurrentVersionText;
                statusLabel.Text = "当前已是最新版本。";
                installButton.Enabled = false;
                return;
            }

            statusHeadingLabel.Text = "发现新版本";
            versionSummaryLabel.Text = "当前版本 " + PluginUpdateService.CurrentVersionText
                + "    最新版本 " + latestManifest.Version;
            statusLabel.Text = "来源：" + latestSource + "\r\n" + (latestManifest.Notes ?? "暂无更新说明。");
            installButton.Enabled = true;
        }

        private void InstallUpdate()
        {
            if (!FoundNewVersion || updateInProgress)
            {
                return;
            }

            updateInProgress = true;
            installButton.Enabled = false;
            checkButton.Enabled = false;
            cancelButton.Enabled = false;
            statusHeadingLabel.Text = "正在下载更新";
            versionSummaryLabel.Text = "最新版本 " + latestManifest.Version;
            statusLabel.Text = "正在下载安装包，请稍候... Word 仍可正常使用。";
            updateProgressBar.Style = ProgressBarStyle.Continuous;
            updateProgressBar.Value = 0;
            updateProgressBar.Visible = true;

            Task.Run(() =>
            {
                string error = string.Empty;
                string packagePath = latestSource == "指定文件夹"
                    ? PluginUpdateService.ResolvePackagePath(latestManifest, localFolderBox.Text.Trim())
                    : PluginUpdateService.DownloadPackage(
                        latestManifest,
                        Path.Combine(Path.GetTempPath(), "DocuLint-updates"),
                        ReportDownloadProgress,
                        out error);
                return Tuple.Create(packagePath, error);
            }).ContinueWith(task =>
            {
                if (IsDisposed || !IsHandleCreated)
                {
                    return;
                }

                BeginInvoke(new Action(() => CompleteInstall(task)));
            }, TaskScheduler.Default);
        }

        private void CompleteInstall(Task<Tuple<string, string>> downloadTask)
        {
            updateInProgress = false;
            updateProgressBar.Visible = false;
            cancelButton.Enabled = true;
            checkButton.Enabled = true;

            if (downloadTask.IsFaulted)
            {
                installButton.Enabled = true;
                statusHeadingLabel.Text = "下载更新失败";
                statusLabel.Text = downloadTask.Exception == null
                    ? "下载安装包失败。"
                    : "下载安装包失败：" + downloadTask.Exception.GetBaseException().Message;
                return;
            }

            string packagePath = downloadTask.Result.Item1;
            string error = downloadTask.Result.Item2;
            if (string.IsNullOrWhiteSpace(packagePath))
            {
                statusHeadingLabel.Text = "安装更新失败";
                statusLabel.Text = "无法取得安装包：" + (error ?? "未找到文件");
                installButton.Enabled = true;
                return;
            }

            SaveSettings();
            try
            {
                Process.Start(new ProcessStartInfo { FileName = packagePath, UseShellExecute = true });
                statusHeadingLabel.Text = "安装程序已启动";
                statusLabel.Text = "安装完成后请重新打开 Word。";
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                statusHeadingLabel.Text = "安装更新失败";
                statusLabel.Text = "启动安装程序失败：" + ex.Message;
            }
        }

        private void ReportDownloadProgress(long receivedBytes, long totalBytes)
        {
            if (IsDisposed || !IsHandleCreated)
            {
                return;
            }

            try
            {
                BeginInvoke(new Action(() =>
                {
                    if (IsDisposed)
                    {
                        return;
                    }

                    if (totalBytes <= 0)
                    {
                        updateProgressBar.Style = ProgressBarStyle.Marquee;
                        statusLabel.Text = "正在下载安装包，请稍候...";
                        return;
                    }

                    updateProgressBar.Style = ProgressBarStyle.Continuous;
                    int percent = (int)Math.Min(100L, receivedBytes * 100L / totalBytes);
                    updateProgressBar.Value = Math.Max(0, percent);
                    statusLabel.Text = "正在下载安装包... " + percent + "%";
                }));
            }
            catch (InvalidOperationException)
            {
                // The form may close while the background download is finishing.
            }
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.UpdateLocalFolder = localFolderBox.Text.Trim();
            Properties.Settings.Default.Save();
        }

        private static string GetLocalFolder() => Properties.Settings.Default.UpdateLocalFolder ?? string.Empty;
    }
}
