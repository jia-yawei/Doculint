using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class PluginUpdateForm : Form
    {
        private readonly TextBox localFolderBox;
        private readonly CheckBox autoCheckBox;
        private readonly CheckBox skipVersionBox;
        private readonly Label statusLabel;
        private PluginUpdateManifest latestManifest;
        private string latestSource;

        internal PluginUpdateForm(bool automaticCheck = false)
        {
            Text = "检查更新";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Microsoft YaHei UI", 9F);
            ClientSize = new Size(640, 360);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 8,
                Padding = new Padding(16)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
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

            layout.Controls.Add(new Label { Text = "内网更新文件夹", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
            localFolderBox = new TextBox { Dock = DockStyle.Fill, Text = GetLocalFolder() };
            layout.Controls.Add(localFolderBox, 1, 3);
            Button browseButton = new Button { Text = "选择...", AutoSize = true, Dock = DockStyle.Fill };
            browseButton.Click += (_, __) => BrowseLocalFolder();
            layout.Controls.Add(browseButton, 2, 3);

            autoCheckBox = new CheckBox { Text = "启动时自动检查 GitHub 更新", AutoSize = true, Checked = GetAutoCheck(), Margin = new Padding(0, 8, 0, 4) };
            layout.Controls.Add(autoCheckBox, 0, 4);
            layout.SetColumnSpan(autoCheckBox, 3);

            skipVersionBox = new CheckBox { Text = "找到新版本时不再提示该版本", AutoSize = true, Visible = false };
            layout.Controls.Add(skipVersionBox, 0, 5);
            layout.SetColumnSpan(skipVersionBox, 3);

            statusLabel = new Label
            {
                Text = automaticCheck ? "正在检查更新..." : "GitHub 更新源已内置，也可以配置内网更新文件夹。",
                AutoSize = false,
                Dock = DockStyle.Fill,
                ForeColor = Color.FromArgb(75, 82, 95),
                Padding = new Padding(0, 8, 0, 8)
            };
            layout.Controls.Add(statusLabel, 0, 6);
            layout.SetColumnSpan(statusLabel, 3);

            FlowLayoutPanel buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, WrapContents = false };
            Button installButton = new Button { Text = "安装更新", AutoSize = true, Enabled = false, MinimumSize = new Size(100, 30) };
            installButton.Click += (_, __) => InstallUpdate();
            Button checkButton = new Button { Text = "检查更新", AutoSize = true, MinimumSize = new Size(100, 30) };
            checkButton.Click += (_, __) => CheckForUpdates();
            Button cancelButton = new Button { Text = "关闭", AutoSize = true, MinimumSize = new Size(76, 30), DialogResult = DialogResult.Cancel };
            buttons.Controls.Add(installButton);
            buttons.Controls.Add(checkButton);
            buttons.Controls.Add(cancelButton);
            layout.Controls.Add(buttons, 0, 7);
            layout.SetColumnSpan(buttons, 3);
            Controls.Add(layout);
            AcceptButton = checkButton;
            CancelButton = cancelButton;

            installButton.Tag = "install";
            Tag = installButton;
            FormClosing += (_, __) =>
            {
                if (skipVersionBox.Checked && latestManifest != null)
                {
                    Properties.Settings.Default.UpdateSkippedVersion = latestManifest.Version ?? string.Empty;
                    Properties.Settings.Default.Save();
                }
            };
            Shown += (_, __) =>
            {
                SaveSettings();
                if (automaticCheck)
                {
                    BeginInvoke(new Action(CheckForUpdates));
                }
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
            statusLabel.Text = "正在检查 GitHub 和内网更新源...";
            Application.DoEvents();

            string githubError = string.Empty;
            string localError = string.Empty;
            PluginUpdateManifest github = null;
            github = PluginUpdateService.LoadFromGitHub(
                PluginUpdateService.DefaultGitHubManifestUrl,
                out githubError);

            PluginUpdateManifest local = null;
            if (!string.IsNullOrWhiteSpace(localFolderBox.Text))
            {
                local = PluginUpdateService.LoadFromFolder(localFolderBox.Text.Trim(), out localError);
            }

            if (github == null && local == null)
            {
                string details = string.IsNullOrWhiteSpace(localFolderBox.Text)
                    ? githubError
                    : (githubError + (string.IsNullOrWhiteSpace(localError) ? string.Empty : "；" + localError));
                statusLabel.Text = "未找到可用更新源。" + (string.IsNullOrWhiteSpace(details) ? string.Empty : "\r\n" + details);
                return;
            }

            latestManifest = SelectNewerManifest(github, local, out latestSource);
            Button installButton = Tag as Button;
            if (latestManifest == null || latestManifest.ParsedVersion <= PluginUpdateService.CurrentVersion)
            {
                statusLabel.Text = "当前已是最新版本。";
                if (installButton != null) installButton.Enabled = false;
                return;
            }

            statusLabel.Text = "发现新版本 " + latestManifest.Version + "（来源：" + latestSource + "）\r\n" + (latestManifest.Notes ?? string.Empty);
            skipVersionBox.Visible = true;
            if (installButton != null) installButton.Enabled = true;
        }

        private void InstallUpdate()
        {
            if (!FoundNewVersion)
            {
                return;
            }

            string error = string.Empty;
            string packagePath = latestSource == "内网"
                ? PluginUpdateService.ResolvePackagePath(latestManifest, localFolderBox.Text.Trim())
                : PluginUpdateService.DownloadPackage(latestManifest, Path.Combine(Path.GetTempPath(), "DocuLint-updates"), out error);
            if (string.IsNullOrWhiteSpace(packagePath))
            {
                statusLabel.Text = "无法取得安装包：" + (error ?? "未找到文件");
                return;
            }

            SaveSettings();
            try
            {
                Process.Start(new ProcessStartInfo { FileName = packagePath, UseShellExecute = true });
                statusLabel.Text = "安装程序已启动。安装完成后请重新打开 Word。";
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                statusLabel.Text = "启动安装程序失败：" + ex.Message;
            }
        }

        private static PluginUpdateManifest SelectNewerManifest(PluginUpdateManifest first, PluginUpdateManifest second, out string source)
        {
            source = null;
            if (first == null) { source = second == null ? null : "内网"; return second; }
            if (second == null) { source = "GitHub"; return first; }
            if (second.ParsedVersion > first.ParsedVersion) { source = "内网"; return second; }
            source = "GitHub";
            return first;
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.UpdateLocalFolder = localFolderBox.Text.Trim();
            Properties.Settings.Default.UpdateAutoCheck = autoCheckBox.Checked;
            Properties.Settings.Default.Save();
        }

        private static string GetLocalFolder() => Properties.Settings.Default.UpdateLocalFolder ?? string.Empty;
        private static bool GetAutoCheck() => Properties.Settings.Default.UpdateAutoCheck;
    }
}
