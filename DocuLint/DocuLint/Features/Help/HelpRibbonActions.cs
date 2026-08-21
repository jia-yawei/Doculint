using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private const string HelpDocumentFileName = "插件帮助文档.html";
        private const string VersionHistoryFileName = "版本更新记录.html";
        private bool helpAboutItemsInitialized;

        private void btnOpenHelpDocument_Click(object sender, RibbonControlEventArgs e)
        {
            OpenLocalHtml(GetHelpDocumentPath(), "帮助文档");
        }

        private void UpdateHelpVersionLabel()
        {
            EnsureHelpAboutItems();
            if (btnHelpVersion != null)
            {
                btnHelpVersion.Label = "版本号：" + GetPluginVersionText();
                btnHelpVersion.Enabled = true;
            }
        }

        private void EnsureHelpAboutItems()
        {
            if (helpAboutItemsInitialized)
            {
                return;
            }

            if (btnHelpVersion != null)
            {
                btnHelpVersion.Enabled = true;
                btnHelpVersion.Click += btnHelpVersion_Click;
            }
            if (btnCheckUpdates != null)
            {
                btnCheckUpdates.Click += btnCheckUpdates_Click;
            }
            helpAboutItemsInitialized = true;
        }

        private void btnCheckUpdates_Click(object sender, RibbonControlEventArgs e)
        {
            using (PluginUpdateForm form = new PluginUpdateForm())
            {
                form.ShowDialog();
            }
        }

        private void btnHelpVersion_Click(object sender, RibbonControlEventArgs e)
        {
            OpenLocalHtml(GetVersionHistoryPath(), "版本更新记录");
        }

        private static string GetPluginVersionText()
        {
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            return version == null ? "未知" : version.ToString();
        }

        private static string GetHelpDocumentPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Help", HelpDocumentFileName);
        }

        private static string GetVersionHistoryPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Help", VersionHistoryFileName);
        }

        private static void OpenLocalHtml(string path, string title)
        {
            if (!File.Exists(path))
            {
                MessageBox.Show(
                    $"未找到{title}：\r\n{path}",
                    title,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = path,
                UseShellExecute = true
            });
        }
    }
}
