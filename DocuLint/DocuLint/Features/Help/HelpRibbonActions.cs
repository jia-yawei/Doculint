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
        private bool helpAboutItemsInitialized;

        private void btnOpenHelpDocument_Click(object sender, RibbonControlEventArgs e)
        {
            string helpPath = GetHelpDocumentPath();
            if (!File.Exists(helpPath))
            {
                MessageBox.Show(
                    $"未找到帮助文档：\r\n{helpPath}",
                    "帮助文档",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = helpPath,
                UseShellExecute = true
            });
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
            helpAboutItemsInitialized = true;
        }

        private void btnHelpVersion_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show(
                BuildVersionUpdateText(),
                "更新内容",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static string BuildVersionUpdateText()
        {
            return "当前版本：" + GetPluginVersionText() +
                   "\r\n作者：软件三室" +
                   "\r\n\r\n更新内容：" +
                   "\r\n- 增加了软件专用的功能。" +
                   "\r\n- 更新了样式管理中的逻辑。" +
                   "\r\n- 修复了部分已知问题。" +
                   "\r\n- 提高了插件执行效率。";
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
    }
}
