using System;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn?.Application?.CommandBars?.ExecuteMso("Numbering");
                TryUpdateStatusBar(Globals.ThisAddIn?.Application, "已应用 Word 原生编号");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入编号失败: {ex.Message}", "文档不加班 快速工具");
            }
        }
    }
}
