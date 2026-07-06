using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private readonly DocumentGroupStore documentGroupStore = new DocumentGroupStore();

        private void InitializeDocumentGroupMenu()
        {
            RefreshDocumentGroupMenu();
            button4.Click += button4_Click;
        }

        internal void RefreshDocumentGroupMenu()
        {
            DocumentGroupCatalog catalog = documentGroupStore.Load();
            DocumentGroupItem activeGroup = documentGroupStore.EnsureActiveGroup(catalog, TryGetCurrentDocumentPathOrEmpty());
            int groupCount = catalog.Groups?.Count ?? 0;
            button27.Label = activeGroup != null
                ? $"当前活动组：{activeGroup.Name}"
                : "当前活动组：未设置";
            button4.ScreenTip = activeGroup != null
                ? $"当前活动文档组：{activeGroup.Name}"
                : groupCount > 0
                    ? $"当前已有 {groupCount} 个文档组，但尚未设置活动组"
                    : "当前还没有文档组";
            button4.SuperTip = activeGroup != null
                ? $"点击后会将当前文档直接加入活动文档组：{activeGroup.Name}。"
                : groupCount > 0
                    ? "点击后可选择一个已有文档组，将当前文档加入其中。"
                : "点击后会先创建一个文档组，再把当前文档加入进去。";
            button1.ScreenTip = "打开文档组管理窗口";
            button1.SuperTip = activeGroup != null
                ? $"当前活动文档组为“{activeGroup.Name}”，可在管理窗口中切换和维护。"
                : groupCount > 0
                    ? $"当前共 {groupCount} 个文档组，可在管理窗口中维护组和组内文档。"
                : "打开管理窗口后可以创建文档组，并维护组内文档。";
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                DocumentGroupCatalog catalog = documentGroupStore.Load();
                List<DocumentGroupItem> groups = catalog.GetOrderedGroups().ToList();
                string currentDocumentPath = GetCurrentDocumentPath();
                DocumentGroupItem activeGroup = documentGroupStore.EnsureActiveGroup(catalog, currentDocumentPath);

                if (groups.Count == 0)
                {
                    using (TextPromptForm prompt = new TextPromptForm("新建文档组", "还没有文档组，请先输入一个文档组名称："))
                    {
                        if (prompt.ShowDialog() != DialogResult.OK)
                        {
                            return;
                        }

                        DocumentGroupItem created = documentGroupStore.CreateGroup(catalog, prompt.ResultText);
                        documentGroupStore.AddDocumentToGroup(catalog, created.Id, currentDocumentPath);
                    }

                    RefreshDocumentGroupMenu();
                    MessageBox.Show("当前文档已加入新建文档组。", "文档组");
                    return;
                }

                if (activeGroup != null)
                {
                    AddCurrentDocumentToGroup(activeGroup.Id);
                    return;
                }

                using (DocumentGroupPickerForm picker = new DocumentGroupPickerForm(
                    groups,
                    System.IO.Path.GetFileName(currentDocumentPath),
                    catalog.ActiveGroupId))
                {
                    if (picker.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(picker.SelectedGroupId))
                    {
                        return;
                    }

                    AddCurrentDocumentToGroup(picker.SelectedGroupId);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "文档组");
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            string currentDocumentPath = TryGetCurrentDocumentPathOrEmpty();
            DocumentGroupCatalog catalog = documentGroupStore.Load();
            documentGroupStore.RefreshDocumentMetadata(catalog);
            documentGroupStore.EnsureActiveGroup(catalog, currentDocumentPath);

            using (DocumentGroupManagerForm form = new DocumentGroupManagerForm(documentGroupStore, catalog, currentDocumentPath))
            {
                form.ShowDialog();
                if (form.DataChanged)
                {
                    RefreshDocumentGroupMenu();
                }
            }
        }

        private void AddCurrentDocumentToGroup(string groupId)
        {
            string currentDocumentPath = GetCurrentDocumentPath();
            DocumentGroupCatalog catalog = documentGroupStore.Load();
            documentGroupStore.AddDocumentToGroup(catalog, groupId, currentDocumentPath);
            documentGroupStore.RefreshDocumentMetadata(catalog);
            RefreshDocumentGroupMenu();

            DocumentGroupItem group = catalog.Groups.FirstOrDefault(item =>
                string.Equals(item.Id, groupId, StringComparison.OrdinalIgnoreCase));
            string groupName = group?.Name ?? "目标组";
            MessageBox.Show($"当前文档已加入文档组：{groupName}", "文档组");
        }

        private static string GetCurrentDocumentPath()
        {
            string path = TryGetCurrentDocumentPathOrEmpty();
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new InvalidOperationException("当前文档还没有保存，请先保存后再加入文档组。");
            }

            return path;
        }

        private static string TryGetCurrentDocumentPathOrEmpty()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            if (app == null)
            {
                return string.Empty;
            }

            try
            {
                if (app.Documents == null || app.Documents.Count < 1)
                {
                    return string.Empty;
                }

                Word.Document doc = app.ActiveDocument;
                if (doc == null || string.IsNullOrWhiteSpace(doc.Path))
                {
                    return string.Empty;
                }

                return doc.FullName ?? string.Empty;
            }
            catch (COMException)
            {
                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
