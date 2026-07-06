using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal static class GroupCapturedContentService
    {
        private static readonly StringComparer PathComparer = StringComparer.OrdinalIgnoreCase;
        private static GroupCapturedContentInjectForm injectForm;

        public static void CaptureCurrentSelection(Word.Application app, IWin32Window owner)
        {
            try
            {
                if (!TryResolveContext(app, out Word.Document doc, out DocumentGroupStore store, out DocumentGroupCatalog catalog, out DocumentGroupItem group, out string error))
                {
                    MessageBox.Show(owner, error, "内容抓取");
                    return;
                }

                Word.Selection selection = app.Selection;
                Word.Range selectedRange = selection?.Range;
                if (selectedRange == null || selectedRange.Start == selectedRange.End)
                {
                    MessageBox.Show(owner, "请先用鼠标选中要抓取的内容。", "内容抓取");
                    return;
                }

                using (TextPromptForm prompt = new TextPromptForm("内容抓取", "请输入抓取标题："))
                {
                    if (prompt.ShowDialog(owner) != DialogResult.OK)
                    {
                        return;
                    }

                    string title = (prompt.ResultText ?? string.Empty).Trim();
                    title = SanitizeForXml(title);
                    if (string.IsNullOrWhiteSpace(title))
                    {
                        MessageBox.Show(owner, "抓取标题不能为空。", "内容抓取");
                        return;
                    }

                    string openXml = selectedRange.WordOpenXML ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(openXml))
                    {
                        MessageBox.Show(owner, "抓取失败：未读取到选区内容。", "内容抓取");
                        return;
                    }

                    group.CapturedContents = group.CapturedContents ?? new List<DocumentGroupCapturedContentItem>();
                    group.CapturedContents.Insert(0, new DocumentGroupCapturedContentItem
                    {
                        Title = title,
                        PreviewText = BuildPreviewText(selectedRange.Text),
                        ContentWordOpenXml = openXml,
                        CreatedAt = DateTime.Now,
                        UpdatedAt = DateTime.Now
                    });

                    store.Save(catalog);
                    MessageBox.Show(owner, $"已抓取到文档组“{group.Name}”。", "内容抓取");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, $"内容抓取失败：{ex.Message}", "内容抓取");
            }
        }

        public static void OpenCaptureManager(Word.Application app, IWin32Window owner)
        {
            try
            {
                if (!TryResolveContext(app, out _, out DocumentGroupStore store, out DocumentGroupCatalog catalog, out DocumentGroupItem group, out string error))
                {
                    MessageBox.Show(owner, error, "抓取管理");
                    return;
                }

                List<DocumentGroupCapturedContentItem> sourceItems = group.CapturedContents ?? new List<DocumentGroupCapturedContentItem>();
                using (GroupCapturedContentManagerForm form = new GroupCapturedContentManagerForm(group.Name, sourceItems))
                {
                    if (form.ShowDialog(owner) != DialogResult.OK)
                    {
                        return;
                    }

                    group.CapturedContents = form.BuildItems();
                    store.Save(catalog);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, $"打开抓取管理失败：{ex.Message}", "抓取管理");
            }
        }

        public static void InjectCapturedContent(Word.Application app, IWin32Window owner)
        {
            try
            {
                if (!TryResolveContext(app, out _, out DocumentGroupStore _, out DocumentGroupCatalog _, out DocumentGroupItem group, out string error))
                {
                    MessageBox.Show(owner, error, "内容注入");
                    return;
                }

                List<DocumentGroupCapturedContentItem> items = group.CapturedContents ?? new List<DocumentGroupCapturedContentItem>();
                if (items.Count == 0)
                {
                    MessageBox.Show(owner, "当前文档组还没有抓取内容。", "内容注入");
                    return;
                }

                if (injectForm != null && !injectForm.IsDisposed)
                {
                    injectForm.Close();
                }

                injectForm = new GroupCapturedContentInjectForm(group.Name, items);
                injectForm.InjectRequested += selected => InjectToCurrentCursor(app, owner, group, selected);
                injectForm.FormClosed += (_, __) =>
                {
                    injectForm = null;
                };
                injectForm.Show(owner);
                injectForm.BringToFront();
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, $"内容注入失败：{ex.Message}", "内容注入");
            }
        }

        private static void InjectToCurrentCursor(
            Word.Application app,
            IWin32Window owner,
            DocumentGroupItem sourceGroup,
            DocumentGroupCapturedContentItem selected)
        {
            try
            {
                if (app == null)
                {
                    MessageBox.Show(owner, "当前没有活动文档。", "内容注入");
                    return;
                }

                if (selected == null || string.IsNullOrWhiteSpace(selected.ContentWordOpenXml))
                {
                    MessageBox.Show(owner, "未选择可注入的内容。", "内容注入");
                    return;
                }

                Word.Document activeDocument = app.ActiveDocument;
                string activePath = TryNormalizePath(activeDocument?.FullName);
                if (string.IsNullOrWhiteSpace(activePath) || !GroupContainsDocument(sourceGroup, activePath))
                {
                    MessageBox.Show(owner, "请先切换到该文档组内的文档，再执行注入。", "内容注入");
                    return;
                }

                Word.Selection selection = app.Selection;
                Word.Range insertionRange = selection?.Range;
                if (insertionRange == null)
                {
                    MessageBox.Show(owner, "当前没有可用光标位置。", "内容注入");
                    return;
                }

                insertionRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                object transform = Type.Missing;
                insertionRange.InsertXML(selected.ContentWordOpenXml, ref transform);
            }
            catch (Exception ex)
            {
                MessageBox.Show(owner, $"内容注入失败：{ex.Message}", "内容注入");
            }
        }

        private static bool TryResolveContext(
            Word.Application app,
            out Word.Document activeDocument,
            out DocumentGroupStore store,
            out DocumentGroupCatalog catalog,
            out DocumentGroupItem group,
            out string error)
        {
            activeDocument = app?.ActiveDocument;
            store = new DocumentGroupStore();
            catalog = store.Load();
            group = null;
            error = string.Empty;

            if (app == null || activeDocument == null)
            {
                error = "当前没有活动文档。";
                return false;
            }

            string fullPath = TryNormalizePath(activeDocument.FullName);
            if (string.IsNullOrWhiteSpace(fullPath))
            {
                error = "当前文档还没有保存，请先保存后再使用该功能。";
                return false;
            }

            group = ResolveGroupByDocumentPath(catalog, fullPath);
            if (group == null)
            {
                error = "当前文档不在任何文档组中，无法访问抓取内容。";
                return false;
            }

            group.CapturedContents = group.CapturedContents ?? new List<DocumentGroupCapturedContentItem>();
            return true;
        }

        private static DocumentGroupItem ResolveGroupByDocumentPath(DocumentGroupCatalog catalog, string fullPath)
        {
            if (catalog == null || string.IsNullOrWhiteSpace(fullPath))
            {
                return null;
            }

            DocumentGroupItem activeGroup = catalog.GetActiveGroup();
            if (GroupContainsDocument(activeGroup, fullPath))
            {
                return activeGroup;
            }

            return (catalog.Groups ?? new List<DocumentGroupItem>())
                .FirstOrDefault(group => GroupContainsDocument(group, fullPath));
        }

        private static bool GroupContainsDocument(DocumentGroupItem group, string fullPath)
        {
            if (group == null || string.IsNullOrWhiteSpace(fullPath))
            {
                return false;
            }

            return (group.Documents ?? new List<DocumentGroupDocumentItem>())
                .Any(item => PathComparer.Equals(TryNormalizePath(item?.FilePath), fullPath));
        }

        private static string BuildPreviewText(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return string.Empty;
            }

            string text = rawText
                .Replace("\r\a", "\r\n")
                .Replace("\v", "\r\n")
                .Replace("\a", "\t")
                .Replace("\r", "\r\n")
                .Replace("\n", " ")
                .Replace("\t", " ")
                .Trim();
            text = SanitizeForXml(text);

            while (text.Contains("  "))
            {
                text = text.Replace("  ", " ");
            }

            const int maxLength = 140;
            return text.Length <= maxLength ? text : text.Substring(0, maxLength) + "...";
        }

        private static string TryNormalizePath(string path)
        {
            try
            {
                return string.IsNullOrWhiteSpace(path) ? string.Empty : Path.GetFullPath(path.Trim());
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SanitizeForXml(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            StringBuilder sb = new StringBuilder(text.Length);
            foreach (char ch in text)
            {
                if (XmlConvert.IsXmlChar(ch))
                {
                    sb.Append(ch);
                }
            }

            return sb.ToString();
        }
    }
}
