using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private sealed class DocumentUndoScope : IDisposable
        {
            private readonly object undoRecord;
            private readonly bool customRecordStarted;

            public DocumentUndoScope(Word.Application app, string actionName)
            {
                if (app == null || string.IsNullOrWhiteSpace(actionName))
                {
                    return;
                }

                try
                {
                    undoRecord = TryGetComProperty(app, "UndoRecord");
                    if (undoRecord == null)
                    {
                        return;
                    }

                    undoRecord.GetType().InvokeMember(
                        "StartCustomRecord",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null,
                        undoRecord,
                        new object[] { actionName });
                    customRecordStarted = true;
                }
                catch
                {
                    customRecordStarted = false;
                }
            }

            public void Dispose()
            {
                if (!customRecordStarted || undoRecord == null)
                {
                    return;
                }

                try
                {
                    undoRecord.GetType().InvokeMember(
                        "EndCustomRecord",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null,
                        undoRecord,
                        null);
                }
                catch
                {
                }
            }
        }

        private void btnInsertTotalPages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ExecuteDocumentCommand("插入总页码", InsertTotalPagesField);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入总页码失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void btnCommonPhrases_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn?.ShowCommonPhrasesPane();
        }

        private void btnInsertSequenceNumber_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ExecuteDocumentCommand("插入域编号", InsertSequenceNumberField);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入域编号失败：{ex.Message}", "文档不加班 快速工具");
            }
        }

        private static void InsertSequenceNumberField()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            Word.Selection selection = app?.Selection;
            string selectedNumber = selection?.Range?.Text?.Trim() ?? string.Empty;
            if (doc == null || selection?.Range == null)
            {
                throw new InvalidOperationException("当前没有可替换的选区。");
            }

            if (selectedNumber.Length == 0 || selectedNumber.Any(character => character < '0' || character > '9'))
            {
                throw new InvalidOperationException("请先选中要生成域编号的数字，例如 001。");
            }

            using (SequenceFieldForm form = new SequenceFieldForm(selectedNumber))
            {
                if (form.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                string identifier = NormalizeSequenceIdentifier(form.SequenceIdentifier);
                if (!IsValidSequenceIdentifier(identifier))
                {
                    throw new InvalidOperationException("SEQ 域名称不能为空，且不能包含反斜杠、双引号或换行。");
                }

                if (!selectedNumber.All(character => character >= '0' && character <= '9'))
                {
                    throw new InvalidOperationException("选中的编号不是有效数字。");
                }

                int width = selectedNumber.Length;
                string format = new string('0', Math.Max(1, width));
                // 直接写入标准 SEQ 域代码，确保 Word 显示为：
                // SEQ STC_TC \\# "000"。
                // 首次出现的域名称始终从 1 开始；选中的数字只用于确定显示位数。
                string fieldCode = $" SEQ {identifier} \\# \"{format}\" ";

                Word.Range replacementRange = selection.Range.Duplicate;
                Word.Field field = replacementRange.Fields.Add(
                    replacementRange,
                    Word.WdFieldType.wdFieldEmpty,
                    fieldCode,
                    false);
                if (field == null)
                {
                    throw new InvalidOperationException("SEQ 域插入失败。");
                }

                field.Update();
                UpdateSequenceFields(doc, identifier);

                Word.Range resultRange = field?.Result?.Duplicate ?? replacementRange;
                resultRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                resultRange.Select();
            }
        }

        private static string NormalizeSequenceIdentifier(string value)
        {
            string identifier = (value ?? string.Empty).Trim();
            if (identifier.StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase))
            {
                identifier = identifier.Substring(4).Trim();
            }

            return identifier;
        }

        private static bool IsValidSequenceIdentifier(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return value.IndexOf('\\') < 0 &&
                   value.IndexOf('"') < 0 &&
                   value.All(character => !char.IsControl(character));
        }

        private static void UpdateSequenceFields(Word.Document doc, string identifier)
        {
            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                try
                {
                    foreach (Word.Field field in storyRange.Fields)
                    {
                        string code = field?.Code?.Text ?? string.Empty;
                        if (Regex.IsMatch(
                            code,
                            $@"^\s*SEQ\s+{Regex.Escape(identifier)}(?:\s|$)",
                            RegexOptions.IgnoreCase))
                        {
                            field.Update();
                        }
                    }
                }
                catch
                {
                }
            }
        }

        private static bool HasSequenceField(Word.Document doc, string identifier)
        {
            if (doc == null || string.IsNullOrWhiteSpace(identifier))
            {
                return false;
            }

            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                try
                {
                    foreach (Word.Field field in storyRange.Fields)
                    {
                        string code = field?.Code?.Text ?? string.Empty;
                        if (Regex.IsMatch(
                            code,
                            $@"^\s*SEQ\s+{Regex.Escape(identifier)}(?:\s|$)",
                            RegexOptions.IgnoreCase))
                        {
                            return true;
                        }
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int updatedCount = ExecuteDocumentCommand("更新总页码", UpdateTotalPagesFields);
                if (updatedCount > 0)
                {
                    MessageBox.Show($"已更新 {updatedCount} 个总页码域。", "文档不加班 快速工具");
                }
                else
                {
                    MessageBox.Show("当前文档中没有找到可更新的总页码域。", "文档不加班 快速工具");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新总页码失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void button26_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int refreshedCount = ExecuteDocumentCommand("刷新目录", RefreshTableOfContents);
                if (refreshedCount > 0)
                {
                    MessageBox.Show($"已刷新 {refreshedCount} 个目录。", "文档不加班 快速工具");
                }
                else
                {
                    MessageBox.Show("当前文档中没有找到可刷新的目录。", "文档不加班 快速工具");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刷新目录失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void btnTogglePageWhitespace_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                bool isVisible = TogglePageWhitespace();
                btnTogglePageWhitespace.Checked = isVisible;
            }
            catch (Exception ex)
            {
                btnTogglePageWhitespace.Checked = IsPageWhitespaceVisible();
                MessageBox.Show($"切换页面间空白失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void button32_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int updatedCount = ExecuteDocumentCommand("图片单倍行距", ApplySingleLineSpacingToPictures);
                MessageBox.Show(
                    updatedCount > 0
                        ? $"已将 {updatedCount} 个图片段落设置为单倍行距。"
                        : "当前文档中没有找到可处理的图片。",
                    "文档不加班 快速工具");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"图片单倍行距失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void btnApplyHeitiXiaosi_Click(object sender, RibbonControlEventArgs e)
        {
            ApplyQuickFont("黑体", 12f, "黑体小四");
        }

        private void btnApplySongtiXiaosi_Click(object sender, RibbonControlEventArgs e)
        {
            ApplyQuickFont("宋体", 12f, "宋体小四");
        }

        private void btnClearFormatting_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int updatedCount = ExecuteDocumentCommand("一键清除格式", ClearSelectionToNormalText);
                MessageBox.Show(
                    updatedCount > 0
                        ? "格式清除成功"
                        : "格式清除失败：当前没有可处理的段落。",
                    "文档不加班 快速工具");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"格式清除失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void btnCleanFieldCodes_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int unlinkedCount = ExecuteDocumentCommand("清理域代码", RemoveSelectedFieldCodes);
                MessageBox.Show(
                    unlinkedCount > 0
                        ? $"已清理 {unlinkedCount} 个域代码，保留当前显示结果。"
                        : "当前选区中没有域代码。",
                    "文档不加班 快速工具");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"清理域代码失败：{ex.Message}", "文档不加班 快速工具");
            }
        }

        private static int RemoveSelectedFieldCodes()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app?.Selection;
            Word.Range range = selection?.Range;
            if (range == null || range.Start == range.End)
            {
                throw new InvalidOperationException("请先选中包含域代码的文本。");
            }

            Word.Fields fields = range.Fields;
            int count = fields?.Count ?? 0;
            int unlinkedCount = 0;
            for (int index = count; index >= 1; index--)
            {
                try
                {
                    Word.Field field = fields[index];
                    field?.Unlink();
                    unlinkedCount++;
                }
                catch
                {
                }
            }

            return unlinkedCount;
        }

        private static void ApplyQuickFont(string fontName, float fontSize, string displayName)
        {
            try
            {
                ExecuteDocumentCommand(displayName, () => ApplyFontToSelection(fontName, fontSize));
                TryUpdateStatusBar(Globals.ThisAddIn.Application, $"已设置为{displayName}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{displayName}设置失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private static void ApplyFontToSelection(string fontName, float fontSize)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app?.Selection;
            if (selection?.Range == null)
            {
                throw new InvalidOperationException("当前没有可设置的选区或输入点。");
            }

            Word.Font font = selection.Range.Font;
            font.NameFarEast = fontName;
            font.Name = fontName;
            font.Size = fontSize;

            // 光标未选中文本时，同时设置 Selection.Font，保证后续输入也使用该字体。
            if (selection.Range.Start == selection.Range.End)
            {
                selection.Font.NameFarEast = fontName;
                selection.Font.Name = fontName;
                selection.Font.Size = fontSize;
            }
        }

        private void btnCleanBlankPages_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int deletedCount = ExecuteDocumentCommand("清理空白页", CleanBlankPages);
                MessageBox.Show(
                    deletedCount > 0
                        ? $"已清理 {deletedCount} 个空白页。"
                        : "未发现可清理的空白页。",
                    "文档不加班 快速工具");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"清理空白页失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void btnCleanInvalidStyles_Click(object sender, RibbonControlEventArgs e)
        {
            DialogResult confirmation = MessageBox.Show(
                "只将第一页没有可见内容且不是正文样式的段落设置为正文。\r\n\r\n只修改段落样式，不删除空段落、自动编号或页面对象。是否继续？",
                "清除首页无效样式",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            if (confirmation != DialogResult.OK)
            {
                return;
            }

            try
            {
                int cleanedCount = ExecuteDocumentCommand("清除首页无效样式", CleanInvalidEmptyHeadingStyles);
                MessageBox.Show(
                    cleanedCount > 0
                        ? $"已将首页 {cleanedCount} 个空白标题段落设置为正文样式。"
                        : "首页未发现需要清除的空白标题样式段落。",
                    "文档不加班 清理",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                TryUpdateStatusBar(Globals.ThisAddIn?.Application, "清除首页无效样式已停止");
                MessageBox.Show("清除首页无效样式已停止。", "文档不加班 清理", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"清除首页无效样式失败: {ex.Message}", "文档不加班 清理");
            }
        }

        private void button23_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!TryUndoLastCommand())
                {
                    MessageBox.Show("当前没有可撤销的文档操作。", "文档不加班 快速工具");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"撤销上次命令失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private static void ExecuteDocumentCommand(string actionName, Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            ResetOperationCancellation(actionName);
            Word.Application app = Globals.ThisAddIn.Application;
            using (new DocumentUndoScope(app, actionName))
            {
                action();
            }
        }

        private static T ExecuteDocumentCommand<T>(string actionName, Func<T> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            ResetOperationCancellation(actionName);
            Word.Application app = Globals.ThisAddIn.Application;
            using (new DocumentUndoScope(app, actionName))
            {
                return action();
            }
        }

        private static bool TryUndoLastCommand()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            if (app == null)
            {
                return false;
            }

            try
            {
                object result = app.GetType().InvokeMember(
                    "Undo",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    app,
                    null);

                if (result is bool boolResult)
                {
                    return boolResult;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void InsertTotalPagesField()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app?.Selection;
            Word.Document doc = app?.ActiveDocument;
            if (selection?.Range == null)
            {
                throw new InvalidOperationException("当前没有可插入的位置。");
            }

            if (doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            PreparePaginationForTotalPages(app, doc);

            Word.Range insertRange = selection.Range.Duplicate;
            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Field field = insertRange.Fields.Add(
                insertRange,
                Word.WdFieldType.wdFieldNumPages,
                Type.Missing,
                false);
            PreparePaginationForTotalPages(app, doc);

            try
            {
                field?.Update();
            }
            catch
            {
            }

            try
            {
                Word.Range fieldRange = field?.Result?.Duplicate ?? field?.Code?.Duplicate;
                if (fieldRange != null)
                {
                    UpdateTotalPagesFieldsInRange(fieldRange);
                }
                else
                {
                    doc.Fields.Update();
                }
            }
            catch
            {
            }

            Word.Range endRange = field?.Result?.Duplicate ?? insertRange;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            endRange.Select();
        }

        private static int UpdateTotalPagesFields()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            PreparePaginationForTotalPages(app, doc);

            int updatedCount = 0;
            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                ThrowIfOperationCancelled();
                updatedCount += UpdateTotalPagesFieldsInRange(storyRange);
            }

            try
            {
                app.ScreenRefresh();
            }
            catch
            {
            }

            return updatedCount;
        }

        private static int RefreshTableOfContents()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            int refreshedCount = 0;
            using (new WordPerformanceScope(app))
            {
                try
                {
                    doc.Repaginate();
                }
                catch
                {
                }

                try
                {
                    if (doc.TablesOfContents != null)
                    {
                        foreach (Word.TableOfContents toc in doc.TablesOfContents)
                        {
                            ThrowIfOperationCancelled();
                            if (toc == null)
                            {
                                continue;
                            }

                            toc.Update();
                            refreshedCount++;
                        }
                    }
                }
                catch
                {
                }

                try
                {
                    doc.Fields.Update();
                }
                catch
                {
                }

                ApplyTableOfContentsFont(doc);

                try
                {
                    app.ScreenRefresh();
                }
                catch
                {
                }

                Application.DoEvents();
            }

            return refreshedCount;
        }

        private static void ApplyTableOfContentsFont(Word.Document doc)
        {
            if (doc?.TablesOfContents == null)
            {
                return;
            }

            try
            {
                foreach (Word.TableOfContents toc in doc.TablesOfContents)
                {
                    ThrowIfOperationCancelled();
                    try
                    {
                        Word.Range range = toc?.Range;
                        if (range == null)
                        {
                            continue;
                        }

                        range.Font.NameFarEast = "宋体";
                        range.Font.Name = "宋体";
                        range.Font.Size = 12f;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        private static int ApplySingleLineSpacingToPictures()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            int updatedCount = 0;
            HashSet<int> paragraphStarts = new HashSet<int>();
            using (new WordPerformanceScope(app))
            {
                try
                {
                    foreach (Word.InlineShape inlineShape in doc.InlineShapes)
                    {
                        ThrowIfOperationCancelled();
                        if (!IsPictureInlineShape(inlineShape))
                        {
                            continue;
                        }

                        if (ApplySingleLineSpacing(inlineShape.Range, paragraphStarts))
                        {
                            updatedCount++;
                        }
                    }
                }
                catch
                {
                }

                try
                {
                    foreach (Word.Shape shape in doc.Shapes)
                    {
                        ThrowIfOperationCancelled();
                        if (!IsPictureShape(shape))
                        {
                            continue;
                        }

                        if (ApplySingleLineSpacing(shape.Anchor, paragraphStarts))
                        {
                            updatedCount++;
                        }
                    }
                }
                catch
                {
                }
            }

            return updatedCount;
        }

        private static int ClearSelectionToNormalText()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app?.Selection;
            if (selection?.Range == null)
            {
                throw new InvalidOperationException("当前没有可处理的选区。");
            }

            int updatedCount = 0;
            HashSet<int> starts = new HashSet<int>();
            Word.Paragraphs paragraphs = selection.Range.Paragraphs;
            if (paragraphs == null || paragraphs.Count == 0)
            {
                return 0;
            }

            using (new WordPerformanceScope(app))
            {
                foreach (Word.Paragraph paragraph in paragraphs)
                {
                    ThrowIfOperationCancelled();
                    try
                    {
                        Word.Range range = paragraph.Range;
                        if (range == null || !starts.Add(range.Start))
                        {
                            continue;
                        }

                        object normalStyle = Word.WdBuiltinStyle.wdStyleNormal;
                        range.set_Style(ref normalStyle);
                        range.ListFormat.RemoveNumbers();
                        paragraph.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
                        range.ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevelBodyText;
                        updatedCount++;
                    }
                    catch
                    {
                    }
                }
            }

            return updatedCount;
        }

        internal static List<NavigationPaneEntry> CollectBrokenReferenceEntries(Word.Document doc)
        {
            List<NavigationPaneEntry> entries = new List<NavigationPaneEntry>();
            HashSet<int> seenStarts = new HashSet<int>();
            string[] brokenTexts =
            {
                "错误！未找到引用源",
                "错误!未找到引用源",
                "未找到引用源",
                "Error! Reference source not found."
            };

            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                ThrowIfOperationCancelled();
                AddBrokenReferenceFieldEntries(storyRange, brokenTexts, entries, seenStarts);
                AddBrokenReferenceEntries(storyRange, brokenTexts, entries, seenStarts);
            }

            return entries.OrderBy(item => item.Start).ToList();
        }

        private static void AddBrokenReferenceFieldEntries(
            Word.Range storyRange,
            string[] brokenTexts,
            List<NavigationPaneEntry> entries,
            ISet<int> seenStarts)
        {
            if (storyRange?.Fields == null || brokenTexts == null)
            {
                return;
            }

            try
            {
                foreach (Word.Field field in storyRange.Fields)
                {
                    ThrowIfOperationCancelled();
                    try
                    {
                        Word.Range result = field?.Result;
                        string text = result?.Text ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(text)
                            || !brokenTexts.Any(item => text.IndexOf(item, StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            continue;
                        }

                        int start = result.Start;
                        if (seenStarts.Add(start))
                        {
                            entries.Add(new NavigationPaneEntry
                            {
                                Start = start,
                                Text = "未更新域：" + BuildBrokenReferenceSnippet(text, 0)
                            });
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        private static void AddBrokenReferenceEntries(
            Word.Range storyRange,
            string[] brokenTexts,
            List<NavigationPaneEntry> entries,
            ISet<int> seenStarts)
        {
            if (storyRange == null || brokenTexts == null || brokenTexts.Length == 0)
            {
                return;
            }

            string storyText;
            try
            {
                storyText = storyRange.Text ?? string.Empty;
            }
            catch
            {
                return;
            }

            foreach (string brokenText in brokenTexts)
            {
                ThrowIfOperationCancelled();
                int offset = 0;
                while (offset < storyText.Length)
                {
                    ThrowIfOperationCancelled();
                    int index = storyText.IndexOf(brokenText, offset, StringComparison.OrdinalIgnoreCase);
                    if (index < 0)
                    {
                        break;
                    }

                    int start = storyRange.Start + index;
                    if (seenStarts.Add(start))
                    {
                        entries.Add(new NavigationPaneEntry
                        {
                            Start = start,
                            Text = "未更新域：" + BuildBrokenReferenceSnippet(storyText, index)
                        });
                    }

                    offset = index + Math.Max(1, brokenText.Length);
                }
            }
        }

        private static string BuildBrokenReferenceSnippet(string storyText, int index)
        {
            try
            {
                int start = Math.Max(0, storyText.LastIndexOf('\r', Math.Max(0, index - 1)) + 1);
                int end = storyText.IndexOf('\r', index);
                if (end < 0)
                {
                    end = storyText.Length;
                }

                string text = storyText.Substring(start, Math.Max(0, end - start))
                    .Replace("\a", string.Empty)
                    .Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    text = storyText.Replace("\r", string.Empty).Replace("\a", string.Empty).Trim();
                }

                return text.Length > 60 ? text.Substring(0, 57) + "..." : text;
            }
            catch
            {
                return "错误！未找到引用源";
            }
        }

        private static int CleanBlankPages()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            int deletedCount = 0;
            using (new WordPerformanceScope(app))
            {
                int pageCount = 0;
                try
                {
                    doc.Repaginate();
                    pageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);
                }
                catch
                {
                    pageCount = 0;
                }

                for (int page = pageCount; page >= 1; page--)
                {
                    ThrowIfOperationCancelled();
                    Word.Range pageRange = TryGetPageRange(doc, page, pageCount);
                    if (pageRange == null || !IsBlankPageRange(doc, pageRange))
                    {
                        continue;
                    }

                    try
                    {
                        pageRange.Delete();
                        deletedCount++;
                    }
                    catch
                    {
                    }
                }
            }

            return deletedCount;
        }

        private static int CleanInvalidEmptyHeadingStyles()
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc?.Content == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            int cleanedCount = 0;

            using (new WordPerformanceScope(app))
            {
                int firstPageEnd = GetFirstPageEnd(doc);
                List<Word.Range> styledBlankRanges = CollectBlankStyledParagraphRanges(doc, firstPageEnd);
                foreach (Word.Range paragraphRange in styledBlankRanges)
                {
                    ThrowIfOperationCancelled();
                    try
                    {
                        object normalStyle = Word.WdBuiltinStyle.wdStyleNormal;
                        paragraphRange.set_Style(ref normalStyle);
                        cleanedCount++;
                    }
                    catch
                    {
                    }
                }
            }

            return cleanedCount;
        }

        private static List<Word.Range> CollectBlankStyledParagraphRanges(Word.Document doc, int pageEnd)
        {
            List<Word.Range> ranges = new List<Word.Range>();
            if (doc == null)
            {
                return ranges;
            }

            try
            {
                int boundedEnd = Math.Max(doc.Content.Start, Math.Min(pageEnd, doc.Content.End));
                Word.Range firstPage = doc.Range(doc.Content.Start, boundedEnd);
                Word.Paragraphs paragraphs = firstPage.Paragraphs;
                int paragraphCount = paragraphs?.Count ?? 0;
                string normalStyleName = GetNormalStyleName(doc);

                for (int index = 1; index <= paragraphCount; index++)
                {
                    ThrowIfOperationCancelled();
                    Word.Paragraph paragraph = paragraphs[index];
                    Word.Range paragraphRange = paragraph?.Range;
                    if (paragraphRange == null || paragraphRange.Start >= boundedEnd)
                    {
                        continue;
                    }

                    if (!IsVisuallyEmptyParagraphText(paragraphRange.Text))
                    {
                        continue;
                    }

                    string styleName = GetParagraphStyleName(paragraph);
                    if (!IsNormalStyleName(styleName, normalStyleName)
                        && paragraphRange.Start < boundedEnd)
                    {
                        ranges.Add(paragraphRange.Duplicate);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch
            {
            }

            return ranges;
        }

        private static string GetNormalStyleName(Word.Document doc)
        {
            try
            {
                object key = Word.WdBuiltinStyle.wdStyleNormal;
                return doc?.Styles?[key]?.NameLocal ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool IsNormalStyleName(string styleName, string normalStyleName)
        {
            string value = (styleName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return true;
            }

            return (!string.IsNullOrWhiteSpace(normalStyleName)
                    && string.Equals(value, normalStyleName, StringComparison.CurrentCultureIgnoreCase))
                || string.Equals(value, "正文", StringComparison.CurrentCultureIgnoreCase)
                || string.Equals(value, "Normal", StringComparison.OrdinalIgnoreCase);
        }

        private static int GetFirstPageEnd(Word.Document doc)
        {
            if (doc?.Content == null)
            {
                return 0;
            }

            int pageEnd = doc.Content.End;
            try
            {
                object what = Word.WdGoToItem.wdGoToPage;
                object which = Word.WdGoToDirection.wdGoToAbsolute;
                object count = 2;
                Word.Range secondPage = doc.GoTo(ref what, ref which, ref count);
                if (secondPage != null
                    && secondPage.Start > doc.Content.Start
                    && secondPage.Start < pageEnd)
                {
                    pageEnd = secondPage.Start;
                }
            }
            catch
            {
            }

            return pageEnd;
        }

        private static bool IsVisuallyEmptyParagraphText(string text)
        {
            string visibleText = (text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\t", string.Empty)
                .Replace("\v", string.Empty)
                .Replace(" ", string.Empty)
                .Replace("\u3000", string.Empty)
                .Replace("\u00A0", string.Empty)
                .Replace("\u200B", string.Empty)
                .Replace("\uFEFF", string.Empty);
            return visibleText.Length == 0;
        }

        private static Word.Range TryGetPageRange(Word.Document doc, int page, int pageCount)
        {
            try
            {
                object what = Word.WdGoToItem.wdGoToPage;
                object which = Word.WdGoToDirection.wdGoToAbsolute;
                object count = page;
                Word.Range startRange = doc.GoTo(ref what, ref which, ref count);
                if (startRange == null)
                {
                    return null;
                }

                int end = doc.Content.End;
                if (page < pageCount)
                {
                    count = page + 1;
                    Word.Range nextRange = doc.GoTo(ref what, ref which, ref count);
                    if (nextRange != null)
                    {
                        end = nextRange.Start;
                    }
                }

                return end > startRange.Start
                    ? doc.Range(startRange.Start, end)
                    : null;
            }
            catch
            {
                return null;
            }
        }

        private static bool IsBlankPageRange(Word.Document doc, Word.Range range)
        {
            try
            {
                if (range == null || range.Start <= doc.Content.Start && range.End >= doc.Content.End)
                {
                    return false;
                }

                if ((range.Tables?.Count ?? 0) > 0 || (range.InlineShapes?.Count ?? 0) > 0)
                {
                    return false;
                }

                if (ContainsAnchoredShape(doc, range))
                {
                    return false;
                }

                string text = (range.Text ?? string.Empty)
                    .Replace("\r", string.Empty)
                    .Replace("\a", string.Empty)
                    .Replace("\f", string.Empty)
                    .Replace("\v", string.Empty)
                    .Replace("\t", string.Empty)
                    .Replace(" ", string.Empty)
                    .Replace("\u3000", string.Empty)
                    .Replace("\u00A0", string.Empty);

                return text.Length == 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool ContainsAnchoredShape(Word.Document doc, Word.Range range)
        {
            try
            {
                foreach (Word.Shape shape in doc.Shapes)
                {
                    int start = shape?.Anchor?.Start ?? -1;
                    if (start >= range.Start && start < range.End)
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool ApplySingleLineSpacing(Word.Range range, ISet<int> paragraphStarts)
        {
            if (range?.Paragraphs == null)
            {
                return false;
            }

            try
            {
                Word.Paragraph paragraph = range.Paragraphs.Count > 0 ? range.Paragraphs[1] : null;
                if (paragraph?.Range == null)
                {
                    return false;
                }

                if (paragraphStarts != null && !paragraphStarts.Add(paragraph.Range.Start))
                {
                    return false;
                }

                paragraph.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void PreparePaginationForTotalPages(Word.Application app, Word.Document doc)
        {
            if (doc == null)
            {
                return;
            }

            try
            {
                doc.Repaginate();
            }
            catch
            {
            }

            try
            {
                doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);
            }
            catch
            {
            }

            try
            {
                app?.ScreenRefresh();
            }
            catch
            {
            }

            Application.DoEvents();
        }

        private static IEnumerable<Word.Range> EnumerateStoryRanges(Word.Document doc)
        {
            if (doc == null)
            {
                yield break;
            }

            Word.WdStoryType[] storyTypes =
            {
                Word.WdStoryType.wdMainTextStory,
                Word.WdStoryType.wdTextFrameStory
            };

            foreach (Word.WdStoryType storyType in storyTypes)
            {
                Word.Range storyRange = null;
                try
                {
                    storyRange = doc.StoryRanges?[storyType];
                }
                catch
                {
                }

                while (storyRange != null)
                {
                    yield return storyRange;

                    Word.Range nextRange = null;
                    try
                    {
                        nextRange = storyRange.NextStoryRange;
                    }
                    catch
                    {
                    }

                    storyRange = nextRange;
                }
            }
        }

        private static int UpdateTotalPagesFieldsInRange(Word.Range range)
        {
            if (range?.Fields == null)
            {
                return 0;
            }

            int updatedCount = 0;
            try
            {
                foreach (Word.Field field in range.Fields)
                {
                    ThrowIfOperationCancelled();
                    if (field == null)
                    {
                        continue;
                    }

                    if (field.Type != Word.WdFieldType.wdFieldNumPages)
                    {
                        continue;
                    }

                    field.Update();
                    updatedCount++;
                }
            }
            catch
            {
            }

            return updatedCount;
        }

        private static bool TogglePageWhitespace()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            if (app?.ActiveWindow == null)
            {
                throw new InvalidOperationException("当前没有活动文档窗口。");
            }

            Word.View view = app.ActiveWindow.View;
            if (view == null)
            {
                throw new InvalidOperationException("未能获取当前页面视图。");
            }

            try
            {
                view.Type = Word.WdViewType.wdPrintView;
            }
            catch
            {
            }

            bool current = IsPageWhitespaceVisible(view, app);
            bool next = !current;
            SetPageWhitespaceVisible(view, app, next);
            return IsPageWhitespaceVisible(view, app);
        }

        private static bool IsPageWhitespaceVisible()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.View view = app?.ActiveWindow?.View;
            return IsPageWhitespaceVisible(view, app);
        }

        private static bool IsPageWhitespaceVisible(object view, Word.Application app)
        {
            object viewValue = TryGetComProperty(view, "DisplayPageBoundaries");
            if (viewValue != null)
            {
                return Convert.ToBoolean(viewValue);
            }

            object options = TryGetComProperty(app, "Options");
            object optionsValue = TryGetComProperty(options, "DisplayPageBoundaries");
            return optionsValue != null && Convert.ToBoolean(optionsValue);
        }

        private static void SetPageWhitespaceVisible(object view, Word.Application app, bool visible)
        {
            TrySetComProperty(view, "DisplayPageBoundaries", visible);

            object options = TryGetComProperty(app, "Options");
            TrySetComProperty(options, "DisplayPageBoundaries", visible);
        }
    }
}
