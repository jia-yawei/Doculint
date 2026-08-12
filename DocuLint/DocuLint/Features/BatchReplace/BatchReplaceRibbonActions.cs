using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private void btnBatchReplace_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            BatchReplaceExecutionRequest request;

            using (BatchReplaceDialog dialog = new BatchReplaceDialog())
            {
                if (dialog.ShowDialog() != DialogResult.OK || dialog.Request == null)
                {
                    return;
                }

                request = dialog.Request;
            }

            List<string> files = request.FilePaths ?? new List<string>();
            if (files.Count == 0)
            {
                MessageBox.Show("未选择可处理文件。", "文档不加班 批量替换");
                return;
            }

            int processedFiles = 0;
            int changedFiles = 0;
            int totalReplacements = 0;
            int totalFileNameHits = 0;
            int renamedFiles = 0;
            List<string> failedFiles = new List<string>();
            List<string> failedDetails = new List<string>();

            bool oldScreenUpdating = app.ScreenUpdating;
            Word.WdAlertLevel oldAlerts = app.DisplayAlerts;

            app.ScreenUpdating = false;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

            try
            {
                foreach (string filePath in files)
                {
                    Word.Document doc = null;
                    bool openedByUs = false;
                    string currentPath = filePath;

                    try
                    {
                        bool hasContentRules = request.Rules.Any(r =>
                            r.Enabled &&
                            !string.IsNullOrWhiteSpace(r.FindText) &&
                            r.Scope != BatchFindScope.None);

                        int replacedCount = 0;
                        int renamedCount = 0;
                        bool fileRenamed = false;

                        if (hasContentRules)
                        {
                            doc = FindOpenedDocumentByPath(app, currentPath);
                            if (doc == null)
                            {
                                doc = app.Documents.Open(currentPath, ReadOnly: false, Visible: false);
                                openedByUs = true;
                            }

                            replacedCount = ReplaceInDocument(doc, request);
                            if (!request.FindOnly && replacedCount > 0)
                            {
                                doc.Save();
                            }
                        }

                        if (!request.FindOnly)
                        {
                            string renamedPath;
                            renamedCount = BuildRenamedPath(currentPath, request, out renamedPath);

                            if (renamedCount > 0 &&
                                !string.Equals(currentPath, renamedPath, StringComparison.OrdinalIgnoreCase))
                            {
                                if (doc != null)
                                {
                                    if (!openedByUs)
                                    {
                                        throw new IOException("文档正在 Word 中打开，无法重命名: " + currentPath);
                                    }

                                    try
                                    {
                                        doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                                    }
                                    catch
                                    {
                                    }
                                    finally
                                    {
                                        ReleaseComObject(doc);
                                        doc = null;
                                    }
                                }

                                if (File.Exists(renamedPath))
                                {
                                    throw new IOException("目标文件名已存在: " + renamedPath);
                                }

                                File.Move(currentPath, renamedPath);
                                fileRenamed = true;
                                currentPath = renamedPath;
                            }

                            if (openedByUs && doc != null)
                            {
                                try
                                {
                                    doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                                }
                                catch
                                {
                                }
                                finally
                                {
                                    ReleaseComObject(doc);
                                    doc = null;
                                }
                            }
                        }

                        if (replacedCount > 0)
                        {
                            changedFiles++;
                            totalReplacements += replacedCount;
                        }
                        else if (fileRenamed)
                        {
                            changedFiles++;
                        }

                        totalReplacements += renamedCount;
                        totalFileNameHits += renamedCount;
                        if (fileRenamed)
                        {
                            renamedFiles++;
                        }

                        processedFiles++;
                    }
                    catch (Exception ex)
                    {
                        failedFiles.Add(currentPath);
                        failedDetails.Add($"{currentPath} | {ex.Message}");
                    }
                    finally
                    {
                        if (doc != null && openedByUs)
                        {
                            try
                            {
                                doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                            }
                            catch
                            {
                            }
                        }

                        ReleaseComObject(doc);
                        doc = null;
                    }

                    // Word can retain a small number of interop wrappers until a GC pass.
                    // Reclaim them periodically so long batches do not grow the Word process.
                    if ((processedFiles + failedFiles.Count) % 5 == 0)
                    {
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
            finally
            {
                app.DisplayAlerts = oldAlerts;
                app.ScreenUpdating = oldScreenUpdating;
            }

            string summary =
                (request.FindOnly ? "批量查找完成。\n\n" : "批量替换完成。\n\n") +
                $"扫描文件：{files.Count} 个\n" +
                $"成功处理：{processedFiles} 个\n" +
                $"发生修改：{changedFiles} 个\n" +
                $"总命中次数：{totalReplacements} 次\n" +
                $"文档名命中：{totalFileNameHits} 次\n" +
                $"成功改名：{renamedFiles} 个\n" +
                $"失败文件：{failedFiles.Count} 个";

            if (failedFiles.Count > 0)
            {
                string failedTop = string.Join("\n", failedFiles.Take(5).ToArray());
                summary += "\n\n失败示例（最多显示 5 个）：\n" + failedTop;

                string reasonTop = string.Join("\n", failedDetails.Take(5).ToArray());
                summary += "\n\n失败原因（最多显示 5 个）：\n" + reasonTop;
            }

            MessageBox.Show(summary, "文档不加班 批量替换");
        }

        private void btnStyleBrush_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn?.Application?.CommandBars?.ExecuteMso("FormatPainter");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动格式刷失败: {ex.Message}", "文档不加班");
            }
        }

        private int ReplaceInDocument(Word.Document doc, BatchReplaceExecutionRequest request)
        {
            int totalCount = 0;

            foreach (BatchReplaceRule rule in request.Rules)
            {
                if (!rule.Enabled || string.IsNullOrWhiteSpace(rule.FindText))
                {
                    continue;
                }

                if (rule.Scope == BatchFindScope.Body || rule.Scope == BatchFindScope.All)
                {
                    Word.Range contentRange = null;
                    try
                    {
                        contentRange = doc.Content;
                        totalCount += ReplaceInRange(contentRange, rule, request.FindOnly);
                    }
                    finally
                    {
                        ReleaseComObject(contentRange);
                    }
                }

                if (rule.Scope == BatchFindScope.HeaderFooter || rule.Scope == BatchFindScope.All)
                {
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdPrimaryHeaderStory, rule, request.FindOnly);
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdEvenPagesHeaderStory, rule, request.FindOnly);
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdFirstPageHeaderStory, rule, request.FindOnly);
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdPrimaryFooterStory, rule, request.FindOnly);
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdEvenPagesFooterStory, rule, request.FindOnly);
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdFirstPageFooterStory, rule, request.FindOnly);
                }

                if (rule.Scope == BatchFindScope.Special || rule.Scope == BatchFindScope.All)
                {
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdFootnotesStory, rule, request.FindOnly);
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdEndnotesStory, rule, request.FindOnly);
                    totalCount += ReplaceByStoryType(doc, Word.WdStoryType.wdTextFrameStory, rule, request.FindOnly);
                }
            }

            return totalCount;
        }
        private int BuildRenamedPath(string filePath, BatchReplaceExecutionRequest request, out string finalPath)
        {
            finalPath = filePath;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return 0;
            }

            string directory = Path.GetDirectoryName(filePath);
            string originalFileName = Path.GetFileName(filePath);
            string originalExtension = Path.GetExtension(filePath);

            if (string.IsNullOrWhiteSpace(directory))
            {
                return 0;
            }

            int totalCount = 0;
            string updatedFileName = originalFileName;

            foreach (BatchReplaceRule rule in request.Rules)
            {
                if (!rule.Enabled || !rule.ApplyToFileName || string.IsNullOrWhiteSpace(rule.FindText))
                {
                    continue;
                }

                int hits = CountMatchesInText(updatedFileName, rule);
                if (hits <= 0)
                {
                    continue;
                }

                totalCount += hits;
                updatedFileName = ReplaceInText(updatedFileName, rule);
            }

            if (totalCount == 0)
            {
                return 0;
            }

            if (string.IsNullOrWhiteSpace(updatedFileName))
            {
                throw new IOException("替换后文件名为空。");
            }

            char[] invalidChars = Path.GetInvalidFileNameChars();
            if (updatedFileName.IndexOfAny(invalidChars) >= 0)
            {
                throw new IOException("替换后文件名包含非法字符: " + updatedFileName);
            }

            if (string.IsNullOrEmpty(Path.GetExtension(updatedFileName)) && !string.IsNullOrEmpty(originalExtension))
            {
                updatedFileName += originalExtension;
            }

            string targetPath = Path.Combine(directory, updatedFileName);
            finalPath = targetPath;
            return totalCount;
        }

        private int CountMatchesInText(string input, BatchReplaceRule rule)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(rule.FindText))
            {
                return 0;
            }

            bool useWholeWordForFileName = rule.MatchWholeWord && ShouldUseWholeWordForFileName(rule.FindText);

            if (rule.FindType == BatchFindType.WordWildcards)
            {
                string pattern = Regex.Escape(rule.FindText)
                    .Replace("\\*", ".*")
                    .Replace("\\?", ".");

                RegexOptions options = rule.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                return Regex.Matches(input, pattern, options).Count;
            }

            StringComparison comparison = rule.MatchCase
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            int count = 0;
            int start = 0;

            while (start <= input.Length)
            {
                int index = input.IndexOf(rule.FindText, start, comparison);
                if (index < 0)
                {
                    break;
                }

                if (!useWholeWordForFileName || IsWholeWordMatch(input, index, rule.FindText.Length))
                {
                    count++;
                }

                start = index + Math.Max(1, rule.FindText.Length);
            }

            return count;
        }

        private string ReplaceInText(string input, BatchReplaceRule rule)
        {
            if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(rule.FindText))
            {
                return input;
            }

            bool useWholeWordForFileName = rule.MatchWholeWord && ShouldUseWholeWordForFileName(rule.FindText);

            if (rule.FindType == BatchFindType.WordWildcards)
            {
                string pattern = Regex.Escape(rule.FindText)
                    .Replace("\\*", ".*")
                    .Replace("\\?", ".");

                RegexOptions options = rule.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase;
                return Regex.Replace(input, pattern, rule.ReplaceText ?? string.Empty, options);
            }

            StringComparison comparison = rule.MatchCase
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            string replacement = rule.ReplaceText ?? string.Empty;
            int start = 0;
            System.Text.StringBuilder sb = new System.Text.StringBuilder();

            while (start < input.Length)
            {
                int index = input.IndexOf(rule.FindText, start, comparison);
                if (index < 0)
                {
                    sb.Append(input, start, input.Length - start);
                    break;
                }

                bool canReplace = !useWholeWordForFileName || IsWholeWordMatch(input, index, rule.FindText.Length);
                if (!canReplace)
                {
                    sb.Append(input, start, index - start + 1);
                    start = index + 1;
                    continue;
                }

                sb.Append(input, start, index - start);
                sb.Append(replacement);
                start = index + rule.FindText.Length;
            }

            return sb.ToString();
        }

        private bool IsWholeWordMatch(string input, int startIndex, int matchLength)
        {
            int left = startIndex - 1;
            int right = startIndex + matchLength;

            bool leftOk = left < 0 || !IsWordChar(input[left]);
            bool rightOk = right >= input.Length || !IsWordChar(input[right]);

            return leftOk && rightOk;
        }

        private bool IsWordChar(char c)
        {
            return char.IsLetterOrDigit(c) || c == '_';
        }

        private bool ShouldUseWholeWordForFileName(string findText)
        {
            if (string.IsNullOrWhiteSpace(findText))
            {
                return false;
            }

            foreach (char c in findText)
            {
                if (c > 127)
                {
                    return false;
                }

                if (!(char.IsLetterOrDigit(c) || c == '_'))
                {
                    return false;
                }
            }

            return true;
        }

        private int ReplaceByStoryType(
            Word.Document doc,
            Word.WdStoryType storyType,
            BatchReplaceRule rule,
            bool findOnly)
        {
            int totalCount = 0;
            Word.StoryRanges storyRanges = null;
            Word.Range current = null;

            try
            {
                storyRanges = doc.StoryRanges;
                current = storyRanges[storyType];
            }
            catch
            {
                return 0;
            }
            finally
            {
                ReleaseComObject(storyRanges);
            }

            while (current != null)
            {
                Word.Range next = null;
                try
                {
                    totalCount += ReplaceInRange(current, rule, findOnly);
                    try
                    {
                        next = current.NextStoryRange;
                    }
                    catch
                    {
                        next = null;
                    }
                }
                finally
                {
                    ReleaseComObject(current);
                }

                current = next;
            }

            return totalCount;
        }

        private int ReplaceInRange(Word.Range range, BatchReplaceRule rule, bool findOnly)
        {
            int count = CountMatches(range, rule, findOnly);
            if (count == 0 || findOnly || rule.HighlightOnly)
            {
                return count;
            }

            Word.Range replaceRange = null;
            Word.Find find = null;
            Word.Replacement replacement = null;

            try
            {
                replaceRange = range.Duplicate;
                find = replaceRange.Find;
                replacement = find.Replacement;
                find.ClearFormatting();
                replacement.ClearFormatting();

                find.Execute(
                    FindText: rule.FindText,
                    MatchCase: rule.MatchCase,
                    MatchWholeWord: rule.MatchWholeWord,
                    MatchWildcards: rule.FindType == BatchFindType.WordWildcards,
                    MatchSoundsLike: rule.MatchSoundsLike,
                    MatchAllWordForms: rule.MatchAllWordForms,
                    Forward: true,
                    Wrap: Word.WdFindWrap.wdFindStop,
                    Format: false,
                    ReplaceWith: rule.ReplaceText ?? string.Empty,
                    Replace: Word.WdReplace.wdReplaceAll);
            }
            finally
            {
                ReleaseComObject(replacement);
                ReleaseComObject(find);
                ReleaseComObject(replaceRange);
            }

            return count;
        }

        private int CountMatches(Word.Range range, BatchReplaceRule rule, bool findOnly)
        {
            if (string.IsNullOrEmpty(rule.FindText))
            {
                return 0;
            }

            int count = 0;
            Word.Range scanRange = null;
            Word.Find find = null;
            Word.Replacement replacement = null;

            try
            {
                scanRange = range.Duplicate;
                find = scanRange.Find;
                replacement = find.Replacement;
                find.ClearFormatting();
                replacement.ClearFormatting();

                while (find.Execute(
                    FindText: rule.FindText,
                    MatchCase: rule.MatchCase,
                    MatchWholeWord: rule.MatchWholeWord,
                    MatchWildcards: rule.FindType == BatchFindType.WordWildcards,
                    MatchSoundsLike: rule.MatchSoundsLike,
                    MatchAllWordForms: rule.MatchAllWordForms,
                    Forward: true,
                    Wrap: Word.WdFindWrap.wdFindStop,
                    Format: false,
                    ReplaceWith: string.Empty,
                    Replace: Word.WdReplace.wdReplaceNone))
                {
                    count++;

                    if (findOnly || rule.HighlightOnly)
                    {
                        scanRange.HighlightColorIndex = Word.WdColorIndex.wdYellow;
                    }

                    scanRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
            }
            finally
            {
                ReleaseComObject(replacement);
                ReleaseComObject(find);
                ReleaseComObject(scanRange);
            }

            return count;
        }
        private Word.Document FindOpenedDocumentByPath(Word.Application app, string fullPath)
        {
            Word.Documents documents = null;

            try
            {
                documents = app.Documents;
                int documentCount = documents.Count;

                for (int index = 1; index <= documentCount; index++)
                {
                    Word.Document candidate = null;
                    bool isMatch = false;

                    try
                    {
                        object itemIndex = index;
                        candidate = documents.get_Item(ref itemIndex);
                        isMatch = string.Equals(candidate.FullName, fullPath, StringComparison.OrdinalIgnoreCase);
                        if (isMatch)
                        {
                            return candidate;
                        }
                    }
                    finally
                    {
                        if (!isMatch)
                        {
                            ReleaseComObject(candidate);
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(documents);
            }

            return null;
        }

        private void ReleaseComObject(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (InvalidComObjectException)
            {
                // The object was already released while Word closed the document.
            }
            catch (COMException)
            {
                // Cleanup must not interrupt processing of the remaining files.
            }
        }

    }
}
