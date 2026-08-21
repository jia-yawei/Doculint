using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        // 正则表达式：匹配标题前的手动编号（如“第1章”“(一)”“1.1.1”“1、”）
        private static readonly Regex[] ManualHeadingPrefixPatterns = new[]
        {
            new Regex(@"^\s*第\s*[0-9一二三四五六七八九十百千万零〇壹贰叁肆伍陆柒捌玖拾]+\s*[章节篇条款项目]\s*", RegexOptions.Compiled),
            new Regex(@"^\s*[（(]\s*[0-9一二三四五六七八九十百千万零〇壹贰叁肆伍陆柒捌玖拾]+\s*[)）]\s*", RegexOptions.Compiled),
            new Regex(@"^\s*[0-9一二三四五六七八九十百千万零〇壹贰叁肆伍陆柒捌玖拾]+(?:\s*[\.．、:：\-]\s*[0-9一二三四五六七八九十百千万零〇壹贰叁肆伍陆柒捌玖拾]+)+(?:\s*[)）])?\s*", RegexOptions.Compiled),
            new Regex(@"^\s*[0-9一二三四五六七八九十百千万零〇壹贰叁肆伍陆柒捌玖拾]+\s*[、:：\-]\s*", RegexOptions.Compiled)
        };
        private static readonly Regex ArabicManualHeadingPrefixPattern =
            new Regex(@"^\s*[0-9]+(?:\s*[\.．、:：\-]\s*[0-9]+)*(?:\s*[\.．、:：\-])?\s*", RegexOptions.Compiled);

        // 点击【更新所选章节号】按钮：更新光标所在段落或当前选区内的标题多级列表
        private void btnRebuildOutlineList_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteRebuildOutlineList();
        }

        // 点击【章节号修复】按钮：检查连续性，发现异常时重建全文标题编号。
        private void btnRepairChapterNumbers_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteRepairChapterNumbers();
        }

        private void btnRepairHeadingStyles_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("当前没有活动文档。", "标题样式修复");
                return;
            }

            try
            {
                string styleLibraryDocumentKey = GetCommonStylesDocumentKey(doc);
                List<string> cachedStyleNames = GetCachedHeadingStyleNames(styleLibraryDocumentKey);
                if (cachedStyleNames.Count == 0)
                {
                    MessageBox.Show(
                        "当前文档尚未加载样式库，请先打开“添加常用样式”面板并点击“加载样式库”。",
                        "标题样式修复",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }
                using (HeadingStyleRepairForm form = new HeadingStyleRepairForm(
                    Enumerable.Range(1, 9),
                    cachedStyleNames))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    IReadOnlyDictionary<int, string> selectedStyles = form.SelectedStyles;
                    List<HeadingStyleEntry> entries;
                    int repairedCount = 0;
                    int scanStart = Math.Max(doc.Content.Start, GetScanStartAfterToc(doc));
                    using (OutlineProgressForm progressForm =
                        new OutlineProgressForm("标题样式修复", "正在定位标题"))
                    {
                        progressForm.Show();
                        progressForm.ReportProgress(
                            0,
                            1,
                            scanStart > doc.Content.Start
                                ? "正在目录之后按所选大纲级别定位标题..."
                                : "正在按所选大纲级别定位标题...");
                        entries = CollectHeadingStyleEntries(
                            doc,
                            new HashSet<int>(selectedStyles.Keys),
                            scanStart,
                            (current, total, message) => progressForm.ReportProgress(current, total, message));

                        if (entries.Count > 0)
                        {
                            progressForm.ReportProgress(0, entries.Count, "正在应用目标样式...");
                            using (new WordPerformanceScope(app))
                            {
                                repairedCount = ApplyHeadingStyles(
                                    entries,
                                    selectedStyles,
                                    (current, total, message) => progressForm.ReportProgress(current, total, message));
                            }

                            app?.ScreenRefresh();
                        }
                    }

                    if (entries.Count == 0)
                    {
                        MessageBox.Show("当前文档中未找到所选大纲级别的标题。", "标题样式修复");
                        return;
                    }

                    TryUpdateStatusBar(app, "标题样式修复完成");
                    MessageBox.Show(
                        repairedCount > 0
                            ? $"标题样式修复完成。已找到 {entries.Count} 个所选级别标题，共修复 {repairedCount} 个标题。"
                            : $"已检查 {entries.Count} 个所选级别标题，均已使用指定样式，无需修复。",
                        "标题样式修复",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException)
            {
                TryUpdateStatusBar(app, "标题样式修复已停止");
                MessageBox.Show("标题样式修复已停止。", "标题样式修复", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("标题样式修复失败：\r\n" + ex.Message, "标题样式修复", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static List<string> GetCachedHeadingStyleNames(string documentKey)
        {
            if (string.IsNullOrWhiteSpace(documentKey) ||
                !headingStyleLibraryCache.TryGetValue(documentKey, out List<string> names))
            {
                return new List<string>();
            }

            return names.ToList();
        }

        private static List<string> LoadHeadingStyleNames(
            Word.Document doc,
            string documentKey,
            Action<int, int> progress)
        {
            List<string> names = GetDocumentParagraphStyleNames(doc, progress);
            if (names.Count > 0 && !string.IsNullOrWhiteSpace(documentKey))
            {
                headingStyleLibraryCache[documentKey] = names.ToList();
            }

            return names;
        }

        private List<HeadingStyleEntry> CollectHeadingStyleEntries(
            Word.Document doc,
            HashSet<int> selectedLevels,
            int scanStart,
            Action<int, int, string> progress)
        {
            List<HeadingStyleEntry> entries = new List<HeadingStyleEntry>();
            if (doc?.Content == null || selectedLevels == null || selectedLevels.Count == 0)
            {
                return entries;
            }

            ResetOperationCancellation("读取标题样式");
            List<Word.Range> headingRanges = CollectOutlineRangesByFind(
                doc,
                selectedLevels,
                Math.Max(doc.Content.Start, scanStart),
                doc.Content.End,
                progress);
            int total = headingRanges.Count;
            for (int index = 0; index < total; index++)
            {
                ThrowIfOperationCancelled();
                try
                {
                    Word.Paragraph paragraph = GetHostParagraph(headingRanges[index]);
                    int level = GetParagraphOutlineLevel(paragraph);
                    if (!selectedLevels.Contains(level) || paragraph?.Range == null)
                    {
                        continue;
                    }

                    string styleName = ResolveStyleName(TryGetStyle(paragraph.Range), doc);
                    if (string.IsNullOrWhiteSpace(styleName))
                    {
                        continue;
                    }

                    entries.Add(new HeadingStyleEntry(level, paragraph.Range.Duplicate, styleName));
                }
                catch
                {
                }

                if (index == 0 || index == total - 1 || index % 10 == 0)
                {
                    progress?.Invoke(index + 1, Math.Max(1, total), "正在读取标题当前样式...");
                }
            }

            return entries;
        }

        private static int ApplyHeadingStyles(
            IEnumerable<HeadingStyleEntry> entries,
            IReadOnlyDictionary<int, string> selectedStyles,
            Action<int, int, string> progress = null)
        {
            if (entries == null || selectedStyles == null || selectedStyles.Count == 0)
            {
                return 0;
            }

            ResetOperationCancellation("标题样式修复");
            int repairedCount = 0;
            List<HeadingStyleEntry> entryList = entries.ToList();
            int total = entryList.Count;
            for (int index = 0; index < total; index++)
            {
                ThrowIfOperationCancelled();
                HeadingStyleEntry entry = entryList[index];
                if (!selectedStyles.TryGetValue(entry.Level, out string targetStyle) ||
                    string.IsNullOrWhiteSpace(targetStyle) ||
                    string.Equals(entry.StyleName, targetStyle, StringComparison.OrdinalIgnoreCase))
                {
                    if (index == 0 || index == total - 1 || index % 10 == 0)
                    {
                        progress?.Invoke(index + 1, Math.Max(1, total), "正在应用目标样式...");
                    }
                    continue;
                }

                if (TrySetStyle(entry.Range, targetStyle))
                {
                    repairedCount++;
                }

                if (index == 0 || index == total - 1 || index % 10 == 0)
                {
                    progress?.Invoke(index + 1, Math.Max(1, total), "正在应用目标样式...");
                }
            }

            try
            {
                Globals.ThisAddIn?.Application?.ScreenRefresh();
            }
            catch
            {
            }

            return repairedCount;
        }

        private sealed class HeadingStyleEntry
        {
            internal HeadingStyleEntry(int level, Word.Range range, string styleName)
            {
                Level = level;
                Range = range;
                StyleName = styleName;
            }

            internal int Level { get; }

            internal Word.Range Range { get; }

            internal string StyleName { get; }
        }

        private void ExecuteRepairChapterNumbers()
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("当前没有活动文档。", "文档不加班 章节号");
                return;
            }

            DialogResult confirmation = MessageBox.Show(
                "章节号修复将执行以下操作：\r\n\r\n" +
                "1. 从目录之后开始检查；未找到目录时从第一页开始。\r\n" +
                "2. 仅检查 1-9 级大纲标题的自动章节号是否缺失、跳号或不连续。\r\n" +
                "3. 发现异常时，按标题大纲级别重新应用多级列表编号，并在完成后再次检查。\r\n\r\n" +
                "此操作可能修改文档中的标题编号，建议先保存当前文档。\r\n\r\n" +
                "是否确认开始修复？",
                "确认章节号修复",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            if (confirmation != DialogResult.OK)
            {
                return;
            }

            OutlineListRebuildOptions options = new OutlineListRebuildOptions
            {
                SelectedLevels = new HashSet<int>(Enumerable.Range(1, 9)),
                ClearManualNumbering = false,
                NumberPattern = outlineNumberPattern,
                Alignment = (int)Word.WdListLevelAlignment.wdListLevelAlignLeft,
                TrailingCharacter = (int)Word.WdTrailingCharacter.wdTrailingNone,
                NumberTextSpacing = outlineNumberTextSpacing
            };

            try
            {
                ResetOperationCancellation("章节号修复");
                int scanStart = Math.Max(doc.Content.Start, GetScanStartAfterToc(doc));
                int initialCheckedHeadingCount;
                List<string> styleConflicts;
                List<string> initialIssues;
                List<Word.Range> initialHeadingRanges;
                using (OutlineProgressForm checkForm =
                    new OutlineProgressForm("章节号修复", "正在检查章节号"))
                {
                    checkForm.Show();
                    checkForm.ReportProgress(0, 1, "正在定位目录之后的标题...");
                    initialHeadingRanges = CollectOutlineRangesByFind(
                        doc,
                        new HashSet<int>(Enumerable.Range(1, 9)),
                        scanStart,
                        doc.Content.End,
                        (current, total, message) => checkForm.ReportProgress(current, total, message));
                    checkForm.ReportProgress(0, Math.Max(1, initialHeadingRanges.Count), "正在检查标题连续性和样式...");
                    initialIssues = VerifyChapterNumberContinuity(
                        doc,
                        scanStart,
                        out initialCheckedHeadingCount,
                        out styleConflicts,
                        (current, total, message) => checkForm.ReportProgress(current, total, message),
                        initialHeadingRanges);
                }
                if (initialCheckedHeadingCount == 0)
                {
                    MessageBox.Show(
                        "当前文档中未找到 1-9 级大纲标题，无法检查章节号。",
                        "章节号修复",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                if (initialIssues.Count == 0)
                {
                    ShowChapterNumberRepairResult(
                        initialCheckedHeadingCount,
                        0,
                        initialIssues,
                        styleConflicts,
                        false);
                    return;
                }

                int finalCheckedHeadingCount;
                List<string> finalStyleConflicts;
                List<string> finalIssues;
                using (OutlineProgressForm progressForm =
                    new OutlineProgressForm("章节号修复", "正在修复章节号"))
                {
                    progressForm.Show();
                    progressForm.ReportProgress(0, 1, "正在准备修复章节号...");
                    using (new WordPerformanceScope(app))
                    {
                        RebuildOutlineListWithOptions(
                            app,
                            doc,
                            options,
                            scanStart,
                            (current, total, message) =>
                            {
                                progressForm.ReportProgress(current, total, message);
                            },
                            initialHeadingRanges);

                        ChapterNumberRepairFormattingSettings formatting = GetChapterNumberRepairFormattingSettings();
                        if (formatting.ApplyFormatting)
                        {
                            progressForm.ReportProgress(0, 1, "正在应用标题文字和段落格式...");
                            ApplyChapterNumberRepairFormatting(doc, scanStart, formatting, initialHeadingRanges);
                        }
                    }

                    app?.ScreenRefresh();
                    progressForm.ReportProgress(0, 1, "正在复查修复结果...");
                    finalIssues = VerifyChapterNumberContinuity(
                        doc,
                        scanStart,
                        out finalCheckedHeadingCount,
                        out finalStyleConflicts,
                        (current, total, message) => progressForm.ReportProgress(current, total, message));
                }

                app?.ScreenRefresh();
                TryUpdateStatusBar(app, finalIssues.Count == 0 ? "章节号修复完成" : "章节号修复后仍有异常");
                ShowChapterNumberRepairResult(
                    initialCheckedHeadingCount,
                    Math.Max(0, initialIssues.Count - finalIssues.Count),
                    initialIssues,
                    styleConflicts,
                    finalIssues.Count > 0);
            }
            catch (OperationCanceledException)
            {
                TryUpdateStatusBar(app, "章节号修复已停止");
                MessageBox.Show("章节号修复已停止。", "章节号修复", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "章节号修复", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static void ApplyChapterNumberRepairFormatting(
            Word.Document doc,
            int scanStart,
            ChapterNumberRepairFormattingSettings settings,
            IReadOnlyList<Word.Range> headingRanges = null)
        {
            if (doc?.Content == null || settings == null)
            {
                return;
            }

            List<Word.Range> headings = headingRanges == null
                ? CollectOutlineRangesByFind(
                    doc,
                    new HashSet<int>(Enumerable.Range(1, 9)),
                    Math.Max(doc.Content.Start, scanStart),
                    doc.Content.End)
                : headingRanges.ToList();
            for (int index = 0; index < headings.Count; index++)
            {
                ThrowIfOperationCancelled();
                Word.Range heading = headings[index];
                Word.Paragraph paragraph = GetHostParagraph(heading);
                int level = GetParagraphOutlineLevel(paragraph);
                if (level < 1 || level > 9 || heading == null)
                {
                    continue;
                }

                string fontName = level == 1 ? settings.LevelOneFontName : settings.OtherLevelsFontName;
                float fontSize = (float)(level == 1 ? settings.LevelOneFontSize : settings.OtherLevelsFontSize);
                try
                {
                    if (!string.IsNullOrWhiteSpace(fontName))
                    {
                        heading.Font.Name = fontName;
                        heading.Font.NameFarEast = fontName;
                    }
                    heading.Font.Size = fontSize;
                    heading.Font.Bold = settings.Bold ? 1 : 0;

                    Word.ParagraphFormat format = heading.ParagraphFormat;
                    format.Alignment = (Word.WdParagraphAlignment)settings.Alignment;
                    format.LineSpacingRule = settings.LineSpacingRule == 1
                        ? Word.WdLineSpacing.wdLineSpace1pt5
                        : settings.LineSpacingRule == 2
                            ? Word.WdLineSpacing.wdLineSpaceDouble
                            : Word.WdLineSpacing.wdLineSpaceSingle;
                    format.SpaceBefore = (float)settings.SpaceBefore;
                    format.SpaceAfter = (float)settings.SpaceAfter;
                }
                catch
                {
                    // Skip an individual protected or unsupported paragraph and continue repairing.
                }
            }
        }

        private static List<string> VerifyChapterNumberContinuity(
            Word.Document doc,
            int scanStart,
            out int checkedHeadingCount,
            out List<string> styleConflicts,
            Action<int, int, string> progress = null,
            IReadOnlyList<Word.Range> headingRangesOverride = null)
        {
            checkedHeadingCount = 0;
            styleConflicts = new List<string>();
            List<string> issues = new List<string>();
            int[] counters = new int[10];
            HashSet<int> levels = new HashSet<int>(Enumerable.Range(1, 9));
            Dictionary<int, Dictionary<string, List<string>>> titlesByLevelAndStyle =
                new Dictionary<int, Dictionary<string, List<string>>>();

            List<Word.Range> headingRanges = (headingRangesOverride == null
                ? CollectOutlineRangesByFind(
                    doc,
                    levels,
                    scanStart,
                    doc.Content.End,
                    progress)
                : headingRangesOverride.ToList())
                .OrderBy(item => item.Start)
                .ToList();
            int headingIndex = 0;
            foreach (Word.Range range in headingRanges)
            {
                ThrowIfOperationCancelled();
                headingIndex++;
                progress?.Invoke(
                    headingIndex,
                    Math.Max(1, headingRanges.Count),
                    "正在检查标题段落...");
                Word.Paragraph paragraph = GetHostParagraph(range);
                if (paragraph?.Range == null || IsInTable(paragraph.Range))
                {
                    continue;
                }

                int level = GetParagraphOutlineLevel(paragraph);
                string styleName = GetParagraphStyleName(paragraph);
                if (!string.IsNullOrWhiteSpace(styleName))
                {
                    if (!titlesByLevelAndStyle.TryGetValue(
                        level,
                        out Dictionary<string, List<string>> titlesByStyle))
                    {
                        titlesByStyle = new Dictionary<string, List<string>>(
                            StringComparer.CurrentCultureIgnoreCase);
                        titlesByLevelAndStyle[level] = titlesByStyle;
                    }

                    if (!titlesByStyle.TryGetValue(styleName, out List<string> titles))
                    {
                        titles = new List<string>();
                        titlesByStyle[styleName] = titles;
                    }

                    string title = CleanParagraphText(range.Text);
                    titles.Add(string.IsNullOrWhiteSpace(title) ? "<空标题>" : title);
                }

                checkedHeadingCount++;
                string listString = GetListString(paragraph);
                int[] actualNumbers = ParseChapterNumber(listString);
                if (actualNumbers.Length < level)
                {
                    issues.Add(BuildChapterNumberIssue(level, range, "没有自动章节号"));
                    continue;
                }

                int[] expectedNumbers = BuildExpectedNumbers(counters, level);
                string expectedText = string.Join(".", expectedNumbers);
                string actualText = string.Join(".", actualNumbers.Take(level));
                if (!string.Equals(actualText, expectedText, StringComparison.Ordinal))
                {
                    issues.Add(BuildChapterNumberIssue(level, range, $"应为 {expectedText}，当前为 {listString}"));
                }

                ApplyActualCounters(counters, actualNumbers, level);
            }

            foreach (KeyValuePair<int, Dictionary<string, List<string>>> item in
                titlesByLevelAndStyle.OrderBy(pair => pair.Key))
            {
                if (item.Value.Count <= 1)
                {
                    continue;
                }

                KeyValuePair<string, List<string>> majority = item.Value
                    .OrderByDescending(pair => pair.Value.Count)
                    .ThenBy(pair => pair.Key, StringComparer.CurrentCultureIgnoreCase)
                    .First();
                foreach (KeyValuePair<string, List<string>> mismatch in item.Value
                    .Where(pair => !string.Equals(
                        pair.Key,
                        majority.Key,
                        StringComparison.CurrentCultureIgnoreCase)))
                {
                    foreach (string title in mismatch.Value.Take(3))
                    {
                        string displayTitle = title.Length > 32 ? title.Substring(0, 32) + "..." : title;
                        styleConflicts.Add(
                            $"{item.Key}级标题“{displayTitle}”使用“{mismatch.Key}”，同级多数标题使用“{majority.Key}”。");
                    }
                }
            }

            return issues;
        }

        private static int[] ParseChapterNumber(string listString)
        {
            string value = (listString ?? string.Empty).Trim();
            value = value.TrimStart('(', '（').TrimEnd(')', '）').Trim();
            return ParseDottedNumber(value);
        }

        private static string BuildChapterNumberIssue(int level, Word.Range range, string reason)
        {
            string title = CleanParagraphText(range?.Text);
            if (title.Length > 32)
            {
                title = title.Substring(0, 32) + "...";
            }

            return string.IsNullOrWhiteSpace(title)
                ? $"{level}级标题：{reason}。"
                : $"{level}级标题“{title}”：{reason}。";
        }

        private static void ShowChapterNumberRepairResult(
            int totalHeadingCount,
            int repairedHeadingCount,
            IEnumerable<string> numberingIssues,
            IEnumerable<string> styleIssues,
            bool hasRemainingNumberingIssues)
        {
            List<string> styles = (styleIssues ?? Enumerable.Empty<string>())
                .Where(issue => !string.IsNullOrWhiteSpace(issue))
                .ToList();
            List<string> problems = (numberingIssues ?? Enumerable.Empty<string>())
                .Concat(styles)
                .Where(issue => !string.IsNullOrWhiteSpace(issue))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            const int maxDisplayedItems = 20;

            StringBuilder message = new StringBuilder();
            message.AppendLine($"总标题数量：{totalHeadingCount}");
            message.AppendLine($"修复数量：{repairedHeadingCount}");

            if (problems.Count > 0)
            {
                message.AppendLine();
                message.AppendLine("发现的问题：");
                foreach (string issue in problems.Take(maxDisplayedItems))
                {
                    message.AppendLine("- " + issue);
                }

                if (problems.Count > maxDisplayedItems)
                {
                    message.AppendLine($"- 另有 {problems.Count - maxDisplayedItems} 项未显示。");
                }
            }
            else
            {
                message.AppendLine();
                message.AppendLine("未发现问题。");
            }

            MessageBox.Show(
                message.ToString().TrimEnd(),
                "章节号修复",
                MessageBoxButtons.OK,
                hasRemainingNumberingIssues || styles.Count > 0
                    ? MessageBoxIcon.Warning
                    : MessageBoxIcon.Information);
        }

        private void ExecuteRebuildOutlineList()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app.ActiveDocument;

            // 无文档则提示
            if (doc == null)
            {
                MessageBox.Show("当前没有活动文档。", "文档不加班 多级列表");
                return;
            }

            DialogResult confirmResult = MessageBox.Show(
                "是否确认更新光标所在段落或当前选区内标题段落的章节号？",
                "更新所选章节号",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            if (confirmResult != DialogResult.OK)
            {
                return;
            }

            HashSet<int> selectedLevels = new HashSet<int>(Enumerable.Range(1, 9));
            OutlineListRebuildOptions defaultOptions = new OutlineListRebuildOptions
            {
                SelectedLevels = selectedLevels,
                ClearManualNumbering = false,
                NumberPattern = outlineNumberPattern,
                Alignment = (int)Word.WdListLevelAlignment.wdListLevelAlignLeft,
                TrailingCharacter = (int)Word.WdTrailingCharacter.wdTrailingNone,
                NumberTextSpacing = outlineNumberTextSpacing
            };

            try
            {
                ResetOperationCancellation("更新所选章节号");
                using (OutlineProgressForm progressForm = new OutlineProgressForm())
                {
                    progressForm.Show();
                    progressForm.ReportProgress(0, 1, "正在准备文档...");

                    using (new WordPerformanceScope(app))
                    {
                        RebuildOutlineListForCurrentSelection(app, doc, defaultOptions, (current, total, message) =>
                        {
                            progressForm.ReportProgress(current, total, message);
                        });
                    }
                }

                MessageBox.Show("章节号更新完成。", "更新所选章节号");
            }
            catch (OperationCanceledException)
            {
                TryUpdateStatusBar(app, "章节号更新已停止");
                MessageBox.Show("章节号更新已停止。", "更新所选章节号", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "更新所选章节号", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnClearManualHeadingNumbers_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ResetOperationCancellation("清除标题手工编号");
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班 清理");
                    return;
                }

                DialogResult confirmResult = MessageBox.Show(
                    "清除标题前的手工编号只会处理已设置为 1-9 级大纲级别的标题段落。\r\n请先确认需要清理的标题段落已经正确设置大纲级别。\r\n\r\n该操作会删除标题开头手动输入的阿拉伯数字，可能误删本来就属于标题内容的数字。\r\n是否确认删除？",
                    "清除标题前的手工编号",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Warning);
                if (confirmResult != DialogResult.OK)
                {
                    return;
                }

                int clearedCount;
                using (new WordPerformanceScope(app))
                {
                    clearedCount = ClearManualHeadingNumbersInDocument(doc);
                }

                MessageBox.Show(
                    clearedCount > 0 ? $"已清除 {clearedCount} 个标题前的手工编号。" : "未发现需要清除的标题手工编号。",
                    "文档不加班 清理");
            }
            catch (OperationCanceledException)
            {
                TryUpdateStatusBar(Globals.ThisAddIn?.Application, "清除标题手工编号已停止");
                MessageBox.Show("清除标题手工编号已停止。", "文档不加班 清理");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"清除标题手工编号失败: {ex.Message}", "文档不加班 清理");
            }
        }

        private int ClearManualHeadingNumbersInDocument(Word.Document doc)
        {
            if (doc?.Content == null)
            {
                return 0;
            }

            ThrowIfOperationCancelled();
            int deepestLevel = FindDeepestOutlineLevel(doc, doc.Content.Start, doc.Content.End);
            if (deepestLevel == 0)
            {
                return 0;
            }

            HashSet<int> seenStarts = new HashSet<int>();
            List<OutlineParagraphSnapshot> targets = CollectOutlineRangesByFind(
                doc,
                new HashSet<int>(Enumerable.Range(1, deepestLevel)),
                doc.Content.Start,
                doc.Content.End)
                .Select(range =>
                {
                    ThrowIfOperationCancelled();
                    Word.Paragraph paragraph = GetHostParagraph(range);
                    Word.Range paragraphRange = paragraph?.Range;
                    if (paragraphRange == null || !seenStarts.Add(paragraphRange.Start))
                    {
                        return null;
                    }

                    return new OutlineParagraphSnapshot
                    {
                        Paragraph = paragraph,
                        Level = GetParagraphOutlineLevel(paragraph),
                        Start = paragraphRange.Start,
                        End = paragraphRange.End,
                        ManualPrefixLength = GetArabicManualHeadingPrefixLength(paragraphRange.Text)
                    };
                })
                .Where(item => item != null && item.HasManualPrefix)
                .ToList();
            if (targets.Count == 0)
            {
                return 0;
            }

            int currentStep = 0;
            return ClearManualHeadingPrefixes(doc, targets, targets.Count, ref currentStep, (current, total, message) => { });
        }

        private OutlineRebuildResult RebuildOutlineListForCurrentSelection(
            Word.Application app,
            Word.Document doc,
            OutlineListRebuildOptions options,
            Action<int, int, string> progress)
        {
            if (options == null || options.SelectedLevels == null || options.SelectedLevels.Count == 0)
            {
                throw new InvalidOperationException("请先选择要重建的级别。");
            }

            ThrowIfOperationCancelled();
            Word.Selection selection = app?.Selection;
            Word.Range selectionRange = selection?.Range;
            if (selectionRange == null)
            {
                throw new InvalidOperationException("当前没有可更新的光标或选区。");
            }

            bool collapsedSelection = selectionRange.Start == selectionRange.End;
            if (collapsedSelection)
            {
                ThrowIfOperationCancelled();
                Word.Paragraph paragraph = GetHostParagraph(selectionRange);
                if (paragraph?.Range == null)
                {
                    throw new InvalidOperationException("当前光标所在位置未找到段落。");
                }

                selectionRange = paragraph.Range;
                return RebuildCurrentOutlineParagraphOnly(doc, paragraph, options, progress);
            }

            int maxOutlineLevel = collapsedSelection
                ? 9
                : FindDeepestOutlineLevel(doc, selectionRange.Start, selectionRange.End);
            HashSet<int> selectedLevels = new HashSet<int>(
                options.SelectedLevels.Where(level => level >= 1 && level <= maxOutlineLevel));
            if (selectedLevels.Count == 0)
            {
                throw new InvalidOperationException("当前选区内未找到 1-9 级大纲标题。");
            }

            OutlineRebuildResult result = new OutlineRebuildResult
            {
                SelectedLevels = selectedLevels.OrderBy(x => x).ToList(),
                ScanScope = "当前选区"
            };

            Stopwatch sw = Stopwatch.StartNew();
            ThrowIfOperationCancelled();
            progress(0, 1, "正在读取当前选区标题...");

            List<OutlineParagraphSnapshot> targets = CollectSelectedOutlineParagraphs(
                doc,
                selectionRange.Start,
                selectionRange.End,
                options.ClearManualNumbering)
                .Where(item => selectedLevels.Contains(item.Level))
                .OrderBy(item => item.Start)
                .ToList();

            result.TargetParagraphCount = targets.Count;
            if (targets.Count == 0)
            {
                throw new InvalidOperationException("当前光标或选区内未找到 1-9 级大纲标题。请先设置段落大纲级别，或手动选择标题范围。");
            }

            int totalSteps = targets.Count;
            int currentStep = 0;
            Stopwatch phaseWatch = Stopwatch.StartNew();
            int manualPrefixCount = targets.Count(item => item.HasManualPrefix);
            if (options.ClearManualNumbering && manualPrefixCount > 0)
            {
                progress(0, totalSteps, "正在清理手工章节号...");
                int cleanupStep = 0;
                result.ClearedManualNumberCount = ClearManualHeadingPrefixes(
                    doc,
                    targets,
                    manualPrefixCount,
                    ref cleanupStep,
                    (current, total, message) => progress(0, totalSteps, message));
            }

            phaseWatch.Stop();
            result.CleanupMilliseconds = phaseWatch.ElapsedMilliseconds;

            currentStep = 0;
            progress(0, totalSteps, "正在建立多级列表模板...");
            ThrowIfOperationCancelled();
            Word.ListTemplate listTemplate = BuildOutlineListTemplate(
                doc,
                options,
                CalculateStartAtByLevel(doc, targets));

            phaseWatch.Restart();
            result.AppliedParagraphCount = ApplyOutlineListToParagraphs(
                targets,
                listTemplate,
                totalSteps,
                ref currentStep,
                progress);
            phaseWatch.Stop();
            result.ApplyMilliseconds = phaseWatch.ElapsedMilliseconds;

            sw.Stop();
            result.DurationMilliseconds = sw.ElapsedMilliseconds;
            progress(totalSteps, totalSteps, "更新完成");
            return result;
        }

        private OutlineRebuildResult RebuildCurrentOutlineParagraphOnly(
            Word.Document doc,
            Word.Paragraph paragraph,
            OutlineListRebuildOptions options,
            Action<int, int, string> progress)
        {
            int level = GetParagraphOutlineLevel(paragraph);
            ThrowIfOperationCancelled();
            if (level < 1 || level > 9)
            {
                throw new InvalidOperationException("当前光标所在段落不是 1-9 级大纲标题。请先设置段落大纲级别。");
            }

            progress(0, 1, "正在接续前一个章节号...");
            ApplyOutlineListToCurrentParagraph(doc, paragraph, level, options);

            progress(1, 1, "更新完成");
            return new OutlineRebuildResult
            {
                SelectedLevels = new List<int> { level },
                ScanScope = "当前段落",
                TargetParagraphCount = 1,
                AppliedParagraphCount = 1
            };
        }

        private void ApplyOutlineListToCurrentParagraph(
            Word.Document doc,
            Word.Paragraph paragraph,
            int level,
            OutlineListRebuildOptions options)
        {
            Word.Range paragraphRange = paragraph?.Range;
            ThrowIfOperationCancelled();
            if (doc == null || paragraphRange == null)
            {
                return;
            }

            Word.ListTemplate template = FindPreviousOutlineListTemplate(doc, paragraphRange.Start);
            bool continuePreviousList = template != null;
            if (template == null)
            {
                template = BuildOutlineListTemplate(doc, options);
            }

            TryRemoveExistingListFormatting(paragraphRange);
            paragraphRange.ParagraphFormat.LeftIndent = 0f;
            paragraphRange.ParagraphFormat.FirstLineIndent = 0f;
            paragraphRange.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate: template,
                ContinuePreviousList: continuePreviousList,
                ApplyTo: Word.WdListApplyTo.wdListApplyToSelection,
                DefaultListBehavior: Word.WdDefaultListBehavior.wdWord10ListBehavior,
                ApplyLevel: level);
        }

        private static Word.ListTemplate FindPreviousOutlineListTemplate(Word.Document doc, int currentStart)
        {
            if (doc?.Content == null || currentStart <= doc.Content.Start)
            {
                return null;
            }

            HashSet<int> levels = new HashSet<int>(Enumerable.Range(1, 9));
            List<Word.Range> previousHeadings = CollectOutlineRangesByFind(doc, levels, doc.Content.Start, currentStart);
            for (int i = previousHeadings.Count - 1; i >= 0; i--)
            {
                ThrowIfOperationCancelled();
                try
                {
                    Word.Paragraph previousParagraph = GetHostParagraph(previousHeadings[i]);
                    Word.ListFormat listFormat = previousParagraph?.Range?.ListFormat;
                    if (listFormat != null
                        && listFormat.ListType != Word.WdListType.wdListNoNumbering
                        && listFormat.ListTemplate != null)
                    {
                        return listFormat.ListTemplate;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private List<OutlineParagraphSnapshot> CollectSelectedOutlineParagraphs(
            Word.Document doc,
            int rangeStart,
            int rangeEnd,
            bool detectManualPrefixes)
        {
            List<OutlineParagraphSnapshot> targets = new List<OutlineParagraphSnapshot>();
            if (doc == null || rangeStart >= rangeEnd)
            {
                return targets;
            }

            Word.Range range = doc.Range(rangeStart, rangeEnd);
            Word.Paragraphs paragraphs = range.Paragraphs;
            int count = paragraphs?.Count ?? 0;
            HashSet<int> seenStarts = new HashSet<int>();
            for (int i = 1; i <= count; i++)
            {
                ThrowIfOperationCancelled();
                try
                {
                    Word.Paragraph paragraph = paragraphs[i];
                    int level = GetParagraphOutlineLevel(paragraph);
                    Word.Range paragraphRange = paragraph?.Range;
                    if (level < 1 || level > 9 || paragraphRange == null || !seenStarts.Add(paragraphRange.Start))
                    {
                        continue;
                    }

                    targets.Add(new OutlineParagraphSnapshot
                    {
                        Paragraph = paragraph,
                        Level = level,
                        Start = paragraphRange.Start,
                        End = paragraphRange.End,
                        ManualPrefixLength = detectManualPrefixes ? GetManualHeadingPrefixLength(paragraphRange.Text) : 0
                    });
                }
                catch
                {
                }
            }

            return targets.OrderBy(item => item.Start).ToList();
        }

        // 按设置重建大纲列表（核心方法）
        private OutlineRebuildResult RebuildOutlineListWithOptions(
            Word.Application app,
            Word.Document doc,
            OutlineListRebuildOptions options,
            int scanStart,
            Action<int, int, string> progress,
            IReadOnlyList<Word.Range> headingRangesOverride = null)
        {
            // 校验参数
            if (options == null || options.SelectedLevels == null || options.SelectedLevels.Count == 0)
            {
                throw new InvalidOperationException("请先选择要重建的级别。");
            }

            ThrowIfOperationCancelled();
            // 只保留1-9级
            HashSet<int> selectedLevels = new HashSet<int>(
                options.SelectedLevels.Where(level => level >= 1 && level <= 9));
            if (selectedLevels.Count == 0)
            {
                throw new InvalidOperationException("仅支持选择 1-9 级，请检查当前设置。");
            }

            // 初始化结果对象
            OutlineRebuildResult result = new OutlineRebuildResult
            {
                SelectedLevels = selectedLevels.OrderBy(x => x).ToList(),
                ScanScope = string.Empty
            };

            Stopwatch sw = Stopwatch.StartNew();
            // 关闭屏幕刷新，提高速度
            bool oldScreenUpdating = app.ScreenUpdating;
            app.ScreenUpdating = false;

            try
            {
                Stopwatch phaseWatch = new Stopwatch();

                Word.Range scanRange = GetMainStoryRange(doc, out string scanScope);
                int boundedScanStart = Math.Max(scanRange.Start, Math.Min(scanStart, scanRange.End));
                if (boundedScanStart > scanRange.Start)
                {
                    scanRange.SetRange(boundedScanStart, scanRange.End);
                    scanScope = "目录之后";
                }
                int currentStep = 0;

                ThrowIfOperationCancelled();
                progress(0, 0, "正在按大纲级别查找目标段落...");
                result.ScanScope = scanScope;

                // 1. 扫描符合条件的标题段落
                phaseWatch.Restart();
                List<OutlineParagraphSnapshot> targets = CollectOutlineParagraphSnapshots(
                    doc,
                    selectedLevels,
                    scanRange.Start,
                    scanRange.End,
                    options.ClearManualNumbering,
                    0,
                    ref currentStep,
                    progress,
                    headingRangesOverride);
                phaseWatch.Stop();
                result.ScanMilliseconds = phaseWatch.ElapsedMilliseconds;
                targets = targets.OrderBy(item => item.Start).ToList();
                result.TargetParagraphCount = targets.Count;

                // 没找到标题则退出
                if (targets.Count == 0)
                {
                    string levelsText = string.Join("、", selectedLevels.OrderBy(level => level));
                    throw new InvalidOperationException($"未找到 {levelsText} 级大纲标题。请确认标题段落已经设置为对应大纲级别。");
                }

                // 2. 清理手动编号
                int totalSteps = targets.Count;
                int appliedStep = 0;
                int manualPrefixCount = targets.Count(item => item.HasManualPrefix);
                result.ClearedManualNumberCount = 0;
                result.CleanupMilliseconds = 0;

                phaseWatch.Restart();
                if (options.ClearManualNumbering && manualPrefixCount > 0)
                {
                    progress(appliedStep, totalSteps, "正在清理手工章节号...");
                    int cleanupStep = 0;
                    result.ClearedManualNumberCount = ClearManualHeadingPrefixes(
                        doc,
                        targets,
                        manualPrefixCount,
                        ref cleanupStep,
                        (current, total, message) => progress(appliedStep, totalSteps, message));
                }
                phaseWatch.Stop();
                result.CleanupMilliseconds = phaseWatch.ElapsedMilliseconds;

                // 3. 按标题出现顺序，把大纲级别直接映射到同级多级列表。
                progress(++appliedStep, totalSteps, "正在建立新的多级列表模板...");
                ThrowIfOperationCancelled();
                Word.ListTemplate listTemplate = BuildOutlineListTemplate(doc, options);

                phaseWatch.Restart();
                result.AppliedParagraphCount = ApplyOutlineListToParagraphs(
                    targets,
                    listTemplate,
                    totalSteps,
                    ref appliedStep,
                    progress);
                phaseWatch.Stop();
                result.ApplyMilliseconds = phaseWatch.ElapsedMilliseconds;

                sw.Stop();
                result.DurationMilliseconds = sw.ElapsedMilliseconds;
                progress(totalSteps, totalSteps, "重建完成");
            }
            finally
            {
                // 恢复屏幕刷新
                app.ScreenUpdating = oldScreenUpdating;
            }

            return result;
        }

        // 段落快照：记录要处理的标题信息
        private sealed class OutlineParagraphSnapshot
        {
            public Word.Paragraph Paragraph { get; set; }    // 段落对象
            public int Level { get; set; }                   // 大纲级别
            public int Start { get; set; }                   // 开始位置
            public int End { get; set; }                     // 结束位置
            public int ManualPrefixLength { get; set; }      // 手动编号长度
            public bool HasManualPrefix => ManualPrefixLength > 0; // 是否有手动编号
        }

        private Word.Range GetMainStoryRange(Word.Document doc, out string scope)
        {
            scope = "全文";
            Word.Range mainStory = doc.StoryRanges[Word.WdStoryType.wdMainTextStory];
            return mainStory?.Duplicate ?? doc.Content.Duplicate;
        }

        private void AutoUpdateOutlineListForParagraphs(List<Word.Paragraph> paragraphs)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null || paragraphs == null || paragraphs.Count == 0)
            {
                return;
            }

            List<OutlineParagraphSnapshot> targets = new List<OutlineParagraphSnapshot>();
            foreach (Word.Paragraph paragraph in paragraphs)
            {
                ThrowIfOperationCancelled();
                int level = GetParagraphOutlineLevel(paragraph);
                if (level < 1 || level > 9 || paragraph?.Range == null)
                {
                    continue;
                }

                targets.Add(new OutlineParagraphSnapshot
                {
                    Paragraph = paragraph,
                    Level = level,
                    Start = paragraph.Range.Start,
                    End = paragraph.Range.End,
                    ManualPrefixLength = GetManualHeadingPrefixLength(paragraph.Range.Text)
                });
            }

            targets = targets.OrderBy(item => item.Start).ToList();
            if (targets.Count == 0)
            {
                return;
            }

            OutlineListRebuildOptions options = new OutlineListRebuildOptions
            {
                SelectedLevels = new HashSet<int>(Enumerable.Range(1, 9)),
                ClearManualNumbering = true,
                NumberPattern = outlineNumberPattern,
                Alignment = (int)Word.WdListLevelAlignment.wdListLevelAlignLeft,
                TrailingCharacter = (int)Word.WdTrailingCharacter.wdTrailingNone,
                NumberTextSpacing = outlineNumberTextSpacing
            };

            using (new WordPerformanceScope(app))
            {
                int currentStep = 0;
                int totalSteps = targets.Count + targets.Count(item => item.HasManualPrefix) + 2;
                ClearManualHeadingPrefixes(doc, targets, totalSteps, ref currentStep, (current, total, message) => { });
                ThrowIfOperationCancelled();
                Word.ListTemplate listTemplate = BuildOutlineListTemplate(
                    doc,
                    options,
                    CalculateStartAtByLevel(doc, targets));
                ApplyOutlineListToParagraphs(targets, listTemplate, totalSteps, ref currentStep, (current, total, message) => { });
            }
        }

        private int[] CalculateStartAtByLevel(Word.Document doc, List<OutlineParagraphSnapshot> targets)
        {
            int[] counters = new int[10];
            if (doc?.Content == null || targets == null || targets.Count == 0)
            {
                return counters;
            }

            int firstStart = targets.Min(item => item.Start);
            HashSet<int> levels = new HashSet<int>(Enumerable.Range(1, 9));
            foreach (Word.Range range in CollectOutlineRangesByFind(doc, levels, doc.Content.Start, firstStart))
            {
                ThrowIfOperationCancelled();
                Word.Paragraph paragraph = GetHostParagraph(range);
                int level = GetParagraphOutlineLevel(paragraph);
                if (level < 1 || level > 9)
                {
                    continue;
                }

                counters[level]++;
                for (int deeperLevel = level + 1; deeperLevel <= 9; deeperLevel++)
                {
                    counters[deeperLevel] = 0;
                }
            }

            int firstLevel = targets.OrderBy(item => item.Start).First().Level;
            if (firstLevel >= 1 && firstLevel <= 9)
            {
                counters[firstLevel]++;
            }

            return counters;
        }

        private static int GetParagraphOutlineLevel(Word.Paragraph paragraph)
        {
            try
            {
                Word.WdOutlineLevel level = paragraph.OutlineLevel;
                return level >= Word.WdOutlineLevel.wdOutlineLevel1 && level <= Word.WdOutlineLevel.wdOutlineLevel9
                    ? (int)level
                    : 0;
            }
            catch
            {
                return 0;
            }
        }

        // 收集所有需要处理的标题段落
        private List<OutlineParagraphSnapshot> CollectOutlineParagraphSnapshots(
            Word.Document doc,
            HashSet<int> selectedLevels,
            int? targetStart,
            int? targetEnd,
            bool detectManualPrefixes,
            int totalSteps,
            ref int currentStep,
            Action<int, int, string> progress,
            IReadOnlyList<Word.Range> headingRangesOverride = null)
        {
            List<OutlineParagraphSnapshot> targets = new List<OutlineParagraphSnapshot>();
            HashSet<int> seenStarts = new HashSet<int>();

            progress(0, 0, "正在按大纲级别定位标题...");
            IEnumerable<Word.Range> outlineRanges = headingRangesOverride ??
                CollectOutlineRangesByFind(doc, selectedLevels, targetStart, targetEnd);
            foreach (Word.Range outlineRange in outlineRanges)
            {
                ThrowIfOperationCancelled();
                TryAddOutlineTarget(
                    outlineRange,
                    selectedLevels,
                    targetStart,
                    targetEnd,
                    detectManualPrefixes,
                    seenStarts,
                    targets);
            }

            currentStep = targets.Count;
            progress(currentStep, currentStep, "已定位标题，准备重刷章节号...");
            return targets.OrderBy(item => item.Start).ToList();
        }

        private static List<Word.Range> CollectOutlineRangesByFind(
            Word.Document doc,
            HashSet<int> selectedLevels,
            int? targetStart,
            int? targetEnd,
            Action<int, int, string> progress = null)
        {
            List<Word.Range> ranges = new List<Word.Range>();
            if (doc?.Content == null || selectedLevels == null || selectedLevels.Count == 0)
            {
                return ranges;
            }

            int searchStart = Math.Max(doc.Content.Start, targetStart ?? doc.Content.Start);
            int searchEnd = Math.Min(doc.Content.End, targetEnd ?? doc.Content.End);
            if (searchStart >= searchEnd)
            {
                return ranges;
            }

            HashSet<int> seenStarts = new HashSet<int>();
            List<int> levels = selectedLevels
                .Where(level => level >= 1 && level <= 9)
                .OrderBy(level => level)
                .ToList();
            int completedLevels = 0;
            progress?.Invoke(0, Math.Max(1, levels.Count), "正在定位标题段落...");

            foreach (int level in levels)
            {
                ThrowIfOperationCancelled();
                try
                {
                    int cursor = searchStart;
                    int iterationCount = 0;
                    int maxIterations = Math.Max(1000, searchEnd - searchStart + 1);
                    while (cursor < searchEnd && iterationCount++ < maxIterations)
                    {
                        ThrowIfOperationCancelled();
                        Word.Range searchRange = doc.Range(cursor, searchEnd);
                        Word.Find find = searchRange.Find;
                        find.ClearFormatting();
                        find.Text = "^p";
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.Format = true;
                        find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)level;
                        if (!find.Execute())
                        {
                            break;
                        }

                        Word.Paragraph paragraph = GetHostParagraph(searchRange);
                        Word.Range paragraphRange = paragraph?.Range;
                        if (paragraphRange != null
                            && paragraphRange.Start >= searchStart
                            && paragraphRange.Start < searchEnd
                            && seenStarts.Add(paragraphRange.Start))
                        {
                            ranges.Add(paragraphRange.Duplicate);
                        }

                        int hitStart = searchRange.Start;
                        int hitEnd = searchRange.End;
                        if (hitStart < cursor || hitEnd <= hitStart)
                        {
                            break;
                        }

                        int nextStart = Math.Max(cursor + 1, hitEnd);
                        if (paragraphRange != null)
                        {
                            nextStart = Math.Max(nextStart, paragraphRange.End);
                        }

                        if (nextStart >= searchEnd)
                        {
                            break;
                        }

                        cursor = nextStart;
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch
                {
                    // 某一级标题检索失败时继续检查其余大纲级别。
                }

                completedLevels++;
                progress?.Invoke(
                    completedLevels,
                    Math.Max(1, levels.Count),
                    $"正在定位标题段落，已找到 {ranges.Count} 个...");
            }

            return ranges.OrderBy(item => item.Start).ToList();
        }

        private static int FindDeepestOutlineLevel(Word.Document doc, int? targetStart, int? targetEnd)
        {
            if (doc?.Content == null)
            {
                return 0;
            }

            int searchStart = Math.Max(doc.Content.Start, targetStart ?? doc.Content.Start);
            int searchEnd = Math.Min(doc.Content.End, targetEnd ?? doc.Content.End);
            if (searchStart >= searchEnd)
            {
                return 0;
            }

            for (int level = 9; level >= 1; level--)
            {
                ThrowIfOperationCancelled();
                try
                {
                    Word.Range searchRange = doc.Range(searchStart, searchEnd);
                    Word.Find find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = "^p";
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = true;
                    find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)level;
                    if (find.Execute())
                    {
                        return level;
                    }
                }
                catch
                {
                }
            }

            return 0;
        }

        private static void TryAddOutlineTarget(
            Word.Range headingRange,
            HashSet<int> selectedLevels,
            int? targetStart,
            int? targetEnd,
            bool detectManualPrefixes,
            HashSet<int> seenStarts,
            List<OutlineParagraphSnapshot> targets)
        {
            try
            {
                ThrowIfOperationCancelled();
                Word.Paragraph paragraph = GetHostParagraph(headingRange);
                int outlineLevel = GetParagraphOutlineLevel(paragraph);
                if (!selectedLevels.Contains(outlineLevel))
                {
                    return;
                }

                Word.Range range = paragraph.Range;
                int start = range.Start;
                if ((targetStart.HasValue && start < targetStart.Value)
                    || (targetEnd.HasValue && start >= targetEnd.Value)
                    || !seenStarts.Add(start))
                {
                    return;
                }

                int manualPrefixLength = detectManualPrefixes
                    ? GetManualHeadingPrefixLength(range.Text)
                    : 0;

                targets.Add(new OutlineParagraphSnapshot
                {
                    Paragraph = paragraph,
                    Level = outlineLevel,
                    Start = start,
                    End = range.End,
                    ManualPrefixLength = manualPrefixLength
                });
            }
            catch
            {
            }
        }

        // 清理段落前的手动编号
        private int ClearManualHeadingPrefixes(
            Word.Document doc,
            List<OutlineParagraphSnapshot> targets,
            int totalSteps,
            ref int currentStep,
            Action<int, int, string> progress)
        {
            if (doc == null || targets == null || targets.Count == 0)
            {
                return 0;
            }

            int clearedCount = 0;
            // 倒序删除，避免位置错乱
            List<OutlineParagraphSnapshot> manualTargets = targets
                .Where(item => item != null && item.HasManualPrefix)
                .OrderByDescending(item => item.Start)
                .ToList();

            if (manualTargets.Count == 0)
            {
                return 0;
            }

            int progressInterval = manualTargets.Count > 5000 ? 500 : 200;
            for (int i = 0; i < manualTargets.Count; i++)
            {
                ThrowIfOperationCancelled();
                OutlineParagraphSnapshot snapshot = manualTargets[i];
                if (snapshot.ManualPrefixLength <= 0)
                {
                    continue;
                }

                try
                {
                    // 删除手动编号部分
                    int deleteEnd = snapshot.Start + snapshot.ManualPrefixLength;
                    if (deleteEnd > snapshot.Start && deleteEnd <= snapshot.End)
                    {
                        doc.Range(snapshot.Start, deleteEnd).Text = string.Empty;
                        clearedCount++;
                    }
                }
                catch
                {
                }

                currentStep++;
                // 更新进度
                if ((i + 1) % progressInterval == 0 || i == manualTargets.Count - 1)
                {
                    progress(currentStep, totalSteps, "正在清理手工章节号...");
                }
            }
            return clearedCount;
        }

        // 获取手动编号的长度（用正则匹配）
        private static int GetManualHeadingPrefixLength(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            // 清理特殊字符
            string normalized = text.Replace("\r", string.Empty).Replace("\a", string.Empty);

            // 匹配正则，返回长度
            foreach (Regex pattern in ManualHeadingPrefixPatterns)
            {
                ThrowIfOperationCancelled();
                Match match = pattern.Match(normalized);
                if (match.Success && match.Index == 0 && match.Length > 0)
                {
                    return match.Length;
                }
            }
            return 0;
        }

        private static int GetArabicManualHeadingPrefixLength(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            string normalized = text.Replace("\r", string.Empty).Replace("\a", string.Empty);
            Match match = ArabicManualHeadingPrefixPattern.Match(normalized);
            return match.Success && match.Index == 0 ? match.Length : 0;
        }

        // 给段落应用多级列表
        private int ApplyOutlineListToParagraphs(
            List<OutlineParagraphSnapshot> targets,
            Word.ListTemplate listTemplate,
            int totalSteps,
            ref int currentStep,
            Action<int, int, string> progress,
            bool continueFirstList = false)
        {
            int count = 0;
            int progressInterval = targets.Count > 5000 ? 500 : 200;

            foreach (OutlineParagraphSnapshot snapshot in targets)
            {
                ThrowIfOperationCancelled();
                try
                {
                    Word.Range paragraphRange = snapshot?.Paragraph?.Range;
                    if (paragraphRange == null)
                    {
                        continue;
                    }

                    TryRemoveExistingListFormatting(paragraphRange);
                    paragraphRange.ParagraphFormat.LeftIndent = 0f;
                    paragraphRange.ParagraphFormat.FirstLineIndent = 0f;
                }
                catch
                {
                }
            }

            for (int i = 0; i < targets.Count; i++)
            {
                ThrowIfOperationCancelled();
                try
                {
                    OutlineParagraphSnapshot snapshot = targets[i];
                    Word.Paragraph paragraph = snapshot.Paragraph;
                    if (paragraph == null || paragraph.Range == null)
                    {
                        continue;
                    }

                    Word.Range paragraphRange = paragraph.Range;

                    bool continuePreviousList = count > 0 || continueFirstList;

                    // 应用列表模板
                    paragraphRange.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate: listTemplate,
                        ContinuePreviousList: continuePreviousList,
                        ApplyTo: Word.WdListApplyTo.wdListApplyToSelection,
                        DefaultListBehavior: Word.WdDefaultListBehavior.wdWord10ListBehavior,
                        ApplyLevel: snapshot.Level);

                    count++;
                }
                catch
                {
                }

                currentStep++;
                // 更新进度
                if ((i + 1) % progressInterval == 0 || i == targets.Count - 1)
                {
                    progress(currentStep, totalSteps, "正在应用新的章节号...");
                }
            }
            return count;
        }

        // 移除段落已有的自动列表格式，避免被旧列表状态影响重新编号。
        private static void TryRemoveExistingListFormatting(Word.Range paragraphRange)
        {
            if (paragraphRange == null)
            {
                return;
            }

            try
            {
                paragraphRange.ListFormat.RemoveNumbers();
            }
            catch
            {
            }
        }

        // 创建多级列表模板（1-9级）
        private Word.ListTemplate BuildOutlineListTemplate(
            Word.Document doc,
            OutlineListRebuildOptions options)
        {
            return BuildOutlineListTemplate(doc, options, null);
        }

        private Word.ListTemplate BuildOutlineListTemplate(
            Word.Document doc,
            OutlineListRebuildOptions options,
            int[] startAtByLevel)
        {
            Word.ListTemplate listTemplate = doc.ListTemplates.Add(OutlineNumbered: true);

            for (int level = 1; level <= 9; level++)
            {
                ThrowIfOperationCancelled();
                Word.ListLevel listLevel = listTemplate.ListLevels[level];
                listLevel.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                // 生成编号格式（1/1.1/1.1.1）
                listLevel.NumberFormat = BuildNumberFormat(level, options.NumberPattern, options.NumberTextSpacing);
                listLevel.TrailingCharacter = (Word.WdTrailingCharacter)options.TrailingCharacter;
                listLevel.Alignment = (Word.WdListLevelAlignment)options.Alignment;
                listLevel.NumberPosition = 0f;
                listLevel.TextPosition = 0f;
                listLevel.TabPosition = 0f;
                listLevel.StartAt = startAtByLevel != null && level < startAtByLevel.Length && startAtByLevel[level] > 0
                    ? startAtByLevel[level]
                    : 1;
                listLevel.ResetOnHigher = level > 1 ? level - 1 : 0;
                ApplyListLevelFont(listLevel, level, GetStyleDefinitionForOutlineLevel(level));
                listLevel.LinkedStyle = string.Empty;
            }
            return listTemplate;
        }

        private static void ApplyListLevelFont(
            Word.ListLevel listLevel,
            int level,
            StyleDefinitionRequest definition)
        {
            if (listLevel == null || definition == null)
            {
                return;
            }

            string fontName = level == 1 ? "黑体" : "宋体";
            float fontSize = definition.ListFontSize > 0f
                ? definition.ListFontSize
                : definition.FontSize;

            try
            {
                if (!string.IsNullOrWhiteSpace(fontName))
                {
                    listLevel.Font.Name = fontName;
                    listLevel.Font.NameFarEast = fontName;
                }

                if (fontSize > 0f)
                {
                    listLevel.Font.Size = fontSize;
                }

                listLevel.Font.Bold = 0;
                listLevel.Font.Color = Word.WdColor.wdColorAutomatic;
            }
            catch
            {
            }
        }

        // 生成编号格式：如 level=3 → 1.1.1 或 (1.1.1)
        private string BuildNumberFormat(int level, OutlineNumberPattern pattern, int spaces)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i <= level; i++)
            {
                if (i > 1)
                {
                    sb.Append('.');
                }
                sb.Append('%').Append(i);
            }

            // 括号格式
            if (pattern == OutlineNumberPattern.Parenthesized)
            {
                return AppendNumberSpacing("(" + sb + ")", spaces);
            }
            if (pattern == OutlineNumberPattern.Dotted)
            {
                return AppendNumberSpacing(sb + ".", spaces);
            }
            // 普通数字格式
            return AppendNumberSpacing(sb.ToString(), spaces);
        }

        private static string AppendNumberSpacing(string numberFormat, int spaces)
        {
            return (numberFormat ?? string.Empty) + new string(' ', Math.Max(0, Math.Min(2, spaces)));
        }
    }
}

