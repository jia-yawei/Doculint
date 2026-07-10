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
                ResetOperationCancellation();
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
                ResetOperationCancellation();
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
                LinkedStyles = new Dictionary<int, string>(),
                ScanScope = "当前选区"
            };

            Stopwatch sw = Stopwatch.StartNew();
            ThrowIfOperationCancelled();
            ConfigureLinkedStyleOutlineLevels(doc);
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
                LinkedStyles = new Dictionary<int, string>(),
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
            Action<int, int, string> progress)
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
                LinkedStyles = new Dictionary<int, string>()
            };

            Stopwatch sw = Stopwatch.StartNew();
            // 关闭屏幕刷新，提高速度
            bool oldScreenUpdating = app.ScreenUpdating;
            app.ScreenUpdating = false;

            try
            {
                Stopwatch phaseWatch = new Stopwatch();

                Word.Range scanRange = GetMainStoryRange(doc, out string scanScope);
                int currentStep = 0;

                ConfigureLinkedStyleOutlineLevels(doc);
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
                    progress);
                phaseWatch.Stop();
                result.ScanMilliseconds = phaseWatch.ElapsedMilliseconds;
                targets = targets.OrderBy(item => item.Start).ToList();
                result.TargetParagraphCount = targets.Count;

                // 没找到标题则退出
                if (targets.Count == 0)
                {
                    string levelsText = string.Join("、", selectedLevels.OrderBy(level => level));
                    throw new InvalidOperationException($"未找到 {levelsText} 级大纲标题。请确认段落已设置为对应大纲级别，或先在样式链接中将样式绑定到大纲级别。");
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
            Action<int, int, string> progress)
        {
            List<OutlineParagraphSnapshot> targets = new List<OutlineParagraphSnapshot>();
            HashSet<int> seenStarts = new HashSet<int>();
            List<Word.Range> headingRanges = CollectHeadingRangesByGoTo(doc, includeBlankHeadings: false);
            int headingCount = headingRanges.Count;
            int progressInterval = headingCount > 1000 ? 100 : 20;

            foreach (Word.Range headingRange in headingRanges)
            {
                ThrowIfOperationCancelled();
                TryAddOutlineTarget(
                    headingRange,
                    selectedLevels,
                    targetStart,
                    targetEnd,
                    detectManualPrefixes,
                    seenStarts,
                    targets);

                currentStep++;
                // 进度更新
                if (currentStep == 1 || currentStep % progressInterval == 0 || currentStep >= headingCount)
                {
                    progress(currentStep, headingCount, "正在扫描标题...");
                }
            }

            foreach (Word.Range outlineRange in CollectOutlineRangesByFind(doc, selectedLevels, targetStart, targetEnd))
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

            return targets.OrderBy(item => item.Start).ToList();
        }

        private static List<Word.Range> CollectOutlineRangesByFind(
            Word.Document doc,
            HashSet<int> selectedLevels,
            int? targetStart,
            int? targetEnd)
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
            foreach (int level in selectedLevels.OrderBy(item => item))
            {
                ThrowIfOperationCancelled();
                Word.Range searchRange = doc.Range(searchStart, searchEnd);
                Word.Find find = searchRange.Find;
                find.ClearFormatting();
                find.Text = string.Empty;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = true;
                find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)level;

                while (find.Execute())
                {
                    ThrowIfOperationCancelled();
                    Word.Paragraph paragraph = GetHostParagraph(searchRange);
                    Word.Range paragraphRange = paragraph?.Range;
                    if (paragraphRange != null && seenStarts.Add(paragraphRange.Start))
                    {
                        ranges.Add(paragraphRange.Duplicate);
                    }

                    int nextStart = Math.Max(searchRange.End, searchRange.Start + 1);
                    if (nextStart >= searchEnd)
                    {
                        break;
                    }

                    searchRange.SetRange(nextStart, searchEnd);
                    find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = string.Empty;
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = true;
                    find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)level;
                }
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
                    find.Text = string.Empty;
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
                ApplyListLevelFont(listLevel, GetStyleDefinitionForOutlineLevel(level));
                string linkedStyle = GetLinkedStyleForOutlineLevel(level);
                if (!string.IsNullOrWhiteSpace(linkedStyle))
                {
                    SetStyleOutlineLevel(doc, linkedStyle, level);
                    listLevel.LinkedStyle = linkedStyle;
                }
                else
                {
                    listLevel.LinkedStyle = string.Empty;
                }
            }
            return listTemplate;
        }

        private static void ApplyListLevelFont(Word.ListLevel listLevel, StyleDefinitionRequest definition)
        {
            if (listLevel == null || definition == null)
            {
                return;
            }

            string fontName = string.IsNullOrWhiteSpace(definition.ListFontName)
                ? definition.FontName
                : definition.ListFontName;
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

