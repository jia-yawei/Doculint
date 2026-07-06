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

        // 点击【更新全部章节号】按钮：一键重建多级列表
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
                "本功能将检查全文中大纲级别为 1-6 级的标题，并重建自动章节号。\r\n是否继续？",
                "更新全部章节号",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            if (confirmResult != DialogResult.OK)
            {
                return;
            }

            // 默认设置：处理常用样式库中的1-6级标题。
            HashSet<int> selectedLevels = new HashSet<int> { 1, 2, 3, 4, 5, 6 };
            OutlineListRebuildOptions defaultOptions = new OutlineListRebuildOptions
            {
                SelectedLevels = selectedLevels,
                ClearManualNumbering = true,
                NumberPattern = commonStyleSettings.NumberPattern,
                Alignment = (int)Word.WdListLevelAlignment.wdListLevelAlignLeft,
                TrailingCharacter = (int)Word.WdTrailingCharacter.wdTrailingSpace
            };

            try
            {
                using (OutlineProgressForm progressForm = new OutlineProgressForm())
                {
                    progressForm.Show();
                    progressForm.ReportProgress(0, 1, "正在准备文档...");

                    using (new WordPerformanceScope(app))
                    {
                        RebuildOutlineListWithOptions(app, doc, defaultOptions, (current, total, message) =>
                        {
                            progressForm.ReportProgress(current, total, message);
                        });
                    }
                }

                MessageBox.Show("章节号更新完成。", "更新全部章节号");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理时出现错误: {ex.Message}", "错误");
            }
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
                Word.Paragraphs scanParagraphs = scanRange.Paragraphs;
                int totalSteps = Math.Max(1, scanParagraphs.Count);
                int currentStep = 0;

                progress(0, totalSteps, "正在按大纲级别查找目标段落...");
                result.ScanScope = scanScope;

                // 1. 扫描符合条件的标题段落
                phaseWatch.Restart();
                List<OutlineParagraphSnapshot> targets = CollectOutlineParagraphSnapshots(
                    scanParagraphs,
                    selectedLevels,
                    scanRange.Start,
                    scanRange.End,
                    options.ClearManualNumbering,
                    totalSteps,
                    ref currentStep,
                    progress);
                phaseWatch.Stop();
                result.ScanMilliseconds = phaseWatch.ElapsedMilliseconds;
                result.TargetParagraphCount = targets.Count;

                // 没找到标题则退出
                if (targets.Count == 0)
                {
                    throw new InvalidOperationException("未找到 1-6 级大纲标题，请先设置段落的大纲级别。");
                }

                // 2. 清理手动编号
                int manualPrefixCount = targets.Count(item => item.HasManualPrefix);
                totalSteps = scanParagraphs.Count + manualPrefixCount + targets.Count + 2;
                result.ClearedManualNumberCount = 0;
                result.CleanupMilliseconds = 0;

                phaseWatch.Restart();
                if (options.ClearManualNumbering && manualPrefixCount > 0)
                {
                    progress(currentStep, totalSteps, "正在清理手工章节号...");
                    result.ClearedManualNumberCount = ClearManualHeadingPrefixes(
                        doc,
                        targets,
                        totalSteps,
                        ref currentStep,
                        progress);
                }
                phaseWatch.Stop();
                result.CleanupMilliseconds = phaseWatch.ElapsedMilliseconds;

                // 3. 创建多级列表模板
                progress(++currentStep, totalSteps, "正在建立新的多级列表模板...");
                Word.ListTemplate listTemplate = BuildOutlineListTemplate(doc, options);

                // 4. 给段落应用列表
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
                int level = GetParagraphOutlineLevel(paragraph);
                if (level < 1 || level > 6 || paragraph?.Range == null)
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
                SelectedLevels = new HashSet<int> { 1, 2, 3, 4, 5, 6 },
                ClearManualNumbering = true,
                NumberPattern = commonStyleSettings.NumberPattern,
                Alignment = (int)Word.WdListLevelAlignment.wdListLevelAlignLeft,
                TrailingCharacter = (int)Word.WdTrailingCharacter.wdTrailingSpace
            };

            using (new WordPerformanceScope(app))
            {
                int currentStep = 0;
                int totalSteps = targets.Count + targets.Count(item => item.HasManualPrefix) + 2;
                ClearManualHeadingPrefixes(doc, targets, totalSteps, ref currentStep, (current, total, message) => { });
                bool continueFromPrevious = TryFindPreviousOutlineListTemplate(doc, targets[0].Start, out Word.ListTemplate listTemplate);
                if (!continueFromPrevious)
                {
                    listTemplate = BuildOutlineListTemplate(doc, options);
                }

                ApplyOutlineListToParagraphs(targets, listTemplate, totalSteps, ref currentStep, (current, total, message) => { }, continueFromPrevious);
            }
        }

        private bool TryFindPreviousOutlineListTemplate(
            Word.Document doc,
            int beforePosition,
            out Word.ListTemplate listTemplate)
        {
            listTemplate = null;
            if (doc == null || beforePosition <= doc.Content.Start)
            {
                return false;
            }

            Word.Range probeRange = doc.Range(doc.Content.Start, Math.Min(beforePosition, doc.Content.End));
            Word.Paragraphs paragraphs = probeRange.Paragraphs;
            for (int i = paragraphs.Count; i >= 1; i--)
            {
                Word.Paragraph paragraph = paragraphs[i];
                int level = GetParagraphOutlineLevel(paragraph);
                if (level < 1 || level > 6)
                {
                    continue;
                }

                try
                {
                    Word.ListFormat listFormat = paragraph.Range.ListFormat;
                    if (listFormat != null
                        && listFormat.ListType != Word.WdListType.wdListNoNumbering
                        && listFormat.ListTemplate != null)
                    {
                        listTemplate = listFormat.ListTemplate;
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
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
            Word.Paragraphs paragraphs,
            HashSet<int> selectedLevels,
            int? targetStart,
            int? targetEnd,
            bool detectManualPrefixes,
            int totalSteps,
            ref int currentStep,
            Action<int, int, string> progress)
        {
            List<OutlineParagraphSnapshot> targets = new List<OutlineParagraphSnapshot>();
            int paragraphCount = paragraphs.Count;
            int progressInterval = paragraphCount > 20000 ? 2000 : paragraphCount > 5000 ? 800 : 300;

            foreach (Word.Paragraph paragraph in paragraphs)
            {
                try
                {
                    Word.WdOutlineLevel level = paragraph.OutlineLevel;
                    // 判断是否是1-9级大纲标题
                    if (level >= Word.WdOutlineLevel.wdOutlineLevel1 && level <= Word.WdOutlineLevel.wdOutlineLevel9)
                    {
                        int outlineLevel = (int)level;
                        // 是否在用户选择的级别里
                        if (selectedLevels.Contains(outlineLevel))
                        {
                            Word.Range range = paragraph.Range;
                            int start = range.Start;
                            // 判断是否在目标范围内
                            if ((!targetStart.HasValue || start >= targetStart.Value) && (!targetEnd.HasValue || start < targetEnd.Value))
                            {
                                // 检测是否有手动编号
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
                        }
                    }
                }
                catch
                {
                    // 出错跳过
                }

                currentStep++;
                // 进度更新
                if (currentStep == 1 || currentStep % progressInterval == 0 || currentStep >= paragraphCount)
                {
                    progress(currentStep, totalSteps, "正在扫描目标级别段落...");
                }
            }
            return targets;
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
                Match match = pattern.Match(normalized);
                if (match.Success && match.Index == 0 && match.Length > 0)
                {
                    return match.Length;
                }
            }
            return 0;
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

            for (int i = 0; i < targets.Count; i++)
            {
                try
                {
                    OutlineParagraphSnapshot snapshot = targets[i];
                    Word.Paragraph paragraph = snapshot.Paragraph;
                    if (paragraph == null || paragraph.Range == null)
                    {
                        continue;
                    }

                    Word.Range paragraphRange = paragraph.Range;

                    // 先去掉旧的自动编号，避免旧列表状态干扰新的续接关系。
                    TryRemoveExistingListFormatting(paragraphRange);

                    // 清除缩进
                    paragraphRange.ParagraphFormat.LeftIndent = 0f;
                    paragraphRange.ParagraphFormat.FirstLineIndent = 0f;

                    bool continuePreviousList = count > 0 || continueFirstList;

                    // 应用列表模板
                    paragraphRange.ListFormat.ApplyListTemplateWithLevel(
                        ListTemplate: listTemplate,
                        ContinuePreviousList: continuePreviousList,
                        ApplyTo: Word.WdListApplyTo.wdListApplyToSelection,
                        DefaultListBehavior: Word.WdDefaultListBehavior.wdWord10ListBehavior,
                        ApplyLevel: snapshot.Level);

                    // 某些宿主对 ApplyLevel 不稳定，这里再显式设置一次级别，确保层级跟标题级别一致。
                    paragraphRange.ListFormat.ListLevelNumber = snapshot.Level;

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
            Word.ListTemplate listTemplate = doc.ListTemplates.Add(OutlineNumbered: true);

            for (int level = 1; level <= 9; level++)
            {
                Word.ListLevel listLevel = listTemplate.ListLevels[level];
                listLevel.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                // 生成编号格式（1/1.1/1.1.1）
                listLevel.NumberFormat = BuildNumberFormat(level, options.NumberPattern);
                listLevel.TrailingCharacter = (Word.WdTrailingCharacter)options.TrailingCharacter;
                listLevel.Alignment = (Word.WdListLevelAlignment)options.Alignment;
                listLevel.NumberPosition = 0f;
                listLevel.TextPosition = 0f;
                listLevel.TabPosition = 0f;
                listLevel.StartAt = 1;
                listLevel.ResetOnHigher = level > 1 ? level - 1 : 0;
                listLevel.LinkedStyle = string.Empty;
            }
            return listTemplate;
        }

        // 生成编号格式：如 level=3 → 1.1.1 或 (1.1.1)
        private string BuildNumberFormat(int level, OutlineNumberPattern pattern)
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
                return "(" + sb + ")";
            }
            // 普通数字格式
            return sb.ToString();
        }
    }
}
