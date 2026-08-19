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
        private void btnStartDocumentCheck_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteDocumentCheck();
        }

        private void btnSoftwareDocumentCheck_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteSoftwareDocumentCheck();
        }

        private void ExecuteDocumentCheck()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            try
            {
                ResetOperationCancellation("文档检查");
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档检查");
                    return;
                }

                bool checkBlankLines = chkNonBodyBlankLine.Checked;
                bool checkCaptions = chkCaptionContinuity.Checked;
                bool checkLists = chkListContinuity.Checked;
                bool checkStyles = chkStyleConsistency.Checked;
                bool checkBrokenReferences = chkBrokenReferences.Checked;
                if (!checkBlankLines && !checkCaptions && !checkLists && !checkStyles && !checkBrokenReferences)
                {
                    MessageBox.Show("请至少选择一个检查项。", "文档检查");
                    return;
                }

                int scanStart = GetDocumentCheckScanStart(doc);
                int totalSteps = CountDocumentCheckSteps(doc, scanStart, checkBlankLines, checkCaptions, checkLists, checkStyles, checkBrokenReferences);
                DocumentCheckProgress progress = new DocumentCheckProgress(app, totalSteps);
                progress.Report("正在检查文档（已忽略首页）");

                List<DocumentCheckIssue> issues;
                using (new WordPerformanceScope(app))
                {
                    issues = CollectDocumentCheckIssues(doc, scanStart, checkBlankLines, checkCaptions, checkLists, checkStyles, checkBrokenReferences, progress);
                }

                progress.Complete(issues.Count);
                Globals.ThisAddIn.ShowDocumentCheckResultPane(doc, ToNavigationEntries(issues));
            }
            catch (OperationCanceledException)
            {
                TrySetStatusBar(app, "文档检查已停止");
                MessageBox.Show("文档检查已停止。", "文档检查");
            }
            catch (Exception ex)
            {
                TrySetStatusBar(app, "文档检查失败");
                MessageBox.Show($"文档检查失败: {ex.Message}", "文档检查");
            }
        }

        private void ExecuteSoftwareDocumentCheck()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            try
            {
                ResetOperationCancellation("软件文档检查");
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档检查");
                    return;
                }

                TrySetStatusBar(app, "软件文档检查：正在识别文档类型...");
                List<DocumentCheckIssue> issues;
                using (new WordPerformanceScope(app))
                {
                    issues = CollectSoftwareDocumentIssues(doc);
                }

                TrySetStatusBar(app, issues.Count > 0
                    ? $"软件文档检查完成：发现 {issues.Count} 个问题"
                    : "软件文档检查完成：未发现问题");
                new SoftwareDocumentCheckResultForm(doc, issues).Show();
            }
            catch (OperationCanceledException)
            {
                TrySetStatusBar(app, "软件文档检查已停止");
                MessageBox.Show("软件文档检查已停止。", "文档检查");
            }
            catch (Exception ex)
            {
                TrySetStatusBar(app, "软件文档检查失败");
                MessageBox.Show($"软件文档检查失败: {ex.Message}", "文档检查");
            }
        }

        private static List<DocumentCheckIssue> CollectDocumentCheckIssues(
            Word.Document doc,
            int scanStart,
            bool checkBlankLines,
            bool checkCaptions,
            bool checkLists,
            bool checkStyles,
            bool checkBrokenReferences,
            DocumentCheckProgress progress)
        {
            List<DocumentCheckIssue> issues = new List<DocumentCheckIssue>();
            Dictionary<string, int> expectedNumbers = new Dictionary<string, int>
            {
                { "图", 1 },
                { "表", 1 }
            };
            int[] counters = new int[10];

            if (checkBlankLines)
            {
                ThrowIfOperationCancelled();
                issues.AddRange(CollectNonBodyBlankLineIssuesByOutlineFind(doc, scanStart, progress));
            }

            if (!checkCaptions && !checkLists && !checkStyles && !checkBrokenReferences)
            {
                return issues;
            }

            if (checkCaptions)
            {
                ThrowIfOperationCancelled();
                issues.AddRange(CollectCaptionContinuityIssues(doc, scanStart, expectedNumbers, progress));
            }

            if (checkLists)
            {
                ThrowIfOperationCancelled();
                issues.AddRange(CollectListContinuityIssuesByOutlineFind(doc, scanStart, counters, progress));
            }

            if (checkStyles)
            {
                ThrowIfOperationCancelled();
                issues.AddRange(CollectStyleConsistencyIssues(doc, scanStart, progress));
            }

            if (checkBrokenReferences)
            {
                ThrowIfOperationCancelled();
                issues.AddRange(CollectBrokenReferenceIssues(doc, scanStart, progress));
            }

            return issues;
        }

        private static List<DocumentCheckIssue> CollectSoftwareDocumentIssues(Word.Document doc)
        {
            List<DocumentCheckIssue> issues = new List<DocumentCheckIssue>();
            SoftwareDocumentSpec spec = DetectSoftwareDocumentSpec(doc);
            int scanStart = GetDocumentCheckScanStart(doc);
            ThrowIfOperationCancelled();
            if (spec == null)
            {
                issues.Add(new DocumentCheckIssue(
                    scanStart,
                    scanStart,
                    "软件文档检查：未能从当前文档标题识别为已支持的软件文档类型。"));
                return issues;
            }

            List<SoftwareDocumentHeading> headings = CollectSoftwareDocumentHeadings(doc, maxOutlineLevel: 3, scanStart);
            foreach (string requiredTitle in spec.RequiredHeadings)
            {
                ThrowIfOperationCancelled();
                string normalized = NormalizeSoftwareHeading(requiredTitle);
                if (string.IsNullOrWhiteSpace(normalized) || SoftwareHeadingExists(headings, normalized))
                {
                    continue;
                }

                issues.Add(new DocumentCheckIssue(
                    scanStart,
                    scanStart,
                    $"软件文档检查：{spec.Name} 缺少章节标题“{requiredTitle}”。"));
            }

            return issues;
        }

        private static List<DocumentCheckIssue> CollectBrokenReferenceIssues(Word.Document doc, int scanStart, DocumentCheckProgress progress)
        {
            progress.Step("正在检查未更新域");
            ThrowIfOperationCancelled();
            return CollectBrokenReferenceEntries(doc)
                .Where(entry => entry != null && entry.Start >= scanStart)
                .Select(entry => new DocumentCheckIssue(entry.Start, entry.Start, entry.Text ?? "未更新域"))
                .ToList();
        }

        private static SoftwareDocumentSpec DetectSoftwareDocumentSpec(Word.Document doc)
        {
            string titleText = NormalizeSoftwareHeading(string.Join(" ", CollectDocumentTitleCandidates(doc)));
            return SoftwareDocumentSpecs.FirstOrDefault(spec =>
                spec.Aliases.Any(alias => titleText.IndexOf(NormalizeSoftwareHeading(alias), StringComparison.OrdinalIgnoreCase) >= 0));
        }

        private static IEnumerable<string> CollectDocumentTitleCandidates(Word.Document doc)
        {
            yield return doc?.Name ?? string.Empty;

            string propertyTitle = TryGetDocumentProperty(doc, "Title");
            if (!string.IsNullOrWhiteSpace(propertyTitle))
            {
                yield return propertyTitle;
            }

            int count = Math.Min(GetParagraphCount(doc), 20);
            for (int i = 1; i <= count; i++)
            {
                ThrowIfOperationCancelled();
                string text = GetParagraphText(doc, i);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    yield return text;
                }
            }
        }

        private static string TryGetDocumentProperty(Word.Document doc, string propertyName)
        {
            try
            {
                object properties = doc?.BuiltInDocumentProperties;
                object property = properties?.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, properties, new object[] { propertyName });
                object value = property?.GetType().InvokeMember("Value", System.Reflection.BindingFlags.GetProperty, null, property, null);
                return value?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static List<SoftwareDocumentHeading> CollectSoftwareDocumentHeadings(Word.Document doc, int maxOutlineLevel, int scanStart)
        {
            List<SoftwareDocumentHeading> headings = new List<SoftwareDocumentHeading>();
            if (doc?.Content == null)
            {
                return headings;
            }

            HashSet<int> seenStarts = new HashSet<int>();
            int highestLevel = Math.Max(1, Math.Min(9, maxOutlineLevel));
            for (int level = 1; level <= highestLevel; level++)
            {
                ThrowIfOperationCancelled();
                Word.Range searchRange = doc.Content.Duplicate;
                searchRange.SetRange(scanStart, doc.Content.End);
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
                    Word.Range found = searchRange.Duplicate;
                    if (!IsInTable(found) && seenStarts.Add(found.Start))
                    {
                        string text = CleanParagraphText(found.Text);
                        string title = StripHeadingNumber(text);
                        if (!string.IsNullOrWhiteSpace(title))
                        {
                            headings.Add(new SoftwareDocumentHeading(title, NormalizeSoftwareHeading(title), level, found.Start));
                        }
                    }

                    int nextStart = Math.Max(found.End, searchRange.Start + 1);
                    if (nextStart >= doc.Content.End)
                    {
                        break;
                    }

                    searchRange.SetRange(nextStart, doc.Content.End);
                    find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = string.Empty;
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = true;
                    find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)level;
                }
            }

            return headings.OrderBy(item => item.Start).ToList();
        }

        private static bool SoftwareHeadingExists(IEnumerable<SoftwareDocumentHeading> headings, string requiredNormalizedTitle)
        {
            foreach (SoftwareDocumentHeading heading in headings)
            {
                string actual = heading?.NormalizedTitle ?? string.Empty;
                if (string.Equals(actual, requiredNormalizedTitle, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (requiredNormalizedTitle.Length >= 4
                    && actual.IndexOf(requiredNormalizedTitle, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string StripHeadingNumber(string text)
        {
            string value = NormalizeDigits(text ?? string.Empty).Trim();
            value = Regex.Replace(value, @"^\d+(\.\d+)*\s*", string.Empty);
            return value.Trim();
        }

        private static string NormalizeSoftwareHeading(string text)
        {
            string value = StripHeadingNumber(text);
            value = value.Replace("文裆", "文档");
            value = value.Replace("影晌", "影响");
            value = value.Replace("其它", "其他");
            value = Regex.Replace(value, @"[（(][^）)]*[）)]", string.Empty);
            value = value.Replace("“", string.Empty).Replace("”", string.Empty).Replace("\"", string.Empty);
            value = Regex.Replace(value, @"\s+", string.Empty);
            return value.Trim('：', ':', '。', '.', '；', ';');
        }

        private static int GetParagraphCount(Word.Document doc)
        {
            try
            {
                return doc?.Paragraphs?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static string GetParagraphText(Word.Document doc, int index)
        {
            try
            {
                return CleanParagraphText(doc.Paragraphs[index].Range.Text);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int GetDocumentStart(Word.Document doc)
        {
            try
            {
                return doc?.Content?.Start ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static int GetDocumentCheckScanStart(Word.Document doc)
        {
            int documentStart = GetDocumentStart(doc);
            try
            {
                return Math.Max(documentStart, GetFirstPageEnd(doc));
            }
            catch
            {
                return documentStart;
            }
        }

        private static List<DocumentCheckIssue> CollectCaptionContinuityIssues(
            Word.Document doc,
            int scanStart,
            Dictionary<string, int> expectedNumbers,
            DocumentCheckProgress progress)
        {
            List<DocumentCheckIssue> issues = new List<DocumentCheckIssue>();
            foreach (CaptionListEntry entry in CollectCaptionListEntries(doc).Where(item => item != null && item.Start >= scanStart))
            {
                ThrowIfOperationCancelled();
                progress.Step("正在检查题注");
                Word.Range entryRange = doc.Range(entry.Start, Math.Min(doc.Content.End, entry.Start + Math.Max(1, entry.Text?.Length ?? 1)));
                if (IsInTable(entryRange))
                {
                    continue;
                }

                string text = CleanParagraphText(entry.Text);
                Match match = CaptionNumberRegex.Match(text);
                if (!match.Success)
                {
                    continue;
                }

                string kind = match.Groups["kind"].Value;
                int actual = ParseCaptionNumber(match.Groups["num"].Value);
                int expected = expectedNumbers[kind];
                if (actual != expected)
                {
                    issues.Add(CreateIssue(entryRange, $"题注连续性：{kind}题注编号应为 {kind}{expected}，当前为 {kind}{actual}。"));
                }

                expectedNumbers[kind] = expected + 1;
            }

            return issues;
        }

        private static List<DocumentCheckIssue> CollectListContinuityIssuesByOutlineFind(
            Word.Document doc,
            int scanStart,
            int[] counters,
            DocumentCheckProgress progress)
        {
            List<DocumentCheckIssue> issues = new List<DocumentCheckIssue>();
            List<Word.Range> headings = CollectOutlineRangesByFind(doc, scanStart, progress);
            foreach (Word.Range range in headings.OrderBy(item => item.Start))
            {
                ThrowIfOperationCancelled();
                progress.Step("正在检查多级列表");
                Word.Paragraph paragraph = GetHostParagraph(range);
                int level = GetDocumentCheckOutlineLevel(paragraph);
                if (level < 1 || level > 9)
                {
                    continue;
                }

                string listString = GetListString(paragraph);
                int[] actualNumbers = ParseDottedNumber(listString);
                if (actualNumbers.Length == 0)
                {
                    issues.Add(CreateIssue(range, "多级列表连续性：该标题没有自动多级列表编号。"));
                    continue;
                }

                int[] expectedListNumbers = BuildExpectedNumbers(counters, level);
                string expectedText = string.Join(".", expectedListNumbers);
                string actualText = string.Join(".", actualNumbers.Take(level));
                if (actualNumbers.Length < level || !string.Equals(actualText, expectedText, StringComparison.Ordinal))
                {
                    issues.Add(CreateIssue(range, $"多级列表连续性：该标题编号应为 {expectedText}，当前为 {listString}。"));
                }

                ApplyActualCounters(counters, actualNumbers, level);
            }

            return issues;
        }

        private static List<DocumentCheckIssue> CollectNonBodyBlankLineIssuesByOutlineFind(Word.Document doc, int scanStart, DocumentCheckProgress progress)
        {
            List<DocumentCheckIssue> issues = new List<DocumentCheckIssue>();
            List<Word.Range> headings = CollectOutlineRangesByFind(doc, scanStart, progress, includeBlankHeadings: true)
                .OrderBy(item => item.Start)
                .ToList();
            foreach (Word.Range range in headings)
            {
                ThrowIfOperationCancelled();
                progress.Step("正在检查章节标题为空");
                string reason = GetMissingHeadingTextReason(range);
                if (!string.IsNullOrWhiteSpace(reason))
                {
                    issues.Add(CreateIssue(range, reason));
                }
            }

            return issues;
        }

        private static string GetMissingHeadingTextReason(Word.Range range)
        {
            Word.Paragraph paragraph = GetHostParagraph(range);
            int level = GetDocumentCheckOutlineLevel(paragraph);
            string prefix = level >= 1 && level <= 9
                ? $"章节标题为空：大纲{level}级"
                : "章节标题为空";
            string listString = GetListString(paragraph);
            string text = CleanParagraphText(range?.Text);
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.IsNullOrWhiteSpace(listString)
                    ? $"{prefix}没有标题文字。"
                    : $"{prefix}只有自动编号“{listString}”，没有标题文字。";
            }

            string normalized = NormalizeDigits(text)
                .Replace('．', '.')
                .Trim();
            if (Regex.IsMatch(normalized, @"^\d+(?:\.\d+)*[\.、]?$"))
            {
                return $"{prefix}只有手工编号“{text}”，没有标题文字。";
            }

            return string.Empty;
        }

        private static List<DocumentCheckIssue> CollectStyleConsistencyIssues(Word.Document doc, int scanStart, DocumentCheckProgress progress)
        {
            List<DocumentCheckIssue> issues = new List<DocumentCheckIssue>();
            Dictionary<int, string> expectedStylesByLevel = new Dictionary<int, string>();
            foreach (Word.Range range in CollectOutlineRangesByFind(doc, scanStart, progress).OrderBy(item => item.Start))
            {
                ThrowIfOperationCancelled();
                progress.Step("正在检查样式一致性");
                Word.Paragraph paragraph = GetHostParagraph(range);
                int level = GetDocumentCheckOutlineLevel(paragraph);
                if (level < 1 || level > 9)
                {
                    continue;
                }

                string styleName = GetParagraphStyleName(paragraph);
                if (string.IsNullOrWhiteSpace(styleName))
                {
                    continue;
                }

                if (!expectedStylesByLevel.TryGetValue(level, out string expectedStyle))
                {
                    expectedStylesByLevel[level] = styleName;
                    continue;
                }

                if (!string.Equals(expectedStyle, styleName, StringComparison.CurrentCultureIgnoreCase))
                {
                    issues.Add(CreateIssue(range, $"样式一致性：大纲{level}级应使用“{expectedStyle}”，当前为“{styleName}”。"));
                }
            }

            return issues;
        }

        private static List<Word.Range> CollectHeadingRangesByGoTo(Word.Document doc, bool includeBlankHeadings)
        {
            List<Word.Range> ranges = new List<Word.Range>();
            if (doc?.Content == null)
            {
                return ranges;
            }

            HashSet<int> seenStarts = new HashSet<int>();
            Word.Range probeRange = null;
            try
            {
                int docEnd = doc.Content.End;
                probeRange = doc.Range(doc.Content.Start, doc.Content.Start);
                for (int guard = 0; guard < 10000; guard++)
                {
                    ThrowIfOperationCancelled();
                    Word.Range headingRange = probeRange.GoTo(
                        Word.WdGoToItem.wdGoToHeading,
                        Word.WdGoToDirection.wdGoToNext);
                    Word.Paragraph paragraph = GetHostParagraph(headingRange);
                    Word.Range paragraphRange = paragraph?.Range;
                    if (paragraphRange == null || paragraphRange.Start >= docEnd || !seenStarts.Add(paragraphRange.Start))
                    {
                        break;
                    }

                    int level = GetDocumentCheckOutlineLevel(paragraph);
                    if (level >= 1 && level <= 9
                        && !IsInTable(paragraphRange)
                        && (includeBlankHeadings || !string.IsNullOrWhiteSpace(CleanParagraphText(paragraphRange.Text))))
                    {
                        ranges.Add(paragraphRange.Duplicate);
                    }

                    int nextStart = Math.Min(Math.Max(paragraphRange.End, paragraphRange.Start + 1), docEnd);
                    if (nextStart >= docEnd)
                    {
                        break;
                    }

                    probeRange.SetRange(nextStart, nextStart);
                }
            }
            catch
            {
            }

            return ranges.OrderBy(item => item.Start).ToList();
        }

        private static List<Word.Range> CollectOutlineRangesByFind(
            Word.Document doc,
            int scanStart,
            DocumentCheckProgress progress,
            bool includeBlankHeadings = false)
        {
            List<Word.Range> ranges = new List<Word.Range>();
            if (doc?.Content == null)
            {
                return ranges;
            }

            HashSet<int> seenStarts = new HashSet<int>();
            for (int level = 1; level <= 9; level++)
            {
                ThrowIfOperationCancelled();
                progress.Step("正在定位标题");
                Word.Range searchRange = doc.Content.Duplicate;
                searchRange.SetRange(scanStart, doc.Content.End);

                Word.Find find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "^p";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = true;
                find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)level;

                while (find.Execute())
                {
                    ThrowIfOperationCancelled();
                    Word.Range found = GetHostParagraph(searchRange)?.Range;
                    if (found != null
                        && seenStarts.Add(found.Start)
                        && !IsInTable(found)
                        && (includeBlankHeadings || !string.IsNullOrWhiteSpace(CleanParagraphText(found.Text))))
                    {
                        ranges.Add(found.Duplicate);
                    }

                    int nextStart = Math.Max(found?.End ?? searchRange.End, searchRange.Start + 1);
                    if (nextStart >= doc.Content.End)
                    {
                        break;
                    }

                    searchRange.SetRange(nextStart, doc.Content.End);
                    find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = "^p";
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = true;
                    find.ParagraphFormat.OutlineLevel = (Word.WdOutlineLevel)level;
                }
            }

            return ranges;
        }

        private static int CountDocumentCheckSteps(Word.Document doc, int scanStart, bool checkBlankLines, bool checkCaptions, bool checkLists, bool checkStyles, bool checkBrokenReferences)
        {
            int captionCount = 1;
            try
            {
                captionCount = checkCaptions
                    ? Math.Max(1, CollectCaptionListEntries(doc).Count(item => item != null && item.Start >= scanStart))
                    : 0;
            }
            catch
            {
            }

            int steps = checkBlankLines ? 9 : 0;
            steps += captionCount;
            steps += checkLists ? 9 : 0;
            steps += checkStyles ? 1 : 0;
            steps += checkBrokenReferences ? 1 : 0;

            return Math.Max(1, steps);
        }

        private static int[] BuildExpectedNumbers(int[] counters, int level)
        {
            int[] expected = new int[level];
            for (int i = 1; i <= level; i++)
            {
                expected[i - 1] = i == level ? counters[i] + 1 : Math.Max(counters[i], 1);
            }

            return expected;
        }

        private static void ApplyActualCounters(int[] counters, int[] actualNumbers, int level)
        {
            for (int i = 1; i <= level && i <= actualNumbers.Length; i++)
            {
                counters[i] = actualNumbers[i - 1];
            }

            for (int i = level + 1; i < counters.Length; i++)
            {
                counters[i] = 0;
            }
        }

        private static int GetDocumentCheckOutlineLevel(Word.Paragraph paragraph)
        {
            try
            {
                Word.WdOutlineLevel level = paragraph.OutlineLevel;
                return level >= Word.WdOutlineLevel.wdOutlineLevel1 && level <= Word.WdOutlineLevel.wdOutlineLevel9 ? (int)level : 0;
            }
            catch
            {
                return 0;
            }
        }

        private static bool IsInTable(Word.Range range)
        {
            try
            {
                return range != null && range.Information[Word.WdInformation.wdWithInTable];
            }
            catch
            {
                return false;
            }
        }

        private static string GetParagraphStyleName(Word.Paragraph paragraph)
        {
            try
            {
                object style = paragraph?.get_Style();
                if (style is Word.Style wordStyle)
                {
                    return wordStyle.NameLocal ?? string.Empty;
                }

                return style?.ToString() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetListString(Word.Paragraph paragraph)
        {
            try
            {
                return paragraph.Range?.ListFormat?.ListString ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int[] ParseDottedNumber(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return new int[0];
            }

            string normalized = NormalizeDigits(value.Trim());
            Match match = Regex.Match(normalized, @"^(\d+(?:\.\d+)*)");
            if (!match.Success)
            {
                return new int[0];
            }

            return match.Groups[1].Value.Split('.').Select(part => int.TryParse(part, out int number) ? number : 0).Where(number => number > 0).ToArray();
        }

        private static int ParseCaptionNumber(string value)
        {
            string normalized = NormalizeDigits(value);
            return int.TryParse(normalized, out int number) ? number : 0;
        }

        private static string NormalizeDigits(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            char[] chars = value.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] >= '０' && chars[i] <= '９')
                {
                    chars[i] = (char)('0' + chars[i] - '０');
                }
            }

            return new string(chars);
        }

        private static string CleanParagraphText(string text)
        {
            return (text ?? string.Empty).Replace("\r", string.Empty).Replace("\a", string.Empty).Trim();
        }

        private static DocumentCheckIssue CreateIssue(Word.Range sourceRange, string reason)
        {
            int start = sourceRange?.Start ?? 0;
            int end = sourceRange?.End ?? start;
            if (end > start)
            {
                end = Math.Max(start, end - 1);
            }

            return new DocumentCheckIssue(start, end, reason);
        }

        private static IList<NavigationPaneEntry> ToNavigationEntries(IList<DocumentCheckIssue> issues)
        {
            if (issues == null || issues.Count == 0)
            {
                return Array.Empty<NavigationPaneEntry>();
            }

            return issues
                .Where(issue => issue != null)
                .Select(issue => new NavigationPaneEntry
                {
                    Start = issue.Start,
                    Text = issue.Reason
                })
                .ToList();
        }

        private static void TrySetStatusBar(Word.Application app, string message)
        {
            if (app == null)
            {
                return;
            }

            try
            {
                app.StatusBar = message;
            }
            catch
            {
            }
        }

        private sealed class DocumentCheckProgress
        {
            private readonly Word.Application app;
            private readonly int total;
            private readonly int interval;
            private int current;

            internal DocumentCheckProgress(Word.Application app, int totalSteps)
            {
                this.app = app;
                total = Math.Max(1, totalSteps);
                interval = total > 3000 ? 100 : 20;
            }

            internal void Report(string message)
            {
                TrySetStatusBar(app, $"文档检查：{message}...");
            }

            internal void Step(string message)
            {
                current++;
                if (current == 1 || current == total || current % interval == 0)
                {
                    int percent = Math.Min(100, Math.Max(0, current * 100 / total));
                    TrySetStatusBar(app, $"文档检查：{message} {current}/{total} ({percent}%)");
                }
            }

            internal void Complete(int issueCount)
            {
                TrySetStatusBar(app, issueCount > 0
                    ? $"文档检查完成：发现 {issueCount} 个问题"
                    : "文档检查完成：未发现问题");
            }
        }

        private sealed class DocumentCheckIssue
        {
            internal DocumentCheckIssue(int start, int end, string reason)
            {
                Start = start;
                End = end;
                Reason = reason;
            }

            internal int Start { get; }

            internal int End { get; }

            internal string Reason { get; }

            public override string ToString()
            {
                return Reason;
            }
        }

        private sealed class SoftwareDocumentCheckResultForm : Form
        {
            private readonly Word.Document doc;
            private readonly ListBox listBox;

            internal SoftwareDocumentCheckResultForm(Word.Document doc, IReadOnlyList<DocumentCheckIssue> issues)
            {
                this.doc = doc;
                Text = "软件文档检查结果";
                Width = 620;
                Height = 420;
                TopMost = true;
                StartPosition = FormStartPosition.CenterScreen;

                TableLayoutPanel layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    Padding = new Padding(10),
                    RowCount = 3,
                    ColumnCount = 1
                };
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                layout.Controls.Add(new Label
                {
                    AutoSize = true,
                    Dock = DockStyle.Fill,
                    Text = issues.Count > 0 ? $"发现 {issues.Count} 个问题。" : "未发现问题。"
                }, 0, 0);

                listBox = new ListBox
                {
                    Dock = DockStyle.Fill,
                    HorizontalScrollbar = true,
                    IntegralHeight = false
                };

                if (issues.Count == 0)
                {
                    listBox.Items.Add("未发现问题");
                }
                else
                {
                    foreach (DocumentCheckIssue issue in issues)
                    {
                        listBox.Items.Add(issue);
                    }
                }

                listBox.DoubleClick += (_, __) => NavigateSelected();
                listBox.KeyDown += (sender, e) =>
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                        NavigateSelected();
                    }
                };
                layout.Controls.Add(listBox, 0, 1);

                FlowLayoutPanel buttons = new FlowLayoutPanel
                {
                    AutoSize = true,
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.RightToLeft
                };
                buttons.Controls.Add(CreateButton("关闭", (_, __) => Close()));
                buttons.Controls.Add(CreateButton("下一个", (_, __) => MoveSelection(1)));
                buttons.Controls.Add(CreateButton("上一个", (_, __) => MoveSelection(-1)));
                layout.Controls.Add(buttons, 0, 2);

                Controls.Add(layout);
                Shown += (_, __) =>
                {
                    if (listBox.Items.Count > 0)
                    {
                        listBox.SelectedIndex = 0;
                    }
                };
            }

            private static Button CreateButton(string text, EventHandler onClick)
            {
                Button button = new Button
                {
                    AutoSize = true,
                    Text = text
                };
                button.Click += onClick;
                return button;
            }

            private void MoveSelection(int offset)
            {
                if (listBox.Items.Count == 0 || !(listBox.Items[0] is DocumentCheckIssue))
                {
                    return;
                }

                int next = listBox.SelectedIndex < 0 ? 0 : listBox.SelectedIndex + offset;
                listBox.SelectedIndex = Math.Max(0, Math.Min(listBox.Items.Count - 1, next));
                NavigateSelected();
            }

            private void NavigateSelected()
            {
                if (!(listBox.SelectedItem is DocumentCheckIssue issue))
                {
                    return;
                }

                try
                {
                    doc?.Activate();
                    Word.Range range = doc.Range(issue.Start, Math.Max(issue.Start, issue.End));
                    range.Select();
                    try
                    {
                        doc.Application.ActiveWindow.ScrollIntoView(range, true);
                    }
                    catch
                    {
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"跳转失败: {ex.Message}", "文档检查");
                }
            }
        }

        private sealed class SoftwareDocumentSpec
        {
            internal SoftwareDocumentSpec(string name, string[] aliases, string[] requiredHeadings)
            {
                Name = name;
                Aliases = aliases;
                RequiredHeadings = requiredHeadings;
            }

            internal string Name { get; }

            internal string[] Aliases { get; }

            internal string[] RequiredHeadings { get; }
        }

        private sealed class SoftwareDocumentHeading
        {
            internal SoftwareDocumentHeading(string title, string normalizedTitle, int level, int start)
            {
                Title = title;
                NormalizedTitle = normalizedTitle;
                Level = level;
                Start = start;
            }

            internal string Title { get; }

            internal string NormalizedTitle { get; }

            internal int Level { get; }

            internal int Start { get; }
        }

        private static readonly SoftwareDocumentSpec[] SoftwareDocumentSpecs =
        {
            new SoftwareDocumentSpec("软件需求规格说明", new[] { "软件需求规格说明" }, new[]
            {
                "范围", "标识", "系统概述", "文档概述", "引用文档", "需求", "要求的状态和方式", "CSCI能力需求", "CSCI外部接口需求", "接口标识和接口图", "CSCI内部接口需求", "CSCI内部数据需求",
                "适应性需求", "保密性(Security)需求", "安全性(Safety)需求", "CSCI环境适应性需求", "其他质量特性", "计算机资源需求", "计算机硬件需求", "计算机硬件资源使用需求", "计算机软件需求", "计算机通信需求",
                "设计和实现约束", "人员相关需求", "培训相关需求", "软件保障需求", "其他需求", "需求的优先顺序和关键性",
                "合格性规定", "需求可追踪性", "注释"
            }),
            new SoftwareDocumentSpec("软件设计说明", new[] { "软件设计说明", "软件概要设计说明", "软件详细设计说明" }, new[]
            {
                "范围", "标识", "系统概述", "文档概述", "引用文档", "系统设计决策", "系统体系结构设计", "需求可追踪性", "注释"
            })
        };

        private static readonly Regex CaptionNumberRegex = new Regex(@"^\s*(?<kind>[图表])\s*(?<num>[0-9０-９]+)(?!\s*[（(]\s*续)", RegexOptions.Compiled);
    }
}
