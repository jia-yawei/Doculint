using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal enum DocumentMarkerDocumentType
    {
        Unknown,
        RequirementSpecification,
        SystemSpecification,
        SoftwareDesign,
        TestSpecification
    }

    internal sealed class DocumentMarkerCollectionResult
    {
        public DocumentMarkerDocumentType DocumentType { get; set; }
        public List<NavigationPaneEntry> Entries { get; set; } = new List<NavigationPaneEntry>();
    }

    internal struct DocumentRangeSpan
    {
        public DocumentRangeSpan(int start, int end)
        {
            Start = start;
            End = end;
        }

        public int Start { get; }

        public int End { get; }
    }

    internal static class DocumentMarkerService
    {
        private static readonly Regex RequirementMarkerRegex = new Regex(
            @"(?<![A-Za-z0-9/])/?(?:(?:[A-Za-z0-9]+(?:/[A-Za-z0-9]+)*-)+)?SRS-[^-\s]+(?:-[^-\s]+)*(?![-A-Za-z0-9])",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex TestMarkerRegex = new Regex(
            @"(?<![A-Za-z0-9/])/?(?:(?:[A-Za-z0-9]+-)+)?SCT_TC_\d{4,6}(?![A-Za-z0-9])",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex SystemMarkerRegex = new Regex(
            @"(?<![A-Za-z0-9/])/?(?:(?:[A-Za-z0-9]+(?:/[A-Za-z0-9]+)*-)+)?(?:SSS|SDTD)-[^-\s]+(?:-[^-\s]+)*(?![-A-Za-z0-9])",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex SoftwareDesignMarkerRegex = new Regex(
            @"(?<![A-Za-z0-9/])/?(?:(?:[A-Za-z0-9]+(?:/[A-Za-z0-9]+)*-)+)?(?:SDS|SDD)-[^-\s]+(?:-[^-\s]+)*(?![-A-Za-z0-9])",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex SlashSeparatedMarkerRegex = new Regex(
            @"(?<![A-Za-z0-9/])/?[A-Za-z][A-Za-z0-9]*/[A-Za-z0-9]+-\d+(?![A-Za-z0-9_-])",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static DocumentMarkerCollectionResult CollectMarkers(Word.Document doc)
        {
            return CollectMarkers(doc, Enumerable.Empty<string>(), 0, 0);
        }

        public static DocumentMarkerCollectionResult CollectMarkers(
            Word.Document doc,
            IEnumerable<string> markerTemplates)
        {
            return CollectMarkers(doc, markerTemplates, 0, 0);
        }

        public static DocumentMarkerCollectionResult CollectMarkers(
            Word.Document doc,
            IEnumerable<string> markerTemplates,
            int startPage,
            int endPage)
        {
            return CollectMarkers(
                doc,
                markerTemplates,
                startPage,
                endPage,
                DocumentMarkerDocumentType.Unknown,
                false);
        }

        public static DocumentMarkerCollectionResult CollectMarkers(
            Word.Document doc,
            IEnumerable<string> markerTemplates,
            int startPage,
            int endPage,
            DocumentMarkerDocumentType presetTemplateType)
        {
            return CollectMarkers(
                doc,
                markerTemplates,
                startPage,
                endPage,
                presetTemplateType,
                false);
        }

        public static DocumentMarkerCollectionResult CollectMarkers(
            Word.Document doc,
            IEnumerable<string> markerTemplates,
            int startPage,
            int endPage,
            DocumentMarkerDocumentType presetTemplateType,
            bool scanFieldResults)
        {
            DocumentMarkerCollectionResult result = new DocumentMarkerCollectionResult
            {
                DocumentType = presetTemplateType == DocumentMarkerDocumentType.Unknown
                    ? DetectDocumentType(doc)
                    : presetTemplateType
            };

            if (doc == null)
            {
                return result;
            }

            List<Regex> customMarkerPatterns = BuildCustomMarkerRegexes(markerTemplates, 0);
            IEnumerable<Regex> patterns = customMarkerPatterns.Count == 0
                ? GetMarkerPatterns(result.DocumentType)
                : customMarkerPatterns;
            int scanStart = GetScanStartAfterToc(doc);
            int scanEnd = doc.Content?.End ?? int.MaxValue;
            bool hasPageRange = startPage > 0 || endPage > 0;
            bool usePrecisePageFilter = false;
            DocumentRangeSpan pageRange = new DocumentRangeSpan(0, int.MaxValue);
            if (hasPageRange)
            {
                // Word 文档可能在分节后重新开始页码。字符位置的 GoTo 页范围
                // 在这种情况下不一定对应界面显示的页码，因此段落扫描时
                // 优先使用 Range.Information 返回的实际页码进行筛选。
                if (!TryGetPageScanRange(doc, startPage, endPage, out pageRange))
                {
                    pageRange = new DocumentRangeSpan(0, scanEnd);
                }

                scanStart = Math.Max(scanStart, pageRange.Start);
                scanEnd = Math.Min(scanEnd, pageRange.End);
                if (scanEnd <= scanStart)
                {
                    return result;
                }

                usePrecisePageFilter = !IsPageRangeBoundaryReliable(
                    doc,
                    pageRange,
                    Math.Max(1, startPage),
                    Math.Max(Math.Max(1, startPage), endPage));
            }

            Dictionary<string, NavigationPaneEntry> entryMap = new Dictionary<string, NavigationPaneEntry>(StringComparer.OrdinalIgnoreCase);
            List<DocumentRangeSpan> traceTableRanges = CollectRequirementTraceTableRanges(doc);
            int chapterCutoff = GetChapterCutoff(result.DocumentType);
            bool hasChapterCutoff = customMarkerPatterns.Count == 0 && chapterCutoff > 0;

            try
            {
                Word.Paragraphs paragraphs = doc.Paragraphs;
                if (hasPageRange)
                {
                    Word.Range scanRange = doc.Range(scanStart, scanEnd);
                    paragraphs = scanRange.Paragraphs;
                }

                foreach (Word.Paragraph paragraph in paragraphs)
                {
                    Word.Range range = paragraph?.Range;
                    if (range == null || range.End <= scanStart)
                    {
                        continue;
                    }

                    if (range.Start >= scanEnd)
                    {
                        break;
                    }

                    if (usePrecisePageFilter && !IsRangeOnRequestedPages(
                        range,
                        Math.Max(1, startPage),
                        Math.Max(Math.Max(1, startPage), endPage),
                        pageRange))
                    {
                        continue;
                    }

                    if (IsWithinAnyRange(range.Start, traceTableRanges))
                    {
                        continue;
                    }

                    string text = NormalizeParagraphText(range.Text);
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    if (hasChapterCutoff && IsAtOrBeyondCutoffChapter(text, chapterCutoff, result.DocumentType))
                    {
                        break;
                    }

                    foreach (Regex pattern in patterns)
                    {
                        AddMarkerEntries(entryMap, range, text, pattern);
                    }
                }
            }
            catch
            {
            }

            if (scanFieldResults)
            {
                ScanFieldResultMarkers(
                    doc,
                    patterns,
                    entryMap,
                    scanStart,
                    scanEnd,
                    hasPageRange,
                    usePrecisePageFilter,
                    startPage,
                    endPage,
                    pageRange,
                    traceTableRanges);
            }

            result.Entries = entryMap.Values
                .OrderBy(item => item.Start)
                .ThenBy(item => item.Text, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return result;
        }

        private static void ScanFieldResultMarkers(
            Word.Document doc,
            IEnumerable<Regex> patterns,
            Dictionary<string, NavigationPaneEntry> entryMap,
            int scanStart,
            int scanEnd,
            bool hasPageRange,
            bool usePrecisePageFilter,
            int requestedStartPage,
            int requestedEndPage,
            DocumentRangeSpan pageRange,
            IEnumerable<DocumentRangeSpan> traceTableRanges)
        {
            if (doc?.Fields == null || patterns == null || entryMap == null)
            {
                return;
            }

            try
            {
                foreach (Word.Field field in doc.Fields)
                {
                    Word.Range resultRange = field?.Result;
                    if (resultRange == null ||
                        resultRange.End <= scanStart ||
                        resultRange.Start >= scanEnd ||
                        IsWithinAnyRange(resultRange.Start, traceTableRanges))
                    {
                        continue;
                    }

                    if (hasPageRange && usePrecisePageFilter && !IsRangeOnRequestedPages(
                        resultRange,
                        Math.Max(1, requestedStartPage),
                        Math.Max(Math.Max(1, requestedStartPage), requestedEndPage),
                        pageRange))
                    {
                        continue;
                    }

                    string text = NormalizeParagraphText(resultRange.Text);
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    foreach (Regex pattern in patterns)
                    {
                        AddMarkerEntries(entryMap, resultRange, text, pattern);
                    }
                }
            }
            catch
            {
            }
        }

        private static bool IsPageRangeBoundaryReliable(
            Word.Document doc,
            DocumentRangeSpan pageRange,
            int requestedStartPage,
            int requestedEndPage)
        {
            if (doc?.Content == null || pageRange.End <= pageRange.Start)
            {
                return false;
            }

            try
            {
                int probeEnd = Math.Min(pageRange.End, doc.Content.End);
                if (probeEnd <= pageRange.Start)
                {
                    return false;
                }

                Word.Range probe = doc.Range(pageRange.Start, probeEnd);
                int actualStartPage = GetRangePageNumber(probe, false);
                int actualEndPage = GetRangePageNumber(probe, true);
                return actualStartPage == requestedStartPage &&
                       actualEndPage >= requestedStartPage &&
                       actualEndPage <= requestedEndPage;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsRangeOnRequestedPages(
            Word.Range range,
            int requestedStartPage,
            int requestedEndPage,
            DocumentRangeSpan fallbackPageRange)
        {
            if (range == null)
            {
                return false;
            }

            int startPage = GetRangePageNumber(range, false);
            int endPage = GetRangePageNumber(range, true);
            if (startPage > 0 || endPage > 0)
            {
                if (startPage <= 0)
                {
                    startPage = endPage;
                }

                if (endPage <= 0)
                {
                    endPage = startPage;
                }

                return endPage >= requestedStartPage && startPage <= requestedEndPage;
            }

            // 某些 Word 兼容实现不支持 Range.Information，回退到字符位置范围。
            return range.End > fallbackPageRange.Start && range.Start < fallbackPageRange.End;
        }

        private static int GetRangePageNumber(Word.Range range, bool atEnd)
        {
            if (range == null)
            {
                return 0;
            }

            try
            {
                Word.Range probe = range.Duplicate;
                if (atEnd)
                {
                    probe.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    if (probe.Start > range.Start)
                    {
                        probe.MoveStart(Word.WdUnits.wdCharacter, -1);
                    }
                }
                else
                {
                    probe.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }

                object value = probe.Information[Word.WdInformation.wdActiveEndPageNumber];
                return Convert.ToInt32(value);
            }
            catch
            {
                return 0;
            }
        }

        private static List<Regex> BuildCustomMarkerRegexes(
            IEnumerable<string> markerIdentifiers,
            int numberDigits)
        {
            List<Regex> patterns = new List<Regex>();
            foreach (string value in (markerIdentifiers ?? Enumerable.Empty<string>())
                .Select(item => (item ?? string.Empty).Trim())
                .Where(item => item.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(3))
            {
                patterns.Add(BuildCustomMarkerRegex(value, numberDigits));
            }

            return patterns;
        }

        private static Regex BuildCustomMarkerRegex(string markerIdentifier, int numberDigits)
        {
            string format = (markerIdentifier ?? string.Empty).Trim().TrimStart('/');
            if (format.IndexOf('#') >= 0)
            {
                return BuildWildcardMarkerRegex(format);
            }

            if (Regex.IsMatch(format, @"^[A-Za-z][A-Za-z0-9_]*$"))
            {
                string digits = numberDigits > 0 ? @"\d{" + numberDigits + "}" : @"\d+";
                return new Regex(
                    @"(?<![A-Za-z0-9/])/?(?:(?:[A-Za-z0-9]+(?:/[A-Za-z0-9]+)*-)+)?"
                    + Regex.Escape(format)
                    + "-" + digits + @"(?![A-Za-z0-9])",
                    RegexOptions.IgnoreCase | RegexOptions.Compiled);
            }

            Match trailingNumber = Regex.Match(format, @"(?<prefix>.*?)(?<number>\d+)$");
            if (!trailingNumber.Success)
            {
                string digits = numberDigits > 0 ? @"\d{" + numberDigits + "}" : @"\d+";
                return new Regex(
                    "(?<![A-Za-z0-9/])/?" + Regex.Escape(format.TrimEnd('-'))
                    + "-" + digits + "(?![A-Za-z0-9])",
                    RegexOptions.IgnoreCase | RegexOptions.Compiled);
            }

            string prefix = Regex.Escape(trailingNumber.Groups["prefix"].Value);
            int digitCount = numberDigits > 0 ? numberDigits : trailingNumber.Groups["number"].Value.Length;
            return new Regex(
                "(?<![A-Za-z0-9/])/?" + prefix + @"\d{" + digitCount + "}(?![A-Za-z0-9])",
                RegexOptions.IgnoreCase | RegexOptions.Compiled);
        }

        private static Regex BuildWildcardMarkerRegex(string template)
        {
            StringBuilder pattern = new StringBuilder(@"(?<![A-Za-z0-9/])/?");
            foreach (char character in template ?? string.Empty)
            {
                pattern.Append(character == '#'
                    ? @"[^-\s]+"
                    : Regex.Escape(character.ToString()));
            }

            pattern.Append(@"(?![-A-Za-z0-9])");
            return new Regex(pattern.ToString(), RegexOptions.IgnoreCase | RegexOptions.Compiled);
        }

        private static List<DocumentRangeSpan> CollectRequirementTraceTableRanges(Word.Document doc)
        {
            List<DocumentRangeSpan> ranges = new List<DocumentRangeSpan>();
            if (doc == null)
            {
                return ranges;
            }

            try
            {
                Word.Bookmarks bookmarks = doc.Bookmarks;
                for (int index = 1; index <= bookmarks.Count; index++)
                {
                    Word.Bookmark bookmark = bookmarks[index];
                    if (bookmark == null ||
                        !bookmark.Name.StartsWith(RequirementTraceTableExporter.TraceTableBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    Word.Range range = bookmark.Range;
                    if (range != null)
                    {
                        ranges.Add(new DocumentRangeSpan(range.Start, range.End));
                    }
                }
            }
            catch
            {
            }

            try
            {
                Word.Tables tables = doc.Tables;
                for (int index = 1; index <= tables.Count; index++)
                {
                    Word.Table table = tables[index];
                    if (!IsLegacyRequirementTraceTable(table))
                    {
                        continue;
                    }

                    Word.Range range = table.Range;
                    if (range != null)
                    {
                        ranges.Add(new DocumentRangeSpan(range.Start, range.End));
                    }
                }
            }
            catch
            {
            }

            return ranges;
        }

        private static bool IsLegacyRequirementTraceTable(Word.Table table)
        {
            if (table == null)
            {
                return false;
            }

            try
            {
                if (table.Rows.Count < 2 || table.Columns.Count != 6)
                {
                    return false;
                }

                string headerText = string.Join("|", Enumerable.Range(1, 6)
                    .Select(column => NormalizeParagraphText(table.Cell(2, column).Range.Text)));
                return headerText.IndexOf("标识", StringComparison.OrdinalIgnoreCase) >= 0 &&
                       headerText.IndexOf("章节号", StringComparison.OrdinalIgnoreCase) >= 0 &&
                       (headerText.IndexOf("要求名称", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        headerText.IndexOf("需求名称", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsWithinAnyRange(int position, IEnumerable<DocumentRangeSpan> ranges)
        {
            return ranges != null && ranges.Any(range => position >= range.Start && position < range.End);
        }

        public static string GetDocumentTypeDisplayName(DocumentMarkerDocumentType documentType)
        {
            switch (documentType)
            {
                case DocumentMarkerDocumentType.RequirementSpecification:
                    return "需求规格说明";
                case DocumentMarkerDocumentType.SystemSpecification:
                    return "系统规格说明/软件研制任务书";
                case DocumentMarkerDocumentType.SoftwareDesign:
                    return "软件设计说明";
                case DocumentMarkerDocumentType.TestSpecification:
                    return "软件测试说明";
                default:
                    return "未识别文档类型";
            }
        }

        public static bool IsMarkerPaneSupportedDocumentType(DocumentMarkerDocumentType documentType)
        {
            return documentType == DocumentMarkerDocumentType.SystemSpecification
                || documentType == DocumentMarkerDocumentType.RequirementSpecification
                || documentType == DocumentMarkerDocumentType.SoftwareDesign
                || documentType == DocumentMarkerDocumentType.TestSpecification;
        }

        private static void AddMarkerEntries(
            Dictionary<string, NavigationPaneEntry> entryMap,
            Word.Range paragraphRange,
            string paragraphText,
            Regex pattern)
        {
            if (entryMap == null || paragraphRange == null || pattern == null || string.IsNullOrWhiteSpace(paragraphText))
            {
                return;
            }

            foreach (Match match in pattern.Matches(paragraphText))
            {
                if (!match.Success)
                {
                    continue;
                }

                string displayText = match.Value.Trim();
                int start = paragraphRange.Start + match.Index;
                if (displayText.StartsWith("/", StringComparison.Ordinal))
                {
                    displayText = displayText.Substring(1);
                    start++;
                }

                string key = $"{start}:{displayText}";
                if (entryMap.ContainsKey(key))
                {
                    continue;
                }

                entryMap[key] = new NavigationPaneEntry
                {
                    Start = start,
                    Text = displayText
                };
            }
        }

        private static IEnumerable<Regex> GetMarkerPatterns(DocumentMarkerDocumentType documentType)
        {
            switch (documentType)
            {
                case DocumentMarkerDocumentType.RequirementSpecification:
                    return new[] { RequirementMarkerRegex, SlashSeparatedMarkerRegex };
                case DocumentMarkerDocumentType.SystemSpecification:
                    return new[] { SystemMarkerRegex, SlashSeparatedMarkerRegex };
                case DocumentMarkerDocumentType.SoftwareDesign:
                    return new[] { SoftwareDesignMarkerRegex, SlashSeparatedMarkerRegex };
                case DocumentMarkerDocumentType.TestSpecification:
                    return new[] { TestMarkerRegex, SlashSeparatedMarkerRegex };
                default:
                    return new[]
                    {
                        RequirementMarkerRegex,
                        SystemMarkerRegex,
                        SoftwareDesignMarkerRegex,
                        TestMarkerRegex,
                        SlashSeparatedMarkerRegex
                    };
            }
        }

        internal static DocumentMarkerDocumentType DetectDocumentType(Word.Document doc)
        {
            string sample = BuildSampleText(doc);
            if (string.IsNullOrWhiteSpace(sample))
            {
                return DocumentMarkerDocumentType.Unknown;
            }

            if (ContainsAny(sample, "软件设计说明", "设计说明") || SoftwareDesignMarkerRegex.IsMatch(sample))
            {
                return DocumentMarkerDocumentType.SoftwareDesign;
            }

            if (ContainsAny(sample, "测试说明", "软件测试说明", "测试文档") || TestMarkerRegex.IsMatch(sample))
            {
                return DocumentMarkerDocumentType.TestSpecification;
            }

            if (ContainsAny(sample, "软件研制任务书", "系统/子系统规格说明", "系统子系统规格说明", "系统规格说明") || SystemMarkerRegex.IsMatch(sample))
            {
                return DocumentMarkerDocumentType.SystemSpecification;
            }

            if (ContainsAny(sample, "软件需求规格说明", "需求规格说明") || RequirementMarkerRegex.IsMatch(sample))
            {
                return DocumentMarkerDocumentType.RequirementSpecification;
            }

            return DocumentMarkerDocumentType.Unknown;
        }

        private static string BuildSampleText(Word.Document doc)
        {
            if (doc == null)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            try
            {
                if (!string.IsNullOrWhiteSpace(doc.Name))
                {
                    builder.AppendLine(doc.Name);
                }
            }
            catch
            {
            }

            int paragraphCount = 0;
            try
            {
                paragraphCount = doc.Paragraphs.Count;
            }
            catch
            {
                paragraphCount = 0;
            }

            for (int i = 1; i <= paragraphCount && builder.Length < 12000 && i <= 80; i++)
            {
                Word.Paragraph paragraph = null;
                try
                {
                    paragraph = doc.Paragraphs[i];
                }
                catch
                {
                    continue;
                }

                string text = NormalizeParagraphText(paragraph?.Range?.Text);
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                builder.AppendLine(text);
            }

            return builder.ToString();
        }

        private static int GetChapterCutoff(DocumentMarkerDocumentType documentType)
        {
            switch (documentType)
            {
                case DocumentMarkerDocumentType.RequirementSpecification:
                case DocumentMarkerDocumentType.TestSpecification:
                    return 5;
                case DocumentMarkerDocumentType.SoftwareDesign:
                    return 6;
                default:
                    return 0;
            }
        }

        private static bool IsAtOrBeyondCutoffChapter(string paragraphText, int cutoffChapter, DocumentMarkerDocumentType documentType)
        {
            if (string.IsNullOrWhiteSpace(paragraphText) || cutoffChapter <= 0)
            {
                return false;
            }

            if (documentType == DocumentMarkerDocumentType.RequirementSpecification || documentType == DocumentMarkerDocumentType.TestSpecification)
            {
                if (paragraphText.IndexOf("需求可追踪性", StringComparison.OrdinalIgnoreCase) >= 0
                    || paragraphText.IndexOf("需求的可追踪性", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            if (documentType == DocumentMarkerDocumentType.SoftwareDesign
                && paragraphText.IndexOf("需求可最终性", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            Match match = ChapterHeadingRegex.Match(paragraphText);
            if (!match.Success)
            {
                return false;
            }

            int chapterNumber = ParseChapterNumber(match.Groups["number"].Value);
            return chapterNumber >= cutoffChapter;
        }

        private static readonly Regex ChapterHeadingRegex = new Regex(
            @"^\s*第\s*(?<number>[0-9一二三四五六七八九十百千万零〇壹贰叁肆伍陆柒捌玖拾]+)\s*章\b",
            RegexOptions.Compiled);

        private static int ParseChapterNumber(string rawNumber)
        {
            if (string.IsNullOrWhiteSpace(rawNumber))
            {
                return 0;
            }

            if (int.TryParse(rawNumber, out int arabicNumber))
            {
                return arabicNumber;
            }

            return ParseChineseNumber(rawNumber);
        }

        private static int ParseChineseNumber(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            int total = 0;
            int current = 0;
            foreach (char ch in text)
            {
                int digit = GetChineseDigit(ch);
                if (digit >= 0)
                {
                    current = digit;
                    continue;
                }

                int unit = GetChineseUnit(ch);
                if (unit > 0)
                {
                    if (current == 0)
                    {
                        current = 1;
                    }

                    total += current * unit;
                    current = 0;
                }
            }

            return total + current;
        }

        private static int GetChineseDigit(char ch)
        {
            switch (ch)
            {
                case '零':
                case '〇':
                    return 0;
                case '一':
                case '壹':
                    return 1;
                case '二':
                case '贰':
                case '两':
                    return 2;
                case '三':
                case '叁':
                    return 3;
                case '四':
                case '肆':
                    return 4;
                case '五':
                case '伍':
                    return 5;
                case '六':
                case '陆':
                    return 6;
                case '七':
                case '柒':
                    return 7;
                case '八':
                case '捌':
                    return 8;
                case '九':
                case '玖':
                    return 9;
                default:
                    return -1;
            }
        }

        private static int GetChineseUnit(char ch)
        {
            switch (ch)
            {
                case '十':
                case '拾':
                    return 10;
                case '百':
                    return 100;
                case '千':
                    return 1000;
                case '万':
                    return 10000;
                default:
                    return 0;
            }
        }

        private static bool ContainsAny(string text, params string[] keywords)
        {
            if (string.IsNullOrWhiteSpace(text) || keywords == null)
            {
                return false;
            }

            foreach (string keyword in keywords)
            {
                if (!string.IsNullOrWhiteSpace(keyword)
                    && text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetPageScanRange(
            Word.Document doc,
            int requestedStartPage,
            int requestedEndPage,
            out DocumentRangeSpan pageRange)
        {
            pageRange = new DocumentRangeSpan(0, 0);
            if (doc?.Content == null)
            {
                return false;
            }

            int startPage = Math.Max(1, requestedStartPage);
            int endPage = Math.Max(startPage, requestedEndPage);
            try
            {
                doc.Repaginate();
                int pageCount = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);
                if (pageCount <= 0 || startPage > pageCount)
                {
                    return false;
                }

                endPage = Math.Min(endPage, pageCount);
                object what = Word.WdGoToItem.wdGoToPage;
                object which = Word.WdGoToDirection.wdGoToAbsolute;
                object count = startPage;
                Word.Range startRange = doc.GoTo(ref what, ref which, ref count);
                if (startRange == null)
                {
                    return false;
                }

                int start = startRange.Start;
                int end = doc.Content.End;
                if (endPage < pageCount)
                {
                    count = endPage + 1;
                    Word.Range nextRange = doc.GoTo(ref what, ref which, ref count);
                    if (nextRange != null && nextRange.Start > start)
                    {
                        end = nextRange.Start;
                    }
                }

                if (end <= start)
                {
                    return false;
                }

                pageRange = new DocumentRangeSpan(start, end);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static int GetScanStartAfterToc(Word.Document doc)
        {
            if (doc == null)
            {
                return 0;
            }

            int scanStart = 0;
            try
            {
                if (doc.TablesOfContents != null && doc.TablesOfContents.Count > 0)
                {
                    scanStart = doc.TablesOfContents.Cast<Word.TableOfContents>()
                        .Where(t => t != null && t.Range != null)
                        .Select(t => t.Range.End)
                        .DefaultIfEmpty(0)
                        .Max();
                }
            }
            catch
            {
            }

            return Math.Max(0, scanStart);
        }

        private static string NormalizeParagraphText(string text)
        {
            return (text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Trim();
        }
    }
}
