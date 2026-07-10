using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal static class RequirementTrackingWordService
    {
        private static readonly Regex RequirementIdRegex = new Regex(
            @"(?<![A-Za-z0-9])(?<id>(?:[A-Za-z0-9]+\s*[-－–—]\s*)*(?<suffix>(?<prefix>SSS|SRS)\s*[-－–—]\s*\d+))(?![A-Za-z0-9])",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex SrsRequirementIdRegex = new Regex(
            @"(?<![A-Za-z0-9])(?<id>SRS\s*[-－–—]\s*\d+)(?![A-Za-z0-9])",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ChapterHeadingRegex = new Regex(
            @"^\s*第\s*(?<number>[0-9一二三四五六七八九十百千万零〇壹贰叁肆伍陆柒捌玖拾]+)\s*章\b",
            RegexOptions.Compiled);

        private static readonly Regex NumberedHeadingRegex = new Regex(
            @"^\s*(?<number>\d+(?:[.．]\d+)+)(?:\s+|[　、．.]|(?=[\u4e00-\u9fff]))(?<title>.+)$",
            RegexOptions.Compiled);

        private static readonly Regex SimpleTopLevelHeadingRegex = new Regex(
            @"^\s*(?<number>\d+|[一二三四五六七八九十百千万零〇壹贰叁肆伍陆柒捌玖拾]+)(?:\s+|[、.．])(?<title>.+)$",
            RegexOptions.Compiled);

        public static IReadOnlyList<RequirementTrackingDocumentOption> CollectOpenDocumentOptions(Word.Application app)
        {
            List<RequirementTrackingDocumentOption> options = new List<RequirementTrackingDocumentOption>();
            if (app == null)
            {
                return options;
            }

            Word.Documents documents = null;
            try
            {
                documents = app.Documents;
                if (documents == null)
                {
                    return options;
                }

                int count = documents.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Document doc = null;
                    try
                    {
                        doc = documents[i];
                        string fullName = TryGetDocumentFullName(doc);
                        if (string.IsNullOrWhiteSpace(fullName))
                        {
                            continue;
                        }

                        options.Add(new RequirementTrackingDocumentOption
                        {
                            FullName = fullName,
                            DisplayName = Path.GetFileName(fullName)
                        });
                    }
                    finally
                    {
                        ReleaseComObject(doc);
                    }
                }
            }
            finally
            {
                ReleaseComObject(documents);
            }

            return options
                .OrderBy(item => item.DisplayName ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public static RequirementTrackingDocumentSnapshot CollectRequirements(Word.Document doc)
        {
            return CollectRequirements(doc, null);
        }

        public static RequirementTrackingDocumentSnapshot CollectRequirements(
            Word.Document doc,
            RequirementTrackingDocumentKind documentKind,
            Action<string> progressReporter)
        {
            switch (documentKind)
            {
                case RequirementTrackingDocumentKind.RequirementSpecification:
                    return CollectRequirements(doc, progressReporter);
                case RequirementTrackingDocumentKind.SystemSpecification:
                    throw new NotSupportedException("系统规格说明的需求解析逻辑尚未实现。");
                case RequirementTrackingDocumentKind.SoftwareDesignDescription:
                    throw new NotSupportedException("软件设计说明的需求解析逻辑尚未实现。");
                case RequirementTrackingDocumentKind.SoftwareTestDescription:
                    throw new NotSupportedException("软件测试说明的需求解析逻辑尚未实现。");
                default:
                    throw new NotSupportedException("当前文档类型不在需求追踪支持范围内。");
            }
        }

        private static Word.Range GoToPreviousHeadingFromPosition(Word.Document doc, int position)
        {
            if (doc == null || position <= 0)
            {
                return null;
            }

            Word.Range probeRange = null;
            try
            {
                int probeStart = Math.Max(doc.Content.Start, position - 1);
                probeRange = doc.Range(probeStart, probeStart);
                Word.Range headingRange = probeRange.GoTo(
                    Word.WdGoToItem.wdGoToHeading,
                    Word.WdGoToDirection.wdGoToPrevious);
                if (headingRange != null && headingRange.Start < position)
                {
                    return headingRange;
                }

                ReleaseComObject(headingRange);
                int fallbackStart = Math.Max(doc.Content.Start, position - 32);
                probeRange.SetRange(fallbackStart, fallbackStart);
                headingRange = probeRange.GoTo(
                    Word.WdGoToItem.wdGoToHeading,
                    Word.WdGoToDirection.wdGoToPrevious);
                return headingRange != null && headingRange.Start < position ? headingRange : null;
            }
            finally
            {
                ReleaseComObject(probeRange);
            }
        }

        private static Word.Range GoToPreviousNumberedHeadingFromPosition(
            Word.Document doc,
            int position,
            out string rawNumber,
            out string headingText)
        {
            rawNumber = string.Empty;
            headingText = string.Empty;
            if (doc == null || position <= 0)
            {
                return null;
            }

            int currentPosition = position;
            HashSet<int> visitedStarts = new HashSet<int>();
            for (int i = 0; i < 12; i++)
            {
                Word.Range headingRange = GoToPreviousHeadingFromPosition(doc, currentPosition);
                if (headingRange == null)
                {
                    return TryFindNearestNumberedListHeadingBeforePosition(
                        doc,
                        currentPosition,
                        out rawNumber,
                        out headingText);
                }

                if (!visitedStarts.Add(headingRange.Start))
                {
                    ReleaseComObject(headingRange);
                    return null;
                }

                string candidateText;
                string candidateNumber = TryResolveHeadingNumber(headingRange, out candidateText);
                if (!string.IsNullOrWhiteSpace(candidateNumber))
                {
                    rawNumber = candidateNumber;
                    headingText = candidateText;
                    return headingRange;
                }

                currentPosition = Math.Max(doc.Content.Start, headingRange.Start - 1);
                ReleaseComObject(headingRange);
            }

            return TryFindNearestNumberedListHeadingBeforePosition(
                doc,
                position,
                out rawNumber,
                out headingText);
        }

        private static string TryResolveHeadingNumber(Word.Range headingRange, out string headingText)
        {
            headingText = string.Empty;
            if (headingRange == null)
            {
                return string.Empty;
            }

            Word.Range paragraphRange = null;
            Word.Paragraph paragraph = null;
            try
            {
                string rangeListString = TryGetRangeListString(headingRange);
                paragraphRange = headingRange.Duplicate;
                paragraphRange.Expand(Word.WdUnits.wdParagraph);
                headingText = NormalizeParagraphText(paragraphRange.Text);
                paragraph = paragraphRange.Paragraphs[1];

                string paragraphListString = TryGetParagraphListNumber(paragraph);
                string textNumber;
                string textTitle;
                if (TryExtractSectionHeadingForContext(headingText, out textNumber, out textTitle))
                {
                    return textNumber;
                }

                if (!string.IsNullOrWhiteSpace(rangeListString))
                {
                    return rangeListString;
                }

                return paragraphListString;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(paragraph);
                ReleaseComObject(paragraphRange);
            }
        }

        private static Word.Range TryFindNearestNumberedListHeadingBeforePosition(
            Word.Document doc,
            int position,
            out string rawNumber,
            out string headingText)
        {
            rawNumber = string.Empty;
            headingText = string.Empty;
            if (doc == null || position <= 0)
            {
                return null;
            }

            Word.Range searchRange = null;
            Word.ListParagraphs listParagraphs = null;
            try
            {
                int searchStart = Math.Max(doc.Content.Start, position - 8000);
                searchRange = doc.Range(searchStart, position);
                listParagraphs = searchRange.ListParagraphs;
                if (listParagraphs == null)
                {
                    return null;
                }

                int count = listParagraphs.Count;
                int minIndex = Math.Max(1, count - 80);
                for (int i = count; i >= minIndex; i--)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = listParagraphs[i];
                        paragraphRange = paragraph?.Range;
                        if (paragraphRange == null ||
                            paragraphRange.Start >= position ||
                            IsRangeInsideTable(paragraphRange))
                        {
                            continue;
                        }

                        string listNumber = TryGetParagraphListNumber(paragraph);
                        if (string.IsNullOrWhiteSpace(listNumber) ||
                            !IsSectionNumberUnder(listNumber, "3"))
                        {
                            continue;
                        }

                        string text = NormalizeParagraphText(paragraphRange.Text);
                        if (string.IsNullOrWhiteSpace(text) ||
                            ContainsRequirementId(text, "SRS") ||
                            LooksLikeTocEntryText(text) ||
                            LooksLikeTableCaptionText(text))
                        {
                            continue;
                        }

                        rawNumber = listNumber;
                        headingText = text;
                        return paragraphRange.Duplicate;
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                    }
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseComObject(listParagraphs);
                ReleaseComObject(searchRange);
            }

            return null;
        }

        private static bool LooksLikeTableCaptionText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            string normalized = NormalizeParagraphText(text);
            return Regex.IsMatch(normalized, @"^表\s*\d+", RegexOptions.IgnoreCase) ||
                   Regex.IsMatch(normalized, @"^图\s*\d+", RegexOptions.IgnoreCase);
        }

        private static int GetHeadingProbeStart(Word.Range foundRange)
        {
            if (foundRange == null)
            {
                return 0;
            }

            int fallbackStart = foundRange.Start;
            Word.Tables tables = null;
            Word.Table table = null;
            Word.Range tableRange = null;
            try
            {
                bool isInTable = false;
                try
                {
                    isInTable = Convert.ToBoolean(foundRange.Information[Word.WdInformation.wdWithInTable]);
                }
                catch
                {
                    isInTable = false;
                }

                if (!isInTable)
                {
                    return fallbackStart;
                }

                tables = foundRange.Tables;
                if (tables == null || tables.Count < 1)
                {
                    return fallbackStart;
                }

                table = tables[1];
                tableRange = table?.Range;
                return tableRange == null ? fallbackStart : Math.Max(0, tableRange.Start);
            }
            catch
            {
                return fallbackStart;
            }
            finally
            {
                ReleaseComObject(tableRange);
                ReleaseComObject(table);
                ReleaseComObject(tables);
            }
        }

        private static string TryGetRangeListString(Word.Range range)
        {
            try
            {
                string listString = range?.ListFormat?.ListString ?? string.Empty;
                return NormalizeParagraphText(listString).Trim(' ', '\t', '.', '．', '、');
            }
            catch
            {
                return string.Empty;
            }
        }

        public static RequirementTrackingDocumentSnapshot CollectRequirements(Word.Document doc, Action<string> progressReporter)
        {
            RequirementTrackingDocumentSnapshot snapshot = CreateSnapshot(doc);
            if (doc == null)
            {
                return snapshot;
            }

            ReportProgress(progressReporter, "正在复用标识窗格识别需求标识...");
            List<NavigationPaneEntry> markerEntries = CollectSrsMarkerEntriesFromMarkerPaneService(doc);
            ReportProgress(progressReporter, $"标识窗格范围内识别到 {markerEntries.Count} 个SRS标识");
            if (markerEntries.Count > 0)
            {
                return CollectSrsThirdChapterRequirements(doc, progressReporter, markerEntries);
            }

            ReportProgress(progressReporter, "标识窗格未返回SRS标识，正在识别文档类型...");
            RequirementTrackingDocumentKind documentKind = DetectDocumentKind(doc);
            if (documentKind == RequirementTrackingDocumentKind.RequirementSpecification)
            {
                return CollectSrsThirdChapterRequirements(doc, progressReporter);
            }

            RequirementTrackingDocumentSnapshot paragraphSnapshot = CollectRequirementsByParagraphScan(doc, progressReporter, DetectPreferredRequirementPrefix(doc));
            if ((paragraphSnapshot.Requirements?.Count ?? 0) > 0)
            {
                return paragraphSnapshot;
            }

            return CollectSrsThirdChapterRequirements(doc, progressReporter);
        }

        public static string ResolveNearestSectionNumber(Word.Document doc, int position)
        {
            Word.Range headingRange = null;
            string rawNumber;
            string headingText;
            try
            {
                string currentParagraphNumber = TryResolveCurrentOutlineHeadingNumber(doc, position);
                if (!string.IsNullOrWhiteSpace(currentParagraphNumber))
                {
                    return currentParagraphNumber;
                }

                headingRange = GoToPreviousOutlineHeadingFromPosition(doc, position);
                if (headingRange != null)
                {
                    string number = TryResolveHeadingNumber(headingRange, out _);
                    if (!string.IsNullOrWhiteSpace(number))
                    {
                        return number;
                    }
                }

                return TryFindNearestOutlineNumberedHeadingBeforePosition(doc, position, out rawNumber, out headingText)
                    ? rawNumber
                    : string.Empty;
            }
            finally
            {
                ReleaseComObject(headingRange);
            }
        }

        public static string ResolveCurrentHeadingNumber(Word.Selection selection)
        {
            Word.Range headingRange = null;
            try
            {
                headingRange = selection?.Range?.Bookmarks?["\\HeadingLevel"]?.Range;
                string number = TryGetRangeListString(headingRange);
                if (!string.IsNullOrWhiteSpace(number))
                {
                    return number;
                }

                Word.Paragraph paragraph = headingRange?.Paragraphs?[1];
                try
                {
                    return TryGetParagraphListNumber(paragraph);
                }
                finally
                {
                    ReleaseComObject(paragraph);
                }
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(headingRange);
            }
        }

        private static string TryResolveCurrentOutlineHeadingNumber(Word.Document doc, int position)
        {
            Word.Range range = null;
            Word.Range paragraphRange = null;
            Word.Paragraph paragraph = null;
            try
            {
                int contentEnd = Math.Max(0, doc.Content.End - 1);
                int safeStart = Math.Max(0, Math.Min(position, contentEnd));
                range = doc.Range(safeStart, safeStart);
                paragraphRange = range.Duplicate;
                paragraphRange.Expand(Word.WdUnits.wdParagraph);
                paragraph = paragraphRange.Paragraphs[1];
                if (!IsOutlineHeadingParagraph(paragraph))
                {
                    return string.Empty;
                }

                string text = NormalizeParagraphText(paragraphRange.Text);
                if (string.IsNullOrWhiteSpace(text) || LooksLikeTableCaptionText(text))
                {
                    return string.Empty;
                }

                string number = TryGetParagraphListNumber(paragraph);
                if (!string.IsNullOrWhiteSpace(number))
                {
                    return number;
                }

                string title;
                return TryExtractSectionHeadingForContext(text, out number, out title) ? number : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(paragraph);
                ReleaseComObject(paragraphRange);
                ReleaseComObject(range);
            }
        }

        private static Word.Range GoToPreviousOutlineHeadingFromPosition(Word.Document doc, int position)
        {
            if (doc == null || position <= 0)
            {
                return null;
            }

            int currentPosition = position;
            HashSet<int> visitedStarts = new HashSet<int>();
            for (int i = 0; i < 20; i++)
            {
                Word.Range headingRange = GoToPreviousHeadingFromPosition(doc, currentPosition);
                if (headingRange == null)
                {
                    return null;
                }

                if (!visitedStarts.Add(headingRange.Start))
                {
                    ReleaseComObject(headingRange);
                    return null;
                }

                if (IsValidOutlineHeadingRange(headingRange))
                {
                    return headingRange;
                }

                currentPosition = Math.Max(doc.Content.Start, headingRange.Start - 1);
                ReleaseComObject(headingRange);
            }

            return null;
        }

        private static bool IsValidOutlineHeadingRange(Word.Range range)
        {
            Word.Range paragraphRange = null;
            Word.Paragraph paragraph = null;
            try
            {
                paragraphRange = range?.Duplicate;
                paragraphRange?.Expand(Word.WdUnits.wdParagraph);
                string text = NormalizeParagraphText(paragraphRange?.Text);
                if (string.IsNullOrWhiteSpace(text) || LooksLikeTableCaptionText(text))
                {
                    return false;
                }

                paragraph = paragraphRange.Paragraphs[1];
                Word.WdOutlineLevel level = paragraph.OutlineLevel;
                return level >= Word.WdOutlineLevel.wdOutlineLevel1 &&
                       level <= Word.WdOutlineLevel.wdOutlineLevel9;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComObject(paragraph);
                ReleaseComObject(paragraphRange);
            }
        }

        private static bool TryFindNearestOutlineNumberedHeadingBeforePosition(
            Word.Document doc,
            int position,
            out string rawNumber,
            out string headingText)
        {
            rawNumber = string.Empty;
            headingText = string.Empty;
            if (doc == null || position <= 0)
            {
                return false;
            }

            Word.Range searchRange = null;
            Word.Paragraphs paragraphs = null;
            try
            {
                int searchStart = Math.Max(doc.Content.Start, position - 8000);
                searchRange = doc.Range(searchStart, position);
                paragraphs = searchRange.Paragraphs;
                int count = paragraphs?.Count ?? 0;
                int minIndex = Math.Max(1, count - 80);
                for (int i = count; i >= minIndex; i--)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                        if (paragraphRange == null ||
                            paragraphRange.Start >= position ||
                            IsRangeInsideTable(paragraphRange) ||
                            !IsOutlineHeadingParagraph(paragraph))
                        {
                            continue;
                        }

                        string text = NormalizeParagraphText(paragraphRange.Text);
                        if (string.IsNullOrWhiteSpace(text) || LooksLikeTableCaptionText(text))
                        {
                            continue;
                        }

                        string number = TryGetParagraphListNumber(paragraph);
                        if (string.IsNullOrWhiteSpace(number))
                        {
                            string textTitle;
                            TryExtractSectionHeadingForContext(text, out number, out textTitle);
                        }

                        if (string.IsNullOrWhiteSpace(number))
                        {
                            continue;
                        }

                        rawNumber = number;
                        headingText = text;
                        return true;
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(paragraphs);
                ReleaseComObject(searchRange);
            }

            return false;
        }

        private static bool IsOutlineHeadingParagraph(Word.Paragraph paragraph)
        {
            try
            {
                Word.WdOutlineLevel level = paragraph.OutlineLevel;
                return level >= Word.WdOutlineLevel.wdOutlineLevel1 &&
                       level <= Word.WdOutlineLevel.wdOutlineLevel9;
            }
            catch
            {
                return false;
            }
        }

        private static RequirementTrackingDocumentSnapshot CollectRequirementsByParagraphScan(
            Word.Document doc,
            Action<string> progressReporter,
            string preferredPrefix)
        {
            RequirementTrackingDocumentSnapshot snapshot = CreateSnapshot(doc);
            if (doc == null)
            {
                return snapshot;
            }

            List<RequirementItem> requirements = new List<RequirementItem>();
            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = doc.Paragraphs;
                if (paragraphs == null)
                {
                    snapshot.Requirements = requirements;
                    return snapshot;
                }

                string currentSectionNumber = string.Empty;
                int paragraphCount = paragraphs.Count;
                int progressInterval = paragraphCount > 3000 ? 250 : paragraphCount > 800 ? 80 : 25;
                for (int i = 1; i <= paragraphCount; i++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                        string paragraphText = NormalizeParagraphText(paragraphRange?.Text);
                        if (string.IsNullOrWhiteSpace(paragraphText))
                        {
                            continue;
                        }

                        string updatedSectionNumber;
                        if (TryExtractSectionNumber(paragraphText, out updatedSectionNumber))
                        {
                            currentSectionNumber = updatedSectionNumber;
                        }

                        MatchCollection matches = RequirementIdRegex.Matches(paragraphText);
                        if (matches.Count == 0)
                        {
                            continue;
                        }

                        for (int matchIndex = 0; matchIndex < matches.Count; matchIndex++)
                        {
                            Match match = matches[matchIndex];
                            if (!match.Success)
                            {
                                continue;
                            }

                            string id = match.Groups["id"].Value.Trim();
                            if (!IsRequirementIdMatched(id, preferredPrefix))
                            {
                                continue;
                            }

                            string name = ExtractRequirementName(paragraphText, match, matches);
                            if (string.IsNullOrWhiteSpace(name))
                            {
                                name = TryGetSupplementalRequirementName(paragraphs, i, preferredPrefix, id);
                            }

                            requirements.Add(CreateRequirementItem(
                                id,
                                name,
                                currentSectionNumber,
                                paragraphRange?.Start ?? 0));
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                        if (i == 1 || i % progressInterval == 0 || i == paragraphCount)
                        {
                            ReportProgress(progressReporter, $"正在扫描段落：{i}/{paragraphCount}，已识别 {requirements.Count} 项");
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(paragraphs);
            }

            snapshot.Requirements = FinalizeRequirements(requirements);
            ReportProgress(progressReporter, $"解析完成，共识别 {snapshot.Requirements.Count} 项需求");
            return snapshot;
        }

        private static RequirementTrackingDocumentSnapshot CollectSrsThirdChapterRequirements(
            Word.Document doc,
            Action<string> progressReporter)
        {
            return CollectSrsThirdChapterRequirements(doc, progressReporter, null);
        }

        private static RequirementTrackingDocumentSnapshot CollectSrsThirdChapterRequirements(
            Word.Document doc,
            Action<string> progressReporter,
            IReadOnlyList<NavigationPaneEntry> knownMarkerEntries)
        {
            RequirementTrackingDocumentSnapshot snapshot = CreateSnapshot(doc);
            if (doc == null)
            {
                return snapshot;
            }

            List<RequirementItem> requirements = new List<RequirementItem>();
            Word.Range contentRange = null;
            try
            {
                ReportProgress(progressReporter, "正在快速读取需求文档全文...");
                contentRange = doc.Content;
                string rawText = contentRange?.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(rawText))
                {
                    snapshot.Requirements = requirements;
                    return snapshot;
                }

                List<TextLineSnapshot> lines = BuildWordTextLineSnapshots(rawText);
                if (lines.Count == 0)
                {
                    snapshot.Requirements = requirements;
                    return snapshot;
                }

                ApplySectionContext(lines);

                ReportProgress(progressReporter, "正在复用标识窗格范围提取需求标识...");
                List<NavigationPaneEntry> markerEntries = (knownMarkerEntries ?? new List<NavigationPaneEntry>())
                    .Where(item => item != null &&
                                   !string.IsNullOrWhiteSpace(item.Text) &&
                                   IsRequirementIdMatched(item.Text, "SRS"))
                    .ToList();
                if (markerEntries.Count == 0)
                {
                    markerEntries = CollectSrsMarkerEntriesFromMarkerPaneService(doc);
                }

                int contentStart = contentRange.Start;
                markerEntries = markerEntries
                    .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Text))
                    .OrderBy(item => item.Start)
                    .ThenBy(item => item.Text ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                ReportProgress(progressReporter, $"准备解析需求名称：{markerEntries.Count} 个标识");

                HashSet<string> processedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, string> sectionLookupByRequirementId =
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                for (int markerIndex = 0; markerIndex < markerEntries.Count; markerIndex++)
                {
                    NavigationPaneEntry marker = markerEntries[markerIndex];
                    if (marker == null || string.IsNullOrWhiteSpace(marker.Text))
                    {
                        continue;
                    }

                    int lineIndex = FindLineIndexForMarker(lines, contentStart, marker.Start, marker.Text);
                    if (lineIndex < 0)
                    {
                        continue;
                    }

                    string sectionNumber;
                    string sectionTitle;
                    ResolveNearestSectionContext(lines, lineIndex, out sectionNumber, out sectionTitle);

                    string requirementId = marker.Text.Trim();
                    if (!IsRequirementIdMatched(requirementId, "SRS"))
                    {
                        continue;
                    }

                    string normalizedRequirementId = NormalizeRequirementId(requirementId);
                    if (!processedIds.Add(normalizedRequirementId))
                    {
                        continue;
                    }

                    string requirementName = ResolveSrsThirdChapterRequirementName(
                        doc,
                        lines,
                        lineIndex,
                        requirementId,
                        marker.Start,
                        sectionNumber,
                        sectionTitle,
                        out sectionNumber,
                        out sectionTitle);

                    string directSectionNumber;
                    if (!sectionLookupByRequirementId.TryGetValue(requirementId, out directSectionNumber))
                    {
                        directSectionNumber = GetChapterNumberByRequirementId(doc, requirementId);
                        sectionLookupByRequirementId[requirementId] = directSectionNumber;
                    }

                    if (!string.IsNullOrWhiteSpace(directSectionNumber))
                    {
                        sectionNumber = directSectionNumber;
                    }

                    requirements.Add(CreateRequirementItem(
                        requirementId,
                        string.IsNullOrWhiteSpace(requirementName) ? requirementId : requirementName,
                        sectionNumber,
                        marker.Start));

                    if (requirements.Count == 1 || requirements.Count % 20 == 0)
                    {
                        ReportProgress(progressReporter, $"已识别需求：{requirements.Count}/{markerEntries.Count}");
                    }
                }

                if (requirements.Count == 0 && markerEntries.Count > 0)
                {
                    ReportProgress(progressReporter, "已找到标识但未能定位文本上下文，正在生成快速预览结果...");
                    AddMarkerPreviewRequirements(requirements, markerEntries);
                }

            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(contentRange);
            }

            snapshot.Requirements = FinalizeRequirements(requirements);
            ReportProgress(progressReporter, $"需求提取完成，共识别 {snapshot.Requirements.Count} 项需求");
            return snapshot;
        }

        private static List<TextLineSnapshot> BuildWordTextLineSnapshots(string rawText)
        {
            List<TextLineSnapshot> lines = new List<TextLineSnapshot>();
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return lines;
            }

            int lineStart = 0;
            int index = 0;
            while (index <= rawText.Length)
            {
                bool atEnd = index == rawText.Length;
                bool isDelimiter = !atEnd && (rawText[index] == '\r' || rawText[index] == '\n');
                if (!atEnd && !isDelimiter)
                {
                    index++;
                    continue;
                }

                string segment = SafeSubstring(rawText, lineStart, index - lineStart);
                bool nextIsCellMarker = !atEnd &&
                                        rawText[index] == '\r' &&
                                        index + 1 < rawText.Length &&
                                        rawText[index + 1] == '\a';
                bool isTableCell = segment.IndexOf('\a') >= 0 || nextIsCellMarker;
                string text = NormalizeParagraphText(segment);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    lines.Add(new TextLineSnapshot
                    {
                        StartIndex = lineStart,
                        Text = text,
                        IsTableCell = isTableCell
                    });
                }

                if (!atEnd && rawText[index] == '\r' && index + 1 < rawText.Length && rawText[index + 1] == '\n')
                {
                    index++;
                }

                lineStart = index + 1;
                index++;
            }

            return lines;
        }

        private static bool TryFindSrsThirdChapterLineBounds(
            IReadOnlyList<TextLineSnapshot> lines,
            out int chapterStartLine,
            out int chapterEndLine)
        {
            chapterStartLine = -1;
            chapterEndLine = -1;
            if (lines == null || lines.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < lines.Count; i++)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (LooksLikeSrsThirdChapterStart(text))
                {
                    chapterStartLine = i;
                    break;
                }
            }

            if (chapterStartLine < 0)
            {
                for (int i = 0; i < lines.Count; i++)
                {
                    string sectionNumber;
                    string sectionTitle;
                    if (TryExtractSectionHeading(lines[i]?.Text, out sectionNumber, out sectionTitle) &&
                        IsSectionNumberUnder(sectionNumber, "3"))
                    {
                        chapterStartLine = i;
                        break;
                    }
                }
            }

            if (chapterStartLine < 0)
            {
                return false;
            }

            chapterEndLine = lines.Count - 1;
            for (int i = chapterStartLine + 1; i < lines.Count; i++)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (LooksLikeNextTopLevelChapter(text))
                {
                    chapterEndLine = i - 1;
                    break;
                }
            }

            return chapterEndLine >= chapterStartLine;
        }

        private static bool TryLocateSrsThirdChapterBoundsFromLines(
            IReadOnlyList<TextLineSnapshot> lines,
            int rawTextLength,
            int contentStart,
            int anchorStart,
            out int chapterStart,
            out int chapterEnd)
        {
            chapterStart = 0;
            chapterEnd = 0;

            int chapterStartLine;
            int chapterEndLine;
            if (!TryFindSrsThirdChapterLineBounds(lines, out chapterStartLine, out chapterEndLine) ||
                chapterStartLine < 0 ||
                chapterStartLine >= (lines?.Count ?? 0))
            {
                return false;
            }

            int anchorRawStart = anchorStart > contentStart ? anchorStart - contentStart : -1;
            if (anchorRawStart >= 0)
            {
                int anchorLine = FindLineIndexByRawPosition(lines, anchorRawStart);
                bool currentBoundsContainAnchor = anchorLine >= chapterStartLine && anchorLine <= chapterEndLine;
                if (!currentBoundsContainAnchor && anchorLine >= 0)
                {
                    int correctedStartLine = FindNearestSrsThirdChapterStartBefore(lines, anchorLine);
                    if (correctedStartLine >= 0)
                    {
                        chapterStartLine = correctedStartLine;
                        chapterEndLine = FindSrsThirdChapterEndLine(lines, chapterStartLine);
                    }
                }
            }

            int rawStart = Math.Max(0, lines[chapterStartLine]?.StartIndex ?? 0);
            int rawEnd = Math.Max(rawStart, rawTextLength);
            int nextLineIndex = chapterEndLine + 1;
            if (lines != null && nextLineIndex >= 0 && nextLineIndex < lines.Count)
            {
                rawEnd = Math.Max(rawStart, lines[nextLineIndex]?.StartIndex ?? rawEnd);
            }

            chapterStart = contentStart + rawStart;
            chapterEnd = contentStart + rawEnd;
            return chapterEnd > chapterStart;
        }

        private static int FindNearestSrsThirdChapterStartBefore(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex)
        {
            if (lines == null || lineIndex < 0)
            {
                return -1;
            }

            for (int i = Math.Min(lineIndex, lines.Count - 1); i >= 0; i--)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (LooksLikeSrsThirdChapterStart(text))
                {
                    return i;
                }

                string sectionNumber;
                string sectionTitle;
                if (TryExtractSectionHeading(text, out sectionNumber, out sectionTitle) &&
                    IsSectionNumberUnder(sectionNumber, "3"))
                {
                    return i;
                }
            }

            return -1;
        }

        private static int FindSrsThirdChapterEndLine(
            IReadOnlyList<TextLineSnapshot> lines,
            int chapterStartLine)
        {
            if (lines == null || chapterStartLine < 0)
            {
                return -1;
            }

            int chapterEndLine = lines.Count - 1;
            for (int i = chapterStartLine + 1; i < lines.Count; i++)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (LooksLikeNextTopLevelChapter(text))
                {
                    return i - 1;
                }
            }

            return chapterEndLine;
        }

        private static int GetFirstRequirementStart(IReadOnlyList<RequirementItem> requirements)
        {
            if (requirements == null)
            {
                return 0;
            }

            RequirementItem first = requirements
                .Where(item => item != null && item.Start > 0)
                .OrderBy(item => item.Start)
                .FirstOrDefault();
            return first?.Start ?? 0;
        }

        private static Dictionary<string, string> BuildSectionLookupByTitle(IReadOnlyList<TextLineSnapshot> lines)
        {
            Dictionary<string, string> lookup = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (lines == null)
            {
                return lookup;
            }

            for (int i = 0; i < lines.Count; i++)
            {
                string sectionNumber;
                string sectionTitle;
                if (!TryExtractSectionHeading(lines[i]?.Text, out sectionNumber, out sectionTitle))
                {
                    continue;
                }

                sectionNumber = TruncateSectionNumberToDepth(sectionNumber, 3);
                string key = NormalizeSectionTitleLookupKey(sectionTitle);
                if (string.IsNullOrWhiteSpace(sectionNumber) ||
                    string.IsNullOrWhiteSpace(key) ||
                    lookup.ContainsKey(key))
                {
                    continue;
                }

                lookup[key] = sectionNumber;
            }

            return lookup;
        }

        private static string FindSectionNumberByTitle(
            IReadOnlyDictionary<string, string> lookup,
            string title)
        {
            if (lookup == null || string.IsNullOrWhiteSpace(title))
            {
                return string.Empty;
            }

            string key = NormalizeSectionTitleLookupKey(title);
            if (string.IsNullOrWhiteSpace(key))
            {
                return string.Empty;
            }

            string sectionNumber;
            if (lookup.TryGetValue(key, out sectionNumber))
            {
                return sectionNumber;
            }

            KeyValuePair<string, string> partialMatch = lookup
                .Where(item => IsConfidentTitleMatch(key, item.Key))
                .OrderByDescending(item => item.Key.Length)
                .FirstOrDefault();
            return partialMatch.Equals(default(KeyValuePair<string, string>)) ? string.Empty : partialMatch.Value;
        }

        private static bool IsConfidentTitleMatch(string requirementNameKey, string headingTitleKey)
        {
            if (string.IsNullOrWhiteSpace(requirementNameKey) || string.IsNullOrWhiteSpace(headingTitleKey))
            {
                return false;
            }

            if (string.Equals(requirementNameKey, headingTitleKey, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            int shorterLength = Math.Min(requirementNameKey.Length, headingTitleKey.Length);
            if (shorterLength < 6)
            {
                return false;
            }

            return requirementNameKey.StartsWith(headingTitleKey, StringComparison.OrdinalIgnoreCase) ||
                   headingTitleKey.StartsWith(requirementNameKey, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeSectionTitleLookupKey(string title)
        {
            string normalized = NormalizeParagraphText(title);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            normalized = Regex.Replace(normalized, @"\.{2,}\s*\d+\s*$", string.Empty);
            normalized = Regex.Replace(normalized, @"\s+\d{1,4}\s*$", string.Empty);
            normalized = RemoveRequirementIds(normalized);
            normalized = Regex.Replace(normalized, @"\s+", string.Empty);
            return normalized.Trim();
        }

        private static void ApplySectionContext(IReadOnlyList<TextLineSnapshot> lines)
        {
            if (lines == null)
            {
                return;
            }

            string currentSectionNumber = string.Empty;
            string currentSectionTitle = string.Empty;
            for (int i = 0; i < lines.Count; i++)
            {
                TextLineSnapshot line = lines[i];
                if (line == null)
                {
                    continue;
                }

                string sectionNumber;
                string sectionTitle;
                if (TryExtractSectionHeadingForContext(line.Text, out sectionNumber, out sectionTitle))
                {
                    currentSectionNumber = sectionNumber;
                    currentSectionTitle = sectionTitle;
                }

                line.SectionNumber = currentSectionNumber;
                line.SectionTitle = currentSectionTitle;
            }
        }

        private static void ResolveNearestSectionContext(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            out string sectionNumber,
            out string sectionTitle)
        {
            sectionNumber = string.Empty;
            sectionTitle = string.Empty;
            if (lines == null || lineIndex < 0)
            {
                return;
            }

            int start = Math.Min(lineIndex, lines.Count - 1);
            sectionNumber = lines[start]?.SectionNumber ?? string.Empty;
            sectionTitle = lines[start]?.SectionTitle ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(sectionNumber))
            {
                return;
            }

            int stop = Math.Max(0, start - 160);
            for (int i = start; i >= stop; i--)
            {
                if (TryExtractSectionHeadingForContext(lines[i]?.Text, out sectionNumber, out sectionTitle))
                {
                    return;
                }
            }
        }

        private static bool TryExtractSectionHeadingForContext(
            string text,
            out string sectionNumber,
            out string sectionTitle)
        {
            sectionNumber = string.Empty;
            sectionTitle = string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            if (LooksLikeTocEntryText(text))
            {
                return false;
            }

            if (TryExtractSectionHeading(text, out sectionNumber, out sectionTitle) &&
                !LooksLikeTocHeadingTitle(sectionTitle))
            {
                sectionNumber = TruncateSectionNumberToDepth(sectionNumber, 3);
                return true;
            }

            string stripped = StripLeadingListMarkerForHeading(text);
            if (!string.Equals(stripped, text, StringComparison.Ordinal) &&
                TryExtractSectionHeading(stripped, out sectionNumber, out sectionTitle) &&
                !LooksLikeTocHeadingTitle(sectionTitle))
            {
                sectionNumber = TruncateSectionNumberToDepth(sectionNumber, 3);
                return true;
            }

            return false;
        }

        private static List<NavigationPaneEntry> CollectSrsMarkerEntriesFromMarkerPaneService(Word.Document doc)
        {
            List<NavigationPaneEntry> entries = new List<NavigationPaneEntry>();
            try
            {
                DocumentMarkerCollectionResult markerResult = DocumentMarkerService.CollectMarkers(doc);
                if (markerResult?.Entries == null)
                {
                    return entries;
                }

                entries.AddRange(markerResult.Entries
                    .Where(item => item != null &&
                                   !string.IsNullOrWhiteSpace(item.Text) &&
                                   IsRequirementIdMatched(item.Text, "SRS")));
            }
            catch
            {
            }

            return entries
                .OrderBy(item => item.Start)
                .ThenBy(item => item.Text ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static List<NavigationPaneEntry> CollectSrsMarkerEntriesFromTextLines(
            IReadOnlyList<TextLineSnapshot> lines,
            int chapterStartLine,
            int chapterEndLine,
            int contentStart)
        {
            List<NavigationPaneEntry> entries = new List<NavigationPaneEntry>();
            if (lines == null || lines.Count == 0)
            {
                return entries;
            }

            int start = Math.Max(0, chapterStartLine);
            int end = Math.Min(lines.Count - 1, chapterEndLine);
            for (int i = start; i <= end; i++)
            {
                TextLineSnapshot line = lines[i];
                if (line == null || string.IsNullOrWhiteSpace(line.Text))
                {
                    continue;
                }

                MatchCollection matches = RequirementIdRegex.Matches(line.Text);
                foreach (Match match in matches)
                {
                    if (!match.Success)
                    {
                        continue;
                    }

                    string requirementId = match.Groups["id"].Value.Trim();
                    if (!IsRequirementIdMatched(requirementId, "SRS"))
                    {
                        continue;
                    }

                    entries.Add(new NavigationPaneEntry
                    {
                        Start = contentStart + line.StartIndex + match.Index,
                        Text = requirementId
                    });
                }
            }

            return entries
                .OrderBy(item => item.Start)
                .ThenBy(item => item.Text ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static int CountSrsMarkersInDocumentText(Word.Document doc)
        {
            if (doc == null)
            {
                return 0;
            }

            Word.Range contentRange = null;
            try
            {
                contentRange = doc.Content;
                string rawText = contentRange?.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(rawText))
                {
                    return 0;
                }

                List<TextLineSnapshot> lines = BuildWordTextLineSnapshots(rawText);
                return CollectSrsMarkerEntriesFromTextLines(lines, 0, lines.Count - 1, contentRange.Start).Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ReleaseComObject(contentRange);
            }
        }

        private static void AddMarkerPreviewRequirements(
            List<RequirementItem> requirements,
            IReadOnlyList<NavigationPaneEntry> markerEntries)
        {
            if (requirements == null || markerEntries == null)
            {
                return;
            }

            HashSet<string> existingIds = new HashSet<string>(
                requirements
                    .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                    .Select(item => NormalizeRequirementId(item.Id)),
                StringComparer.OrdinalIgnoreCase);

            foreach (NavigationPaneEntry marker in markerEntries)
            {
                if (marker == null || string.IsNullOrWhiteSpace(marker.Text))
                {
                    continue;
                }

                string normalizedId = NormalizeRequirementId(marker.Text);
                if (string.IsNullOrWhiteSpace(normalizedId) || !existingIds.Add(normalizedId))
                {
                    continue;
                }

                requirements.Add(CreateRequirementItem(
                    marker.Text,
                    ExtractRequirementIdSuffix(marker.Text, "SRS"),
                    string.Empty,
                    marker.Start));
            }
        }

        private static int FindLineIndexForMarker(
            IReadOnlyList<TextLineSnapshot> lines,
            int contentStart,
            int markerStart,
            string markerText)
        {
            if (lines == null || lines.Count == 0)
            {
                return -1;
            }

            int rawMarkerStart = Math.Max(0, markerStart - contentStart);
            int estimatedIndex = FindLineIndexByRawPosition(lines, rawMarkerStart);
            if (estimatedIndex >= 0 && LineContainsMarker(lines[estimatedIndex], markerText))
            {
                return estimatedIndex;
            }

            int nearbyStart = Math.Max(0, estimatedIndex - 8);
            int nearbyEnd = Math.Min(lines.Count - 1, estimatedIndex + 8);
            for (int i = nearbyStart; i <= nearbyEnd; i++)
            {
                if (LineContainsMarker(lines[i], markerText))
                {
                    return i;
                }
            }

            for (int i = 0; i < lines.Count; i++)
            {
                if (LineContainsMarker(lines[i], markerText))
                {
                    return i;
                }
            }

            return estimatedIndex;
        }

        private static int FindLineIndexByRawPosition(IReadOnlyList<TextLineSnapshot> lines, int rawPosition)
        {
            if (lines == null || lines.Count == 0)
            {
                return -1;
            }

            int left = 0;
            int right = lines.Count - 1;
            int best = 0;
            while (left <= right)
            {
                int middle = left + ((right - left) / 2);
                int start = lines[middle]?.StartIndex ?? 0;
                if (start <= rawPosition)
                {
                    best = middle;
                    left = middle + 1;
                }
                else
                {
                    right = middle - 1;
                }
            }

            return best;
        }

        private static bool LineContainsMarker(TextLineSnapshot line, string markerText)
        {
            if (line == null || string.IsNullOrWhiteSpace(line.Text) || string.IsNullOrWhiteSpace(markerText))
            {
                return false;
            }

            string normalizedLine = NormalizeRequirementId(line.Text);
            string normalizedMarker = NormalizeRequirementId(markerText);
            if (!string.IsNullOrWhiteSpace(normalizedMarker) &&
                normalizedLine.IndexOf(normalizedMarker, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            string shortMarker = ExtractRequirementIdSuffix(normalizedMarker, "SRS");
            return !string.IsNullOrWhiteSpace(shortMarker) &&
                   normalizedLine.IndexOf(shortMarker, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string ResolveSrsThirdChapterRequirementName(
            Word.Document doc,
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            string requirementId,
            int markerStart,
            string initialSectionNumber,
            string initialSectionTitle,
            out string resolvedSectionNumber,
            out string resolvedSectionTitle)
        {
            resolvedSectionNumber = initialSectionNumber ?? string.Empty;
            resolvedSectionTitle = initialSectionTitle ?? string.Empty;
            string candidate = string.Empty;

            bool markerInsideTable = IsMarkerPositionInsideTable(doc, markerStart, lines, lineIndex);
            if (markerInsideTable)
            {
                candidate = TryResolveNameFromMarkerTableFirstTwoColumns(doc, markerStart, requirementId);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            if (!markerInsideTable)
            {
                string localSectionNumber;
                string localSectionTitle;
                bool hasLocalHeading = TryFindNearestHeadingBeforePosition(
                    doc,
                    markerStart,
                    out localSectionNumber,
                    out localSectionTitle);
                if (hasLocalHeading)
                {
                    if (!string.IsNullOrWhiteSpace(localSectionNumber))
                    {
                        resolvedSectionNumber = localSectionNumber;
                    }

                    if (!string.IsNullOrWhiteSpace(localSectionTitle))
                    {
                        resolvedSectionTitle = localSectionTitle;
                    }
                }

                if (!markerInsideTable)
                {
                    candidate = hasLocalHeading ? localSectionTitle : string.Empty;

                    if (string.IsNullOrWhiteSpace(candidate))
                    {
                        candidate = FindNearestRequirementTitleBefore(lines, lineIndex);
                    }

                    if (string.IsNullOrWhiteSpace(candidate))
                    {
                        candidate = LooksLikeTocHeadingTitle(resolvedSectionTitle)
                            ? string.Empty
                            : NormalizeRequirementNameCell(resolvedSectionTitle, false);
                    }

                    if (!string.IsNullOrWhiteSpace(candidate))
                    {
                        return candidate;
                    }
                }
            }

            if (IsSrsStatusRequirementSection(resolvedSectionNumber, resolvedSectionTitle))
            {
                candidate = FindNearestSubsectionTitleBefore(lines, lineIndex, "3.1");
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    candidate = LooksLikeTocHeadingTitle(resolvedSectionTitle)
                        ? string.Empty
                        : NormalizeRequirementNameCell(resolvedSectionTitle, false);
                }

                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            if (!markerInsideTable &&
                IsSrsSafetyOrCriticalRequirementContext(lines, lineIndex, resolvedSectionNumber))
            {
                candidate = TryGetSecondColumnNameFromTableRow(lines, lineIndex);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            if (!markerInsideTable &&
                IsSectionNumberUnder(resolvedSectionNumber, "3.2"))
            {
                candidate = TryGetNameAfterTableLabel(lines, lineIndex, "能力名称", "需求名称");
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    candidate = TryGetFirstRowSecondCellName(lines, lineIndex);
                }

                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            if (!markerInsideTable &&
                IsSectionNumberUnder(resolvedSectionNumber, "3.3"))
            {
                candidate = TryGetNameByHeaderOffset(lines, lineIndex, "接口名称", "能力名称", "需求名称");
                if (string.IsNullOrWhiteSpace(candidate))
                {
                    candidate = TryGetNameAfterTableLabel(lines, lineIndex, "能力名称", "需求名称", "接口名称");
                }

                if (string.IsNullOrWhiteSpace(candidate))
                {
                    candidate = TryGetNearbyShortNameCell(lines, lineIndex);
                }

                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            candidate = markerInsideTable ? string.Empty : TryGetGenericTableRequirementName(lines, lineIndex);
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            candidate = ExtractNameFromText(lines[lineIndex]?.Text, requirementId);
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            candidate = FindNearestRequirementTitleBefore(lines, lineIndex);
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            candidate = NormalizeRequirementNameCell(resolvedSectionTitle, false);
            return string.IsNullOrWhiteSpace(candidate) ? requirementId : candidate;
        }

        private static string TryResolveNameFromMarkerTableFirstTwoColumns(
            Word.Document doc,
            int markerStart,
            string requirementId)
        {
            if (doc == null || markerStart <= 0)
            {
                return string.Empty;
            }

            Word.Range markerRange = null;
            Word.Cells cells = null;
            Word.Cell cell = null;
            Word.Row row = null;
            try
            {
                int end = Math.Min(doc.Content.End, markerStart + 1);
                markerRange = doc.Range(markerStart, end);
                if (markerRange == null)
                {
                    return string.Empty;
                }

                cells = markerRange.Cells;
                if (cells == null || cells.Count <= 0)
                {
                    return string.Empty;
                }

                cell = cells[1];
                row = cell?.Row;
                if (row == null)
                {
                    return string.Empty;
                }

                List<string> rowCells = GetRowCellTexts(row);
                if (rowCells.Count <= 0)
                {
                    return string.Empty;
                }

                int idCellIndex = FindCellIndexContainingId(rowCells, requirementId);
                int maxNameColumnIndex = Math.Min(1, rowCells.Count - 1);
                for (int i = 0; i <= maxNameColumnIndex; i++)
                {
                    if (i == idCellIndex)
                    {
                        continue;
                    }

                    string candidate = NormalizeRequirementNameCell(rowCells[i], true);
                    if (!string.IsNullOrWhiteSpace(candidate))
                    {
                        return candidate;
                    }
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(row);
                ReleaseComObject(cell);
                ReleaseComObject(cells);
                ReleaseComObject(markerRange);
            }
        }

        private static bool IsLineInsideTable(IReadOnlyList<TextLineSnapshot> lines, int lineIndex)
        {
            return lines != null &&
                   lineIndex >= 0 &&
                   lineIndex < lines.Count &&
                   lines[lineIndex] != null &&
                   lines[lineIndex].IsTableCell;
        }

        private static bool IsMarkerPositionInsideTable(
            Word.Document doc,
            int markerStart,
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex)
        {
            if (doc != null && markerStart > 0)
            {
                Word.Range markerRange = null;
                try
                {
                    markerRange = doc.Range(markerStart, Math.Min(markerStart + 1, doc.Content.End));
                    return IsRangeInsideTable(markerRange);
                }
                catch
                {
                }
                finally
                {
                    ReleaseComObject(markerRange);
                }
            }

            return IsLineInsideTable(lines, lineIndex);
        }

        private static bool TryFindNearestHeadingBeforePosition(
            Word.Document doc,
            int markerStart,
            out string sectionNumber,
            out string sectionTitle)
        {
            sectionNumber = string.Empty;
            sectionTitle = string.Empty;
            if (doc == null || markerStart <= 0)
            {
                return false;
            }

            Word.Range searchRange = null;
            Word.Paragraphs paragraphs = null;
            try
            {
                int searchStart = Math.Max(doc.Content.Start, markerStart - 5000);
                searchRange = doc.Range(searchStart, markerStart);
                paragraphs = searchRange?.Paragraphs;
                if (paragraphs == null)
                {
                    return false;
                }

                for (int i = paragraphs.Count; i >= 1; i--)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                        if (paragraphRange == null || IsRangeInsideTable(paragraphRange))
                        {
                            continue;
                        }

                        string text = NormalizeParagraphText(paragraphRange.Text);
                        string candidateNumber;
                        string candidateTitle;
                        if (TryExtractHeadingFromParagraph(paragraph, text, out candidateNumber, out candidateTitle))
                        {
                            sectionNumber = TruncateSectionNumberToDepth(candidateNumber, 3);
                            sectionTitle = candidateTitle;
                            return true;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(paragraphs);
                ReleaseComObject(searchRange);
            }

            return false;
        }

        private static string GetChapterNumberByRequirementId(Word.Document doc, string requirementId)
        {
            if (doc == null || string.IsNullOrWhiteSpace(requirementId))
            {
                return string.Empty;
            }

            Word.Range searchRange = null;
            Word.Find find = null;
            Word.Range headingRange = null;
            try
            {
                searchRange = doc.Content;
                find = searchRange.Find;
                if (find == null)
                {
                    return string.Empty;
                }

                find.ClearFormatting();
                find.Text = requirementId;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;

                if (!find.Execute())
                {
                    string shortId = ExtractRequirementIdSuffix(requirementId, "SRS");
                    if (string.IsNullOrWhiteSpace(shortId) || string.Equals(shortId, requirementId, StringComparison.OrdinalIgnoreCase))
                    {
                        return string.Empty;
                    }

                    searchRange = doc.Content;
                    find = searchRange.Find;
                    find.ClearFormatting();
                    find.Text = shortId;
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = false;
                    find.MatchCase = false;
                    find.MatchWholeWord = false;
                    find.MatchWildcards = false;
                    if (!find.Execute())
                    {
                        return string.Empty;
                    }
                }

                int headingProbeStart = GetHeadingProbeStart(searchRange);
                string resolvedHeadingText;
                string resolvedListString;
                headingRange = GoToPreviousNumberedHeadingFromPosition(
                    doc,
                    headingProbeStart,
                    out resolvedListString,
                    out resolvedHeadingText);
                if (headingRange == null)
                {
                    return string.Empty;
                }

                string rawListString = resolvedListString;
                try
                {
                    if (string.IsNullOrWhiteSpace(rawListString))
                    {
                        rawListString = headingRange.ListFormat?.ListString ?? string.Empty;
                    }
                }
                catch
                {
                    rawListString = resolvedListString;
                }

                if (string.IsNullOrWhiteSpace(rawListString))
                {
                    Word.Paragraph paragraph = null;
                    try
                    {
                        headingRange.Expand(Word.WdUnits.wdParagraph);
                        paragraph = headingRange.Paragraphs[1];
                        rawListString = TryGetParagraphListNumber(paragraph);
                    }
                    finally
                    {
                        ReleaseComObject(paragraph);
                    }
                }

                return TruncateSectionNumberToDepth(rawListString, 3);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(headingRange);
                ReleaseComObject(find);
                ReleaseComObject(searchRange);
            }
        }

        private static bool TryFindNearestListHeadingBeforePosition(
            Word.Document doc,
            int position,
            out string sectionNumber,
            out string sectionTitle)
        {
            sectionNumber = string.Empty;
            sectionTitle = string.Empty;
            if (doc == null || position <= 0)
            {
                return false;
            }

            Word.Range searchRange = null;
            Word.Paragraphs paragraphs = null;
            try
            {
                int searchStart = Math.Max(doc.Content.Start, position - 2500);
                searchRange = doc.Range(searchStart, position);
                paragraphs = searchRange?.Paragraphs;
                if (paragraphs == null)
                {
                    return false;
                }

                for (int i = paragraphs.Count; i >= 1; i--)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                        if (paragraphRange == null || IsRangeInsideTable(paragraphRange))
                        {
                            continue;
                        }

                        string listNumber = TryGetParagraphListNumber(paragraph);
                        if (string.IsNullOrWhiteSpace(listNumber))
                        {
                            continue;
                        }

                        int[] parts = ParseSectionNumberParts(listNumber);
                        if (parts.Length < 3)
                        {
                            continue;
                        }

                        string text = NormalizeParagraphText(paragraphRange.Text);
                        string title = TrimRequirementNameFragment(text);
                        if (string.IsNullOrWhiteSpace(title) ||
                            ContainsRequirementId(title, "SRS") ||
                            LooksLikeTocHeadingTitle(title))
                        {
                            continue;
                        }

                        sectionNumber = TruncateSectionNumberToDepth(listNumber, 3);
                        sectionTitle = title;
                        return true;
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                    }
                }
            }
            finally
            {
                ReleaseComObject(paragraphs);
                ReleaseComObject(searchRange);
            }

            return false;
        }

        private static bool TryExtractHeadingFromParagraph(
            Word.Paragraph paragraph,
            string text,
            out string sectionNumber,
            out string sectionTitle)
        {
            sectionNumber = string.Empty;
            sectionTitle = string.Empty;
            string textNumber;
            string textTitle;
            if (TryExtractSectionHeadingForContext(text, out textNumber, out textTitle))
            {
                sectionNumber = textNumber;
                sectionTitle = textTitle;
                return true;
            }

            string listNumber = TryGetParagraphListNumber(paragraph);
            string title = TrimRequirementNameFragment(text);
            bool hasHeadingOutline = false;

            try
            {
                Word.WdOutlineLevel level = paragraph.OutlineLevel;
                hasHeadingOutline = level != Word.WdOutlineLevel.wdOutlineLevelBodyText;
            }
            catch
            {
                hasHeadingOutline = false;
            }

            if (string.IsNullOrWhiteSpace(listNumber))
            {
                Match simpleMatch = SimpleTopLevelHeadingRegex.Match(text);
                if (simpleMatch.Success)
                {
                    listNumber = simpleMatch.Groups["number"].Value.Trim();
                    title = TrimRequirementNameFragment(simpleMatch.Groups["title"].Value);
                }
            }
            
            Match chapterMatch = ChapterHeadingRegex.Match(text);
            if (chapterMatch.Success)
            {
                listNumber = chapterMatch.Groups["number"].Value.Trim();
                string titlePart = text.Substring(chapterMatch.Length);
                title = TrimRequirementNameFragment(titlePart);
            }

            if (string.IsNullOrWhiteSpace(listNumber) && !hasHeadingOutline)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(title) ||
                ContainsRequirementId(title, "SRS") ||
                LooksLikeTocHeadingTitle(title))
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(listNumber))
            {
                sectionNumber = TruncateSectionNumberToDepth(listNumber, 3);
            }

            sectionTitle = title;
            return true;
        }

        private static string TryGetParagraphListNumber(Word.Paragraph paragraph)
        {
            if (paragraph == null)
            {
                return string.Empty;
            }

            try
            {
                string listString = paragraph.Range?.ListFormat?.ListString ?? string.Empty;
                listString = NormalizeParagraphText(listString).Trim(' ', '\t', '.', '．', '、');
                return Regex.IsMatch(listString, @"^\d+(?:[.．]\d+)*$")
                    ? listString.Replace('．', '.')
                    : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool IsStandaloneRequirementMarkerLine(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            string requirementId)
        {
            if (lines == null || lineIndex < 0 || lineIndex >= lines.Count)
            {
                return false;
            }

            string text = lines[lineIndex]?.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text) || !ContainsRequirementId(text, "SRS"))
            {
                return false;
            }

            string cleaned = RemoveRequirementIds(text);
            if (string.IsNullOrWhiteSpace(cleaned))
            {
                return true;
            }

            cleaned = cleaned.Trim('：', ':', '，', ',', '。', '.', '；', ';', '、', '-', '—', '－', ' ', '\t', '[', ']', '【', '】', '(', ')', '（', '）');
            return string.IsNullOrWhiteSpace(cleaned) ||
                   ContainsAny(cleaned, "需求标识", "唯一标识", "标识", "编号");
        }

        private static string TryGetGenericTableRequirementName(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex)
        {
            string candidate = TryGetNameByHeaderOffset(lines, lineIndex, "接口名称", "能力名称", "需求名称");
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            candidate = TryGetNameAfterTableLabel(lines, lineIndex, "能力名称", "需求名称", "接口名称");
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            candidate = TryGetSecondColumnNameFromTableRow(lines, lineIndex);
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            return TryGetNearbyShortNameCell(lines, lineIndex);
        }

        private static bool IsSrsStatusRequirementSection(string sectionNumber, string sectionTitle)
        {
            return IsSectionNumberUnder(sectionNumber, "3.1") ||
                   ContainsAny(sectionTitle, "要求的状态和方式");
        }

        private static string FindNearestSubsectionTitleBefore(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            string parentSectionNumber)
        {
            if (lines == null || lineIndex < 0 || string.IsNullOrWhiteSpace(parentSectionNumber))
            {
                return string.Empty;
            }

            int start = Math.Min(lineIndex, lines.Count - 1);
            int stop = Math.Max(0, start - 12);
            for (int i = start; i >= stop; i--)
            {
                string sectionNumber;
                string sectionTitle;
                if (!TryExtractStatusSubsectionHeading(lines[i]?.Text, parentSectionNumber, out sectionNumber, out sectionTitle))
                {
                    if (i < start && ContainsRequirementId(lines[i]?.Text, "SRS"))
                    {
                        break;
                    }

                    continue;
                }

                if (!IsSectionNumberUnder(sectionNumber, parentSectionNumber))
                {
                    break;
                }

                if (string.Equals(sectionNumber, parentSectionNumber, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string candidate = NormalizeRequirementNameCell(sectionTitle, false);
                if (!string.IsNullOrWhiteSpace(candidate) && !LooksLikeTocHeadingTitle(candidate))
                {
                    return candidate;
                }
            }

            return string.Empty;
        }

        private static string FindNearestRequirementTitleBefore(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex)
        {
            if (lines == null || lineIndex < 0)
            {
                return string.Empty;
            }

            int start = Math.Min(lineIndex, lines.Count - 1);
            int stop = Math.Max(0, start - 80);
            for (int i = start; i >= stop; i--)
            {
                string sectionNumber;
                string sectionTitle;
                if (!TryExtractSectionHeadingForContext(lines[i]?.Text, out sectionNumber, out sectionTitle))
                {
                    if (i < start && ContainsRequirementId(lines[i]?.Text, "SRS"))
                    {
                        break;
                    }

                    continue;
                }

                string candidate = NormalizeRequirementNameCell(sectionTitle, false);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            return string.Empty;
        }

        private static bool TryExtractStatusSubsectionHeading(
            string text,
            string parentSectionNumber,
            out string sectionNumber,
            out string sectionTitle)
        {
            sectionNumber = string.Empty;
            sectionTitle = string.Empty;
            if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(parentSectionNumber))
            {
                return false;
            }

            if (TryExtractSectionHeading(text, out sectionNumber, out sectionTitle) &&
                IsSectionNumberUnder(sectionNumber, parentSectionNumber))
            {
                return true;
            }

            string stripped = StripLeadingListMarkerForHeading(text);
            if (string.Equals(stripped, text, StringComparison.Ordinal))
            {
                return false;
            }

            return TryExtractSectionHeading(stripped, out sectionNumber, out sectionTitle) &&
                   IsSectionNumberUnder(sectionNumber, parentSectionNumber);
        }

        private static bool LooksLikeTocHeadingTitle(string title)
        {
            string normalized = NormalizeParagraphText(title);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return true;
            }

            if (Regex.IsMatch(normalized, @"\.{2,}\s*\d+\s*$"))
            {
                return true;
            }

            if (Regex.IsMatch(normalized, @"\s\d{1,4}\s*$"))
            {
                return true;
            }

            return false;
        }

        private static bool LooksLikeTocEntryText(string text)
        {
            string normalized = NormalizeParagraphText(text);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            if (Regex.IsMatch(normalized, @"\.{2,}\s*\d{1,4}\s*$"))
            {
                return true;
            }

            bool hasTrailingPageNumber = Regex.IsMatch(normalized, @"\s\d{1,4}\s*$");
            if (!hasTrailingPageNumber)
            {
                return false;
            }

            return NumberedHeadingRegex.IsMatch(normalized) ||
                   SimpleTopLevelHeadingRegex.IsMatch(normalized) ||
                   ChapterHeadingRegex.IsMatch(normalized);
        }

        private static string StripLeadingListMarkerForHeading(string text)
        {
            string normalized = NormalizeParagraphText(text);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return string.Empty;
            }

            int firstDigitIndex = -1;
            for (int i = 0; i < normalized.Length; i++)
            {
                if (char.IsDigit(normalized[i]))
                {
                    firstDigitIndex = i;
                    break;
                }
            }

            if (firstDigitIndex <= 0)
            {
                return normalized;
            }

            string prefix = normalized.Substring(0, firstDigitIndex).Trim();
            if (prefix.Any(ch => char.IsLetterOrDigit(ch) || IsCjkCharacter(ch)))
            {
                return normalized;
            }

            return normalized.Substring(firstDigitIndex).TrimStart();
        }

        private static bool IsCjkCharacter(char ch)
        {
            return ch >= '\u4e00' && ch <= '\u9fff';
        }

        private static bool IsSrsSafetyOrCriticalRequirementContext(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            string sectionNumber)
        {
            // Do not route by section number alone (e.g. 3.17) to avoid overriding
            // the normal name resolution path for non-safety/non-reliability tables.
            return TableContextContains(lines, lineIndex, "安全性需求", "安全性", "关键性需求", "关键需求", "关键性", "可靠性需求", "可靠性");
        }

        private static bool TableContextContains(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            params string[] keywords)
        {
            int tableStart;
            int tableEnd;
            if (!TryGetTableCellBounds(lines, lineIndex, out tableStart, out tableEnd))
            {
                return false;
            }

            int start = Math.Max(0, tableStart - 5);
            int end = Math.Min(lines.Count - 1, tableEnd);
            for (int i = start; i <= end; i++)
            {
                if (ContainsAny(lines[i]?.Text, keywords))
                {
                    return true;
                }
            }

            return false;
        }

        private static string TryGetSecondColumnNameFromTableRow(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex)
        {
            int tableStart;
            int tableEnd;
            if (!TryGetTableCellBounds(lines, lineIndex, out tableStart, out tableEnd))
            {
                return string.Empty;
            }

            int idHeaderIndex;
            int nameHeaderIndex;
            if (TryFindIdentifierAndNameHeaders(lines, tableStart, tableEnd, lineIndex, out idHeaderIndex, out nameHeaderIndex))
            {
                int nameOffset = nameHeaderIndex - idHeaderIndex;
                int minColumnCount = Math.Max(2, nameOffset + 1);
                for (int columnCount = minColumnCount; columnCount <= 8; columnCount++)
                {
                    int offsetFromHeader = lineIndex - idHeaderIndex;
                    if (offsetFromHeader <= 0 || offsetFromHeader % columnCount != 0)
                    {
                        continue;
                    }

                    int candidateIndex = lineIndex + nameOffset;
                    if (candidateIndex <= lineIndex || candidateIndex > tableEnd)
                    {
                        continue;
                    }

                    string candidate = NormalizeRequirementNameCell(lines[candidateIndex]?.Text, true);
                    if (!string.IsNullOrWhiteSpace(candidate))
                    {
                        return candidate;
                    }
                }
            }

            int fallbackIndex = lineIndex + 1;
            if (fallbackIndex <= tableEnd)
            {
                return NormalizeRequirementNameCell(lines[fallbackIndex]?.Text, true);
            }

            return string.Empty;
        }

        private static bool TryFindIdentifierAndNameHeaders(
            IReadOnlyList<TextLineSnapshot> lines,
            int tableStart,
            int tableEnd,
            int lineIndex,
            out int idHeaderIndex,
            out int nameHeaderIndex)
        {
            idHeaderIndex = -1;
            nameHeaderIndex = -1;
            if (lines == null)
            {
                return false;
            }

            int start = Math.Max(tableStart, lineIndex - 80);
            for (int i = start; i < lineIndex; i++)
            {
                if (!IsRequirementIdentifierHeader(lines[i]?.Text))
                {
                    continue;
                }

                int searchEnd = Math.Min(tableEnd, i + 6);
                for (int j = i + 1; j <= searchEnd; j++)
                {
                    if (ContainsAny(lines[j]?.Text, "需求名称", "能力名称", "接口名称", "名称"))
                    {
                        idHeaderIndex = i;
                        nameHeaderIndex = j;
                        return true;
                    }
                }
            }

            return false;
        }

        private static string TryGetNameAfterTableLabel(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            params string[] labels)
        {
            int tableStart;
            int tableEnd;
            if (!TryGetTableCellBounds(lines, lineIndex, out tableStart, out tableEnd))
            {
                return string.Empty;
            }

            int start = Math.Max(tableStart, lineIndex - 40);
            for (int i = lineIndex - 1; i >= start; i--)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (!ContainsAny(text, labels))
                {
                    continue;
                }

                int end = Math.Min(tableEnd, i + 3);
                for (int probe = i + 1; probe <= end; probe++)
                {
                    string candidate = NormalizeRequirementNameCell(lines[probe]?.Text, true);
                    if (!string.IsNullOrWhiteSpace(candidate))
                    {
                        return candidate;
                    }
                }
            }

            return string.Empty;
        }

        private static string TryGetFirstRowSecondCellName(IReadOnlyList<TextLineSnapshot> lines, int lineIndex)
        {
            int tableStart;
            int tableEnd;
            if (!TryGetTableCellBounds(lines, lineIndex, out tableStart, out tableEnd))
            {
                return string.Empty;
            }

            if (tableStart + 1 <= tableEnd &&
                ContainsAny(lines[tableStart]?.Text ?? string.Empty, "能力名称", "需求名称", "接口名称"))
            {
                return NormalizeRequirementNameCell(lines[tableStart + 1]?.Text, true);
            }

            return string.Empty;
        }

        private static string TryGetNameByHeaderOffset(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            params string[] nameHeaders)
        {
            int tableStart;
            int tableEnd;
            if (!TryGetTableCellBounds(lines, lineIndex, out tableStart, out tableEnd))
            {
                return string.Empty;
            }

            int searchStart = Math.Max(tableStart, lineIndex - 120);
            int bestNameHeaderIndex = -1;
            int bestIdHeaderIndex = -1;
            int bestDistance = int.MaxValue;

            for (int i = searchStart; i < lineIndex; i++)
            {
                string nameHeaderText = lines[i]?.Text ?? string.Empty;
                if (!ContainsAny(nameHeaderText, nameHeaders))
                {
                    continue;
                }

                int idSearchEnd = Math.Min(lineIndex - 1, i + 20);
                for (int idHeader = i + 1; idHeader <= idSearchEnd; idHeader++)
                {
                    if (!IsRequirementIdentifierHeader(lines[idHeader]?.Text))
                    {
                        continue;
                    }

                    int distance = lineIndex - idHeader;
                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        bestNameHeaderIndex = i;
                        bestIdHeaderIndex = idHeader;
                    }
                }
            }

            if (bestNameHeaderIndex < 0 || bestIdHeaderIndex <= bestNameHeaderIndex)
            {
                return string.Empty;
            }

            int nameOffsetFromId = bestIdHeaderIndex - bestNameHeaderIndex;
            int candidateIndex = lineIndex - nameOffsetFromId;
            if (candidateIndex < tableStart || candidateIndex > tableEnd || candidateIndex == lineIndex)
            {
                return string.Empty;
            }

            return NormalizeRequirementNameCell(lines[candidateIndex]?.Text, true);
        }

        private static string TryGetNearbyShortNameCell(IReadOnlyList<TextLineSnapshot> lines, int lineIndex)
        {
            int tableStart;
            int tableEnd;
            if (!TryGetTableCellBounds(lines, lineIndex, out tableStart, out tableEnd))
            {
                return string.Empty;
            }

            int start = Math.Max(tableStart, lineIndex - 10);
            for (int i = lineIndex - 1; i >= start; i--)
            {
                string candidate = NormalizeShortCapabilityName(lines[i]?.Text);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            return string.Empty;
        }

        private static bool TryGetTableCellBounds(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            out int tableStart,
            out int tableEnd)
        {
            tableStart = -1;
            tableEnd = -1;
            if (lines == null ||
                lineIndex < 0 ||
                lineIndex >= lines.Count ||
                !lines[lineIndex].IsTableCell)
            {
                return false;
            }

            tableStart = lineIndex;
            while (tableStart > 0 && lines[tableStart - 1].IsTableCell)
            {
                tableStart--;
            }

            tableEnd = lineIndex;
            while (tableEnd + 1 < lines.Count && lines[tableEnd + 1].IsTableCell)
            {
                tableEnd++;
            }

            return tableEnd >= tableStart;
        }

        private static bool IsRequirementIdentifierHeader(string text)
        {
            return ContainsAny(text, "唯一标识", "需求标识", "标识", "编号");
        }

        private static string NormalizeRequirementNameCell(string text, bool allowLong)
        {
            if (string.IsNullOrWhiteSpace(text) ||
                LooksLikeTableHeaderText(text) ||
                ContainsRequirementId(text, "SRS") ||
                IsHeadingLikeText(text))
            {
                return string.Empty;
            }

            string candidate = TrimRequirementNameFragment(RemoveRequirementIds(text));
            if (string.IsNullOrWhiteSpace(candidate) ||
                Regex.IsMatch(candidate, @"^\d+(?:\.\d+)*$"))
            {
                return string.Empty;
            }

            int maxLength = allowLong ? 80 : 30;
            return candidate.Length > maxLength ? string.Empty : candidate;
        }

        private static bool IsSectionNumberUnder(string sectionNumber, string prefix)
        {
            if (string.IsNullOrWhiteSpace(sectionNumber) || string.IsNullOrWhiteSpace(prefix))
            {
                return false;
            }

            return string.Equals(sectionNumber, prefix, StringComparison.OrdinalIgnoreCase) ||
                   sectionNumber.StartsWith(prefix + ".", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSectionDepthAtMost(string sectionNumber, int maxDepth)
        {
            if (string.IsNullOrWhiteSpace(sectionNumber) || maxDepth <= 0)
            {
                return false;
            }

            return ParseSectionNumberParts(sectionNumber).Length <= maxDepth;
        }

        private static List<RequirementItem> CollectSrsRequirementsFastPreview(
            Word.Document doc,
            Action<string> progressReporter)
        {
            List<RequirementItem> requirements = new List<RequirementItem>();
            if (doc == null)
            {
                return requirements;
            }

            Word.Range contentRange = null;
            try
            {
                ReportProgress(progressReporter, "正在快速读取需求文档...");
                contentRange = doc.Content;
                string rawText = contentRange?.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(rawText))
                {
                    return requirements;
                }

                string textForParsing = NormalizeRawTextForLineParsing(rawText);
                List<TextLineSnapshot> lines = BuildTextLineSnapshots(textForParsing);
                if (lines.Count == 0)
                {
                    return requirements;
                }

                bool thirdChapterStarted = false;
                HashSet<string> processedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                string currentSectionNumber = string.Empty;
                string currentSectionTitle = string.Empty;

                for (int i = 0; i < lines.Count; i++)
                {
                    TextLineSnapshot line = lines[i];
                    if (line == null || string.IsNullOrWhiteSpace(line.Text))
                    {
                        continue;
                    }

                    if (!thirdChapterStarted)
                    {
                        if (LooksLikeSrsThirdChapterStart(line.Text))
                        {
                            thirdChapterStarted = true;
                        }
                        else
                        {
                            continue;
                        }
                    }
                    else if (LooksLikeNextTopLevelChapter(line.Text))
                    {
                        break;
                    }

                    string sectionNumber;
                    string sectionTitle;
                    if (TryExtractSectionHeading(line.Text, out sectionNumber, out sectionTitle))
                    {
                        currentSectionNumber = sectionNumber;
                        currentSectionTitle = sectionTitle;
                    }

                    MatchCollection matches = RequirementIdRegex.Matches(line.Text);
                    for (int matchIndex = 0; matchIndex < matches.Count; matchIndex++)
                    {
                        Match match = matches[matchIndex];
                        if (!match.Success)
                        {
                            continue;
                        }

                        string requirementId = match.Groups["id"].Value.Trim();
                        if (!IsRequirementIdMatched(requirementId, "SRS") || !processedIds.Add(requirementId))
                        {
                            continue;
                        }

                        string requirementName = ResolveFastPreviewRequirementName(
                            lines,
                            i,
                            requirementId,
                            currentSectionNumber,
                            currentSectionTitle);
                        requirements.Add(CreateRequirementItem(
                            requirementId,
                            string.IsNullOrWhiteSpace(requirementName) ? requirementId : requirementName,
                            currentSectionNumber,
                            line.StartIndex + match.Index));
                    }

                    if (i == 0 || (i + 1) % 40 == 0)
                    {
                        ReportProgress(progressReporter, $"正在快速提取需求：{i + 1}/{lines.Count}");
                    }
                }

                if (requirements.Count == 0)
                {
                    return CollectSrsRequirementsFromWholeText(textForParsing, progressReporter);
                }

                return requirements;
            }
            catch
            {
                return requirements;
            }
            finally
            {
                ReleaseComObject(contentRange);
            }
        }

        private static void ApplyStructuredSrsTableOverrides(
            Word.Document doc,
            List<RequirementItem> requirements,
            Action<string> progressReporter)
        {
            if (doc == null || requirements == null)
            {
                return;
            }

            Dictionary<string, RequirementItem> existingLookup = requirements
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
            int overrideStart = FindFirstStructuredOverrideStart(requirements);
            if (overrideStart <= 0)
            {
                return;
            }

            Word.Tables tables = null;
            try
            {
                tables = doc.Tables;
                if (tables == null)
                {
                    return;
                }

                int tableCount = tables.Count;
                for (int i = 1; i <= tableCount; i++)
                {
                    Word.Table table = null;
                    Word.Range tableRange = null;
                    try
                    {
                        table = tables[i];
                        tableRange = table?.Range;
                        if (tableRange == null)
                        {
                            continue;
                        }

                        if (tableRange.End < overrideStart)
                        {
                            continue;
                        }

                        if (!LooksLikeStructuredRequirementTable(table))
                        {
                            continue;
                        }

                        List<RequirementItem> extracted = new List<RequirementItem>();
                        ExtractRequirementsFromTable(table, string.Empty, "SRS", extracted);
                        for (int itemIndex = 0; itemIndex < extracted.Count; itemIndex++)
                        {
                            RequirementItem extractedItem = extracted[itemIndex];
                            if (extractedItem == null ||
                                string.IsNullOrWhiteSpace(extractedItem.Id) ||
                                string.IsNullOrWhiteSpace(extractedItem.Name) ||
                                string.Equals(extractedItem.Name, "未命名需求", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            RequirementItem existingItem;
                            if (existingLookup.TryGetValue(extractedItem.Id, out existingItem))
                            {
                                existingItem.Name = extractedItem.Name;
                            }
                            else
                            {
                                requirements.Add(extractedItem);
                                existingLookup[extractedItem.Id] = extractedItem;
                            }
                        }
                    }
                    finally
                    {
                        ReleaseComObject(tableRange);
                        ReleaseComObject(table);
                        if (i == 1 || i % 10 == 0 || i == tableCount)
                        {
                            ReportProgress(progressReporter, $"正在校正表格需求：{i}/{tableCount}");
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(tables);
            }
        }

        private static int FindFirstStructuredOverrideStart(IEnumerable<RequirementItem> requirements)
        {
            return (requirements ?? Enumerable.Empty<RequirementItem>())
                .Where(item => item != null && IsSectionNumberAtOrAfter(item.SectionNumber, "3.3.7"))
                .Where(item => item.Start > 0)
                .Select(item => item.Start)
                .DefaultIfEmpty(0)
                .Min();
        }

        private static string FindNearestSectionNumberBeforePosition(Word.Document doc, int position)
        {
            if (doc == null || position <= 0)
            {
                return string.Empty;
            }

            string sectionNumber;
            string sectionTitle;
            if (TryFindNearestHeadingBeforePosition(doc, position, out sectionNumber, out sectionTitle))
            {
                return sectionNumber;
            }

            return string.Empty;
        }

        private static bool LooksLikeStructuredRequirementTable(Word.Table table)
        {
            if (table == null)
            {
                return false;
            }

            TableHeaderInfo headerInfo = AnalyzeTableHeader(table);
            return headerInfo.IdColumnIndex >= 0 &&
                   headerInfo.NameColumnIndex >= 0;
        }

        private static List<RequirementItem> CollectSrsRequirementsFromWholeText(
            string rawText,
            Action<string> progressReporter)
        {
            List<RequirementItem> requirements = new List<RequirementItem>();
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return requirements;
            }

            List<TextLineSnapshot> lines = BuildTextLineSnapshots(rawText);
            HashSet<string> processedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string currentSectionNumber = string.Empty;
            string currentSectionTitle = string.Empty;

            for (int i = 0; i < lines.Count; i++)
            {
                TextLineSnapshot line = lines[i];
                if (line == null || string.IsNullOrWhiteSpace(line.Text))
                {
                    continue;
                }

                string sectionNumber;
                string sectionTitle;
                if (TryExtractSectionHeading(line.Text, out sectionNumber, out sectionTitle))
                {
                    currentSectionNumber = sectionNumber;
                    currentSectionTitle = sectionTitle;
                }

                MatchCollection matches = RequirementIdRegex.Matches(line.Text);
                for (int matchIndex = 0; matchIndex < matches.Count; matchIndex++)
                {
                    Match match = matches[matchIndex];
                    if (!match.Success)
                    {
                        continue;
                    }

                    string requirementId = match.Groups["id"].Value.Trim();
                    if (!IsRequirementIdMatched(requirementId, "SRS") || !processedIds.Add(requirementId))
                    {
                        continue;
                    }

                    string requirementName = ResolveFastPreviewRequirementName(
                        lines,
                        i,
                        requirementId,
                        currentSectionNumber,
                        currentSectionTitle);
                    requirements.Add(CreateRequirementItem(
                        requirementId,
                        string.IsNullOrWhiteSpace(requirementName) ? requirementId : requirementName,
                        currentSectionNumber,
                        line.StartIndex + match.Index));
                }
            }

            ReportProgress(progressReporter, $"全文快速提取完成，共识别 {requirements.Count} 项");
            return requirements;
        }

        private static string ResolveFastPreviewRequirementName(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            string requirementId,
            string currentSectionNumber,
            string currentSectionTitle)
        {
            string candidate = TryResolveFlatTableRequirementName(lines, lineIndex, currentSectionNumber);
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            candidate = FindNearbyRequirementNameByLabel(lines, lineIndex, currentSectionNumber);
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            int[] offsets = new[] { 0, -1, 1, -2, 2 };
            for (int i = 0; i < offsets.Length; i++)
            {
                int probeIndex = lineIndex + offsets[i];
                if (probeIndex < 0 || probeIndex >= lines.Count)
                {
                    continue;
                }

                candidate = ExtractNameFromText(lines[probeIndex].Text, requirementId);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            if (!string.IsNullOrWhiteSpace(currentSectionTitle) &&
                !LooksLikeNoiseText(currentSectionTitle))
            {
                return currentSectionTitle;
            }

            return requirementId;
        }

        private static string TryResolveFlatTableRequirementName(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            string sectionNumber)
        {
            if (!IsSectionNumberAtOrAfter(sectionNumber, "3.3.7") ||
                lines == null ||
                lineIndex < 0 ||
                lineIndex >= lines.Count)
            {
                return string.Empty;
            }

            int headerIndex = FindCapabilityNameHeaderLineIndex(lines, lineIndex);
            if (headerIndex < 0)
            {
                return string.Empty;
            }

            int forwardEnd = Math.Min(lines.Count - 1, lineIndex + 2);
            for (int i = lineIndex + 1; i <= forwardEnd; i++)
            {
                string text = lines[i]?.Text ?? string.Empty;
                string candidate = NormalizeShortCapabilityName(text);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            int headerEnd = Math.Min(lines.Count - 1, headerIndex + 6);
            for (int i = headerIndex + 1; i <= headerEnd; i++)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (ContainsRequirementId(text, "SRS"))
                {
                    if (i > lineIndex)
                    {
                        break;
                    }

                    continue;
                }

                string candidate = NormalizeShortCapabilityName(text);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            return string.Empty;
        }

        private static int FindCapabilityNameHeaderLineIndex(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex)
        {
            if (lines == null || lineIndex < 0 || lineIndex >= lines.Count)
            {
                return -1;
            }

            int start = Math.Max(0, lineIndex - 6);
            int end = Math.Min(lines.Count - 1, lineIndex + 2);
            for (int i = start; i <= end; i++)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (ContainsAny(text, "能力名称"))
                {
                    return i;
                }
            }

            return -1;
        }

        private static string NormalizeShortCapabilityName(string text)
        {
            if (string.IsNullOrWhiteSpace(text) ||
                LooksLikeTableHeaderText(text) ||
                ContainsRequirementId(text, "SRS") ||
                IsHeadingLikeText(text))
            {
                return string.Empty;
            }

            string candidate = TrimRequirementNameFragment(text);
            if (string.IsNullOrWhiteSpace(candidate) || candidate.Length > 15)
            {
                return string.Empty;
            }

            return candidate;
        }

        private static string FindNearbyRequirementNameByLabel(
            IReadOnlyList<TextLineSnapshot> lines,
            int lineIndex,
            string sectionNumber)
        {
            if (lines == null || lineIndex < 0 || lineIndex >= lines.Count)
            {
                return string.Empty;
            }

            int start = Math.Max(0, lineIndex - 4);
            int end = Math.Min(lines.Count - 1, lineIndex + 4);
            for (int i = start; i <= end; i++)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (!IsRequirementNameLabel(text, sectionNumber))
                {
                    continue;
                }

                string nextValue = GetNearestMeaningfulLine(lines, i + 1, end, 1);
                if (!string.IsNullOrWhiteSpace(nextValue))
                {
                    return nextValue;
                }

                string previousValue = GetNearestMeaningfulLine(lines, i - 1, start, -1);
                if (!string.IsNullOrWhiteSpace(previousValue))
                {
                    return previousValue;
                }
            }

            return string.Empty;
        }

        private static string GetNearestMeaningfulLine(
            IReadOnlyList<TextLineSnapshot> lines,
            int startIndex,
            int boundaryIndex,
            int step)
        {
            for (int i = startIndex; step > 0 ? i <= boundaryIndex : i >= boundaryIndex; i += step)
            {
                string text = lines[i]?.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text) ||
                    LooksLikeTableHeaderText(text) ||
                    ContainsRequirementId(text, "SRS"))
                {
                    continue;
                }

                return TrimRequirementNameFragment(text);
            }

            return string.Empty;
        }

        private static bool LooksLikeSrsThirdChapterStart(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return ContainsAny(text, "第3章", "第三章") ||
                   (text.StartsWith("3 ", StringComparison.OrdinalIgnoreCase) && ContainsAny(text, "需求", "要求", "规格", "说明")) ||
                   (text.StartsWith("3.", StringComparison.OrdinalIgnoreCase) && ContainsAny(text, "需求", "要求", "状态", "能力", "接口"));
        }

        private static bool LooksLikeNextTopLevelChapter(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return ContainsAny(text, "第4章", "第四章") ||
                   text.StartsWith("4 ", StringComparison.OrdinalIgnoreCase) ||
                   text.StartsWith("4.", StringComparison.OrdinalIgnoreCase);
        }

        private static List<TextLineSnapshot> BuildTextLineSnapshots(string rawText)
        {
            List<TextLineSnapshot> lines = new List<TextLineSnapshot>();
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return lines;
            }

            int lineStart = 0;
            int index = 0;
            while (index <= rawText.Length)
            {
                bool atEnd = index == rawText.Length;
                if (!atEnd && rawText[index] != '\n')
                {
                    index++;
                    continue;
                }

                string lineText = NormalizeParagraphText(SafeSubstring(rawText, lineStart, index - lineStart));
                if (!string.IsNullOrWhiteSpace(lineText))
                {
                    lines.Add(new TextLineSnapshot
                    {
                        StartIndex = lineStart,
                        Text = lineText
                    });
                }

                lineStart = index + 1;
                index++;
            }

            return lines;
        }

        private static string NormalizeRawTextForLineParsing(string rawText)
        {
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return string.Empty;
            }

            string normalized = rawText.Replace("\r\n", "\n").Replace('\r', '\n');
            normalized = normalized.Replace('\a', '\n').Replace('\u0007', '\n');
            return normalized;
        }

        private static List<RequirementItem> CollectSrsRequirementsFromTextSnapshot(
            Word.Document doc,
            int chapterStart,
            int chapterEnd,
            IReadOnlyList<SectionHeadingSnapshot> headings,
            Action<string> progressReporter)
        {
            List<RequirementItem> requirements = new List<RequirementItem>();
            if (doc == null)
            {
                return requirements;
            }

            Word.Range chapterRange = null;
            try
            {
                ReportProgress(progressReporter, "正在快速读取文档文本...");
                chapterRange = doc.Range(chapterStart, chapterEnd);
                string rawText = chapterRange?.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(rawText))
                {
                    return requirements;
                }

                MatchCollection matches = RequirementIdRegex.Matches(rawText);
                HashSet<string> processedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < matches.Count; i++)
                {
                    Match match = matches[i];
                    if (!match.Success)
                    {
                        continue;
                    }

                    string requirementId = match.Groups["id"].Value.Trim();
                    if (!IsRequirementIdMatched(requirementId, "SRS"))
                    {
                        continue;
                    }

                    if (!processedIds.Add(requirementId))
                    {
                        continue;
                    }

                    RequirementItem requirement = ResolveSrsRequirementByIdentifier(
                        doc,
                        chapterStart,
                        chapterEnd,
                        headings,
                        requirementId,
                        rawText,
                        match.Index);
                    if (requirement != null)
                    {
                        requirements.Add(requirement);
                    }

                    if (i == 0 || (i + 1) % 30 == 0)
                    {
                        ReportProgress(progressReporter, $"正在快速提取需求：{i + 1}/{matches.Count}");
                    }
                }

                return requirements;
            }
            catch
            {
                return requirements;
            }
            finally
            {
                ReleaseComObject(chapterRange);
            }
        }

        private static RequirementItem ResolveSrsRequirementByIdentifier(
            Word.Document doc,
            int chapterStart,
            int chapterEnd,
            IReadOnlyList<SectionHeadingSnapshot> headings,
            string requirementId,
            string chapterText,
            int fallbackIndex)
        {
            if (doc == null || string.IsNullOrWhiteSpace(requirementId))
            {
                return null;
            }

            Word.Range searchRange = null;
            try
            {
                searchRange = doc.Range(chapterStart, chapterEnd);
                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                RequirementItem bestCandidate = null;
                int bestQuality = int.MinValue;
                while (FindNextSpecificRequirementId(searchRange, chapterEnd, requirementId, out Word.Range hitRange))
                {
                    try
                    {
                        int quality;
                        RequirementItem candidate = BuildRequirementFromSpecificHitRange(
                            hitRange,
                            requirementId,
                            headings,
                            out quality);
                        if (candidate != null && quality > bestQuality)
                        {
                            bestCandidate = candidate;
                            bestQuality = quality;
                            if (quality >= 120)
                            {
                                return bestCandidate;
                            }
                        }

                        AdvanceSearchRange(searchRange, hitRange, chapterEnd);
                    }
                    finally
                    {
                        ReleaseComObject(hitRange);
                    }
                }

                if (bestCandidate != null)
                {
                    return bestCandidate;
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(searchRange);
            }

            int fallbackStart = chapterStart + Math.Max(0, fallbackIndex);
            string fallbackSectionNumber = FindNearestHeadingSectionNumber(headings, fallbackStart);
            string fallbackSectionTitle = FindNearestHeadingTitle(headings, fallbackStart);
            string fallbackName = ResolveQuickRequirementName(chapterText, fallbackIndex, requirementId);
            if (string.IsNullOrWhiteSpace(fallbackName))
            {
                fallbackName = fallbackSectionTitle;
            }

            return CreateRequirementItem(requirementId, fallbackName, fallbackSectionNumber, fallbackStart);
        }

        private static bool FindNextSpecificRequirementId(
            Word.Range searchRange,
            int chapterEnd,
            string requirementId,
            out Word.Range hitRange)
        {
            hitRange = null;
            if (searchRange == null || string.IsNullOrWhiteSpace(requirementId))
            {
                return false;
            }

            Word.Find find = null;
            try
            {
                find = searchRange.Find;
                if (find == null)
                {
                    return false;
                }

                find.ClearFormatting();
                find.Text = requirementId;
                find.MatchWildcards = false;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                if (!find.Execute())
                {
                    return false;
                }

                hitRange = searchRange.Duplicate;
                return hitRange != null && hitRange.Start < chapterEnd;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComObject(find);
            }
        }

        private static RequirementItem BuildRequirementFromSpecificHitRange(
            Word.Range hitRange,
            string requirementId,
            IReadOnlyList<SectionHeadingSnapshot> headings,
            out int quality)
        {
            quality = 0;
            if (hitRange == null || string.IsNullOrWhiteSpace(requirementId))
            {
                return null;
            }

            string sectionNumber = FindNearestHeadingSectionNumber(headings, hitRange.Start);
            string sectionTitle = FindNearestHeadingTitle(headings, hitRange.Start);
            bool isTableHit = IsRangeInsideTable(hitRange);
            string explicitName = isTableHit
                ? TryResolveNameFromTableHit(hitRange, requirementId, sectionNumber)
                : TryResolveNameFromParagraphHit(hitRange, requirementId);

            string name = explicitName;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = sectionTitle;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                name = requirementId;
            }

            if (!string.IsNullOrWhiteSpace(explicitName))
            {
                quality = isTableHit ? 120 : 100;
            }
            else if (!string.IsNullOrWhiteSpace(sectionTitle))
            {
                quality = 40;
            }
            else
            {
                quality = 10;
            }

            return CreateRequirementItem(requirementId, name, sectionNumber, hitRange.Start);
        }

        private static string ResolveQuickRequirementName(string rawText, int matchIndex, string requirementId)
        {
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return string.Empty;
            }

            string currentLine = GetContextLine(rawText, matchIndex, 0);
            string previousLine = GetContextLine(rawText, matchIndex, -1);
            string nextLine = GetContextLine(rawText, matchIndex, 1);

            string name = ExtractNameFromText(currentLine, requirementId);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            name = ExtractNameFromText(previousLine, requirementId);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            name = ExtractNameFromText(nextLine, requirementId);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            return string.Empty;
        }

        private static string GetContextLine(string rawText, int matchIndex, int relativeLineOffset)
        {
            if (string.IsNullOrWhiteSpace(rawText))
            {
                return string.Empty;
            }

            int currentLineStart = FindLineStart(rawText, matchIndex);
            int currentLineEnd = FindLineEnd(rawText, matchIndex);
            if (relativeLineOffset == 0)
            {
                return NormalizeParagraphText(SafeSubstring(rawText, currentLineStart, currentLineEnd - currentLineStart));
            }

            if (relativeLineOffset < 0)
            {
                int probeEnd = Math.Max(0, currentLineStart - 1);
                for (int step = 0; step < -relativeLineOffset; step++)
                {
                    probeEnd = FindPreviousLineBreak(rawText, probeEnd);
                    if (probeEnd < 0)
                    {
                        return string.Empty;
                    }
                }

                int previousStart = FindLineStart(rawText, probeEnd);
                int previousEnd = FindLineEnd(rawText, previousStart);
                return NormalizeParagraphText(SafeSubstring(rawText, previousStart, previousEnd - previousStart));
            }

            int probeStart = currentLineEnd;
            for (int step = 0; step < relativeLineOffset; step++)
            {
                probeStart = FindNextLineStart(rawText, probeStart);
                if (probeStart < 0)
                {
                    return string.Empty;
                }
            }

            int nextEnd = FindLineEnd(rawText, probeStart);
            return NormalizeParagraphText(SafeSubstring(rawText, probeStart, nextEnd - probeStart));
        }

        private static int FindLineStart(string text, int index)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            int current = Math.Max(0, Math.Min(index, text.Length - 1));
            while (current > 0 && !IsLineBreakChar(text[current - 1]))
            {
                current--;
            }

            return current;
        }

        private static int FindLineEnd(string text, int index)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            int current = Math.Max(0, Math.Min(index, text.Length));
            while (current < text.Length && !IsLineBreakChar(text[current]))
            {
                current++;
            }

            return current;
        }

        private static int FindPreviousLineBreak(string text, int index)
        {
            if (string.IsNullOrEmpty(text))
            {
                return -1;
            }

            int current = Math.Max(0, Math.Min(index, text.Length - 1));
            while (current >= 0)
            {
                if (IsLineBreakChar(text[current]))
                {
                    while (current >= 0 && IsLineBreakChar(text[current]))
                    {
                        current--;
                    }

                    return current;
                }

                current--;
            }

            return -1;
        }

        private static int FindNextLineStart(string text, int index)
        {
            if (string.IsNullOrEmpty(text))
            {
                return -1;
            }

            int current = Math.Max(0, Math.Min(index, text.Length));
            while (current < text.Length && IsLineBreakChar(text[current]))
            {
                current++;
            }

            return current < text.Length ? current : -1;
        }

        private static bool IsLineBreakChar(char value)
        {
            return value == '\r' || value == '\n' || value == '\a' || value == '\u0007';
        }

        private static string SafeSubstring(string text, int startIndex, int length)
        {
            if (string.IsNullOrEmpty(text) || startIndex >= text.Length || length <= 0)
            {
                return string.Empty;
            }

            int safeStart = Math.Max(0, startIndex);
            int safeLength = Math.Min(length, text.Length - safeStart);
            return safeLength > 0 ? text.Substring(safeStart, safeLength) : string.Empty;
        }

        private static List<RequirementItem> CollectSrsRequirementsByIdentifierSearch(
            Word.Document doc,
            int chapterStart,
            int chapterEnd,
            IReadOnlyList<SectionHeadingSnapshot> headings,
            Action<string> progressReporter)
        {
            List<RequirementItem> requirements = new List<RequirementItem>();
            if (doc == null || chapterEnd <= chapterStart)
            {
                return requirements;
            }

            Word.Range chapterRange = null;
            Word.Range searchRange = null;
            try
            {
                chapterRange = doc.Range(chapterStart, chapterEnd);
                if (chapterRange == null)
                {
                    return requirements;
                }

                searchRange = chapterRange.Duplicate;
                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                int foundCount = 0;
                while (FindNextRequirementId(searchRange, chapterEnd, out Word.Range hitRange, out string requirementId))
                {
                    if (hitRange == null)
                    {
                        break;
                    }

                    if (hitRange.Start >= chapterEnd)
                    {
                        ReleaseComObject(hitRange);
                        break;
                    }

                    string sectionNumber = FindNearestHeadingSectionNumber(headings, hitRange.Start);
                    string sectionTitle = FindNearestHeadingTitle(headings, hitRange.Start);
                    RequirementItem requirement = BuildRequirementFromHitRange(hitRange, requirementId, sectionNumber, sectionTitle);
                    if (requirement != null)
                    {
                        requirements.Add(requirement);
                    }

                    foundCount++;
                    if (foundCount == 1 || foundCount % 20 == 0)
                    {
                        ReportProgress(progressReporter, $"正在按标识解析第三章：已识别 {foundCount} 项");
                    }

                    AdvanceSearchRange(searchRange, hitRange, chapterEnd);
                    ReleaseComObject(hitRange);
                }

                return requirements;
            }
            finally
            {
                ReleaseComObject(searchRange);
                ReleaseComObject(chapterRange);
            }
        }

        private static List<SectionHeadingSnapshot> CollectSrsSectionHeadings(
            Word.Document doc,
            int chapterStart,
            int chapterEnd,
            Action<string> progressReporter)
        {
            List<SectionHeadingSnapshot> headings = new List<SectionHeadingSnapshot>();
            Word.Range chapterRange = null;
            Word.Paragraphs paragraphs = null;
            try
            {
                chapterRange = doc.Range(chapterStart, chapterEnd);
                paragraphs = chapterRange?.Paragraphs;
                if (paragraphs == null)
                {
                    return headings;
                }

                int paragraphCount = paragraphs.Count;
                int progressInterval = paragraphCount > 1000 ? 100 : paragraphCount > 400 ? 50 : 20;
                for (int i = 1; i <= paragraphCount; i++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                        if (paragraphRange == null || paragraphRange.Start < chapterStart || paragraphRange.Start >= chapterEnd)
                        {
                            continue;
                        }

                        if (IsRangeInsideTable(paragraphRange))
                        {
                            continue;
                        }

                        string paragraphText = NormalizeParagraphText(paragraphRange.Text);
                        string sectionNumber;
                        string sectionTitle;
                        if (TryExtractHeadingFromParagraph(paragraph, paragraphText, out sectionNumber, out sectionTitle))
                        {
                            headings.Add(new SectionHeadingSnapshot
                            {
                                Start = paragraphRange.Start,
                                SectionNumber = sectionNumber,
                                SectionTitle = sectionTitle
                            });
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                        if (i == 1 || i % progressInterval == 0 || i == paragraphCount)
                        {
                            ReportProgress(progressReporter, $"正在提取第三章标题：{i}/{paragraphCount}");
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(paragraphs);
                ReleaseComObject(chapterRange);
            }

            return headings;
        }

        private static bool FindNextRequirementId(
            Word.Range searchRange,
            int chapterEnd,
            out Word.Range hitRange,
            out string requirementId)
        {
            hitRange = null;
            requirementId = string.Empty;
            if (searchRange == null)
            {
                return false;
            }

            Word.Find find = null;
            try
            {
                find = searchRange.Find;
                if (find == null)
                {
                    return false;
                }

                find.ClearFormatting();
                find.Text = "SRS-[0-9][0-9][0-9][0-9]";
                find.MatchWildcards = true;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                if (!find.Execute())
                {
                    return false;
                }

                hitRange = searchRange.Duplicate;
                if (hitRange == null)
                {
                    return false;
                }

                string rawText = NormalizeParagraphText(hitRange.Text);
                Match match = RequirementIdRegex.Match(rawText);
                requirementId = match.Success ? match.Groups["id"].Value.Trim() : rawText.Trim();
                return !string.IsNullOrWhiteSpace(requirementId) && hitRange.Start < chapterEnd;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseComObject(find);
            }
        }

        private static void AdvanceSearchRange(Word.Range searchRange, Word.Range hitRange, int chapterEnd)
        {
            if (searchRange == null || hitRange == null)
            {
                return;
            }

            try
            {
                int nextStart = Math.Min(hitRange.End, chapterEnd);
                searchRange.SetRange(nextStart, chapterEnd);
                searchRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            }
            catch
            {
            }
        }

        private static RequirementItem BuildRequirementFromHitRange(
            Word.Range hitRange,
            string requirementId,
            string sectionNumber,
            string sectionTitle)
        {
            if (hitRange == null || string.IsNullOrWhiteSpace(requirementId))
            {
                return null;
            }

            string name = string.Empty;
            if (IsRangeInsideTable(hitRange))
            {
                name = TryResolveNameFromTableHit(hitRange, requirementId, sectionNumber);
            }
            else
            {
                name = TryResolveNameFromParagraphHit(hitRange, requirementId);
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                name = sectionTitle;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                name = "未命名需求";
            }

            return CreateRequirementItem(requirementId, name, sectionNumber, hitRange.Start);
        }

        private static string TryResolveNameFromParagraphHit(Word.Range hitRange, string requirementId)
        {
            Word.Paragraph paragraph = null;
            Word.Range paragraphRange = null;
            try
            {
                paragraph = hitRange?.Paragraphs[1];
                paragraphRange = paragraph?.Range;
                string paragraphText = NormalizeParagraphText(paragraphRange?.Text);
                string name = ExtractNameFromText(paragraphText, requirementId);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    return name;
                }

                Word.Range localRange = paragraphRange?.Duplicate;
                try
                {
                    if (localRange != null)
                    {
                        localRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        localRange.MoveStart(Word.WdUnits.wdParagraph, -1);
                        localRange.MoveEnd(Word.WdUnits.wdParagraph, 1);
                        name = ExtractNameFromText(NormalizeParagraphText(localRange.Text), requirementId);
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            return name;
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(localRange);
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(paragraphRange);
                ReleaseComObject(paragraph);
            }

            return string.Empty;
        }

        private static string TryResolveNameFromTableHit(Word.Range hitRange, string requirementId, string sectionNumber)
        {
            Word.Cell cell = null;
            Word.Row row = null;
            Word.Table table = null;
            try
            {
                cell = hitRange?.Cells[1];
                row = cell?.Row;
                table = cell?.Range?.Tables[1];
                if (row == null || table == null)
                {
                    return string.Empty;
                }

                string keyValueName = TryResolveNameFromKeyValueTable(table, requirementId, sectionNumber);
                if (!string.IsNullOrWhiteSpace(keyValueName))
                {
                    return keyValueName;
                }

                List<string> cellTexts = GetRowCellTexts(row);
                if (cellTexts.Count == 0)
                {
                    return string.Empty;
                }

                TableHeaderInfo headerInfo = AnalyzeTableHeader(table);
                int idCellIndex = FindCellIndexContainingId(cellTexts, requirementId);
                string name = PickRequirementNameFromRow(cellTexts, headerInfo, idCellIndex, requirementId);
                return name;
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ReleaseComObject(table);
                ReleaseComObject(row);
                ReleaseComObject(cell);
            }
        }

        private static string TryResolveNameFromKeyValueTable(
            Word.Table table,
            string requirementId,
            string sectionNumber)
        {
            if (table == null || string.IsNullOrWhiteSpace(requirementId))
            {
                return string.Empty;
            }

            string resolvedName = string.Empty;
            bool idMatched = false;
            
            List<TableRowData> tableData = ParseTableData(table);
            for (int rowIndex = 0; rowIndex < tableData.Count; rowIndex++)
            {
                List<string> cellTexts = tableData[rowIndex].CellTexts;
                if (cellTexts.Count < 2)
                {
                    continue;
                }

                for (int cellIndex = 0; cellIndex < cellTexts.Count - 1; cellIndex++)
                {
                    string label = NormalizeParagraphText(cellTexts[cellIndex]);
                    string value = NormalizeParagraphText(cellTexts[cellIndex + 1]);
                    if (string.IsNullOrWhiteSpace(label) || string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    if (IsRequirementIdLabel(label) &&
                        value.IndexOf(requirementId, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        idMatched = true;
                    }

                    if (IsRequirementNameLabel(label, sectionNumber) &&
                        !LooksLikeTableHeaderText(value) &&
                        !ContainsRequirementId(value, "SRS"))
                    {
                        resolvedName = TrimRequirementNameFragment(value);
                    }
                }
            }

            return idMatched ? resolvedName : string.Empty;
        }

        private static bool IsRequirementIdLabel(string text)
        {
            return ContainsAny(text, "唯一标识", "需求标识", "标识", "编号");
        }

        private static bool IsRequirementNameLabel(string text, string sectionNumber)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            if (IsSrsInterfaceSection(sectionNumber))
            {
                return ContainsAny(text, "接口名称", "能力名称", "需求名称", "功能名称", "名称");
            }

            return ContainsAny(text, "能力名称", "需求名称", "功能名称", "名称");
        }

        private static bool IsSrsInterfaceSection(string sectionNumber)
        {
            return !string.IsNullOrWhiteSpace(sectionNumber) &&
                   sectionNumber.StartsWith("3.3", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSectionNumberAtOrAfter(string sectionNumber, string baseline)
        {
            int[] left = ParseSectionNumberParts(sectionNumber);
            int[] right = ParseSectionNumberParts(baseline);
            if (left.Length == 0 || right.Length == 0)
            {
                return false;
            }

            int count = Math.Max(left.Length, right.Length);
            for (int i = 0; i < count; i++)
            {
                int leftPart = i < left.Length ? left[i] : 0;
                int rightPart = i < right.Length ? right[i] : 0;
                if (leftPart == rightPart)
                {
                    continue;
                }

                return leftPart > rightPart;
            }

            return true;
        }

        private static int[] ParseSectionNumberParts(string sectionNumber)
        {
            if (string.IsNullOrWhiteSpace(sectionNumber))
            {
                return new int[0];
            }

            string[] tokens = sectionNumber.Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            List<int> parts = new List<int>();
            for (int i = 0; i < tokens.Length; i++)
            {
                int value;
                if (int.TryParse(tokens[i], out value))
                {
                    parts.Add(value);
                }
                else
                {
                    break;
                }
            }

            return parts.ToArray();
        }

        private static string TruncateSectionNumberToDepth(string sectionNumber, int maxDepth)
        {
            if (string.IsNullOrWhiteSpace(sectionNumber) || maxDepth <= 0)
            {
                return string.Empty;
            }

            int[] parts = ParseSectionNumberParts(sectionNumber);
            if (parts.Length == 0)
            {
                return string.Empty;
            }

            return string.Join(".", parts.Take(maxDepth));
        }

        private static string ExtractNameFromText(string text, string requirementId)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            string cleaned = RequirementIdRegex.Replace(text, " ");
            if (!string.IsNullOrWhiteSpace(requirementId))
            {
                cleaned = cleaned.Replace(requirementId, string.Empty);
            }

            cleaned = TrimRequirementNameFragment(cleaned);
            if (LooksLikeNoiseText(cleaned))
            {
                return string.Empty;
            }

            return cleaned;
        }

        private static bool LooksLikeNoiseText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            return string.Equals(text, "需求标识", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "唯一标识", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "标识", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "编号", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "能力名称", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "接口名称", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "用途", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "来源", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "接收者", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "描述", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "说明", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "备注", StringComparison.OrdinalIgnoreCase);
        }

        private static int FindCellIndexContainingId(IReadOnlyList<string> cellTexts, string requirementId)
        {
            if (cellTexts == null || string.IsNullOrWhiteSpace(requirementId))
            {
                return -1;
            }

            for (int i = 0; i < cellTexts.Count; i++)
            {
                string text = cellTexts[i] ?? string.Empty;
                if (text.IndexOf(requirementId, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return i;
                }
            }

            return -1;
        }

        private static string FindNearestHeadingSectionNumber(IReadOnlyList<SectionHeadingSnapshot> headings, int start)
        {
            SectionHeadingSnapshot heading = FindNearestHeading(headings, start);
            return heading?.SectionNumber ?? string.Empty;
        }

        private static string FindNearestHeadingTitle(IReadOnlyList<SectionHeadingSnapshot> headings, int start)
        {
            SectionHeadingSnapshot heading = FindNearestHeading(headings, start);
            return heading?.SectionTitle ?? string.Empty;
        }

        private static SectionHeadingSnapshot FindNearestHeading(IReadOnlyList<SectionHeadingSnapshot> headings, int start)
        {
            if (headings == null || headings.Count == 0)
            {
                return null;
            }

            SectionHeadingSnapshot candidate = null;
            for (int i = 0; i < headings.Count; i++)
            {
                SectionHeadingSnapshot heading = headings[i];
                if (heading == null || heading.Start > start)
                {
                    break;
                }

                candidate = heading;
            }

            return candidate;
        }

        private static void CollectSrsParagraphRequirements(
            Word.Document doc,
            int chapterStart,
            int chapterEnd,
            List<RequirementItem> requirements,
            Action<string> progressReporter)
        {
            Word.Range chapterRange = null;
            Word.Paragraphs paragraphs = null;
            try
            {
                chapterRange = doc.Range(chapterStart, chapterEnd);
                paragraphs = chapterRange.Paragraphs;
                if (paragraphs == null)
                {
                    return;
                }

                string currentSectionNumber = string.Empty;
                string currentSectionTitle = string.Empty;
                int paragraphCount = paragraphs.Count;
                int progressInterval = paragraphCount > 1200 ? 120 : paragraphCount > 400 ? 50 : 20;

                for (int i = 1; i <= paragraphCount; i++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                        if (paragraphRange == null || paragraphRange.Start < chapterStart || paragraphRange.Start >= chapterEnd)
                        {
                            continue;
                        }

                        if (IsRangeInsideTable(paragraphRange))
                        {
                            continue;
                        }

                        string paragraphText = NormalizeParagraphText(paragraphRange.Text);
                        if (string.IsNullOrWhiteSpace(paragraphText))
                        {
                            continue;
                        }

                        string updatedSectionNumber;
                        string updatedSectionTitle;
                        if (TryExtractSectionHeading(paragraphText, out updatedSectionNumber, out updatedSectionTitle))
                        {
                            currentSectionNumber = updatedSectionNumber;
                            if (!string.IsNullOrWhiteSpace(updatedSectionTitle))
                            {
                                currentSectionTitle = updatedSectionTitle;
                            }
                        }

                        MatchCollection matches = RequirementIdRegex.Matches(paragraphText);
                        if (matches.Count == 0)
                        {
                            continue;
                        }

                        for (int matchIndex = 0; matchIndex < matches.Count; matchIndex++)
                        {
                            Match match = matches[matchIndex];
                            if (!match.Success)
                            {
                                continue;
                            }

                            string id = match.Groups["id"].Value.Trim();
                            if (!IsRequirementIdMatched(id, "SRS"))
                            {
                                continue;
                            }

                            string name = ResolveParagraphRequirementName(
                                paragraphs,
                                i,
                                currentSectionTitle,
                                paragraphText,
                                match,
                                matches,
                                id);

                            requirements.Add(CreateRequirementItem(
                                id,
                                name,
                                currentSectionNumber,
                                paragraphRange.Start));
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                        if (i == 1 || i % progressInterval == 0 || i == paragraphCount)
                        {
                            ReportProgress(progressReporter, $"正在解析第三章段落：{i}/{paragraphCount}，已识别 {requirements.Count} 项");
                        }
                    }
                }
            }
            finally
            {
                ReleaseComObject(paragraphs);
                ReleaseComObject(chapterRange);
            }
        }

        private static void CollectSrsTableRequirements(
            Word.Document doc,
            int chapterStart,
            int chapterEnd,
            List<RequirementItem> requirements,
            IReadOnlyList<SectionHeadingSnapshot> headings,
            Action<string> progressReporter)
        {
            Word.Range chapterRange = null;
            Word.Tables tables = null;
            try
            {
                chapterRange = doc.Range(chapterStart, chapterEnd);
                tables = chapterRange.Tables;
                if (tables == null)
                {
                    return;
                }

                int tableCount = tables.Count;
                for (int i = 1; i <= tableCount; i++)
                {
                    Word.Table table = null;
                    Word.Range tableRange = null;
                    try
                    {
                        table = tables[i];
                        tableRange = table?.Range;
                        if (tableRange == null)
                        {
                            continue;
                        }

                        if (tableRange.Start < chapterStart || tableRange.Start >= chapterEnd)
                        {
                            continue;
                        }

                        string sectionNumber = FindNearestHeadingSectionNumber(headings, tableRange.Start);
                        ExtractRequirementsFromTable(table, sectionNumber, "SRS", requirements);
                    }
                    finally
                    {
                        ReleaseComObject(tableRange);
                        ReleaseComObject(table);
                        ReportProgress(progressReporter, $"正在解析第三章表格：{i}/{tableCount}，已识别 {requirements.Count} 项");
                    }
                }
            }
            finally
            {
                ReleaseComObject(tables);
                ReleaseComObject(chapterRange);
            }
        }

        private sealed class TableRowData
        {
            public int Start { get; set; }
            public List<string> CellTexts { get; set; } = new List<string>();
        }

        private static List<TableRowData> ParseTableData(Word.Table table)
        {
            List<TableRowData> grid = new List<TableRowData>();
            if (table == null) return grid;

            Word.Cells cells = null;
            try
            {
                cells = table.Range.Cells;
                if (cells == null) return grid;

                int count = cells.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Cell cell = null;
                    Word.Range cellRange = null;
                    try
                    {
                        cell = cells[i];
                        int r = cell.RowIndex - 1;
                        int c = cell.ColumnIndex - 1;
                        cellRange = cell.Range;
                        string text = NormalizeParagraphText(cellRange?.Text) ?? string.Empty;
                        int start = cellRange?.Start ?? 0;

                        while (grid.Count <= r)
                        {
                            grid.Add(new TableRowData());
                        }

                        TableRowData rowData = grid[r];
                        if (rowData.Start == 0 || (start > 0 && start < rowData.Start))
                        {
                            rowData.Start = start;
                        }

                        while (rowData.CellTexts.Count <= c)
                        {
                            rowData.CellTexts.Add(string.Empty);
                        }

                        if (string.IsNullOrWhiteSpace(rowData.CellTexts[c]))
                        {
                            rowData.CellTexts[c] = text;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(cellRange);
                        ReleaseComObject(cell);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(cells);
            }

            return grid;
        }

        private static void ExtractRequirementsFromTable(
            Word.Table table,
            string sectionNumber,
            string preferredPrefix,
            List<RequirementItem> requirements)
        {
            if (table == null || requirements == null) return;

            List<TableRowData> tableData = ParseTableData(table);
            TableHeaderInfo headerInfo = AnalyzeTableHeaderData(tableData);

            for (int rowIndex = 0; rowIndex < tableData.Count; rowIndex++)
            {
                TableRowData rowData = tableData[rowIndex];
                List<string> cellTexts = rowData.CellTexts;
                if (cellTexts.Count == 0) continue;

                string requirementId;
                int idCellIndex;
                if (!TryExtractRequirementId(cellTexts, headerInfo, preferredPrefix, out requirementId, out idCellIndex))
                {
                    continue;
                }

                string requirementName = PickRequirementNameFromRow(cellTexts, headerInfo, idCellIndex, requirementId);
                requirements.Add(CreateRequirementItem(
                    requirementId,
                    requirementName,
                    sectionNumber,
                    rowData.Start));
            }
        }

        private static TableHeaderInfo AnalyzeTableHeaderData(List<TableRowData> tableData)
        {
            TableHeaderInfo info = new TableHeaderInfo();
            int headerRows = Math.Min(2, tableData.Count);

            for (int rowIndex = 0; rowIndex < headerRows; rowIndex++)
            {
                List<string> headerTexts = tableData[rowIndex].CellTexts;
                for (int cellIndex = 0; cellIndex < headerTexts.Count; cellIndex++)
                {
                    string cellText = headerTexts[cellIndex];
                    if (string.IsNullOrWhiteSpace(cellText)) continue;

                    if (info.IdColumnIndex < 0 && ContainsAny(cellText, "唯一标识", "需求标识", "标识", "编号"))
                    {
                        info.IdColumnIndex = cellIndex;
                    }

                    if (info.NameColumnIndex < 0 && ContainsAny(cellText, "能力名称", "接口名称", "需求名称", "功能名称", "名称"))
                    {
                        info.NameColumnIndex = cellIndex;
                    }

                    if (ContainsAny(cellText, "能力名称", "能力需求描述"))
                    {
                        info.IsCapabilityTable = true;
                    }

                    if (ContainsAny(cellText, "接口名称", "用途", "来源", "接收者"))
                    {
                        info.IsInterfaceTable = true;
                    }

                    if (ContainsAny(cellText, "安全性", "可靠性"))
                    {
                        info.IsSafetyOrReliabilityTable = true;
                    }
                }
            }

            return info;
        }

        private static TableHeaderInfo AnalyzeTableHeader(Word.Table table)
        {
            return AnalyzeTableHeaderData(ParseTableData(table));
        }

        private static List<string> GetRowCellTexts(Word.Row row)
        {
            List<string> cellTexts = new List<string>();
            if (row == null)
            {
                return cellTexts;
            }

            int cellCount = 0;
            try
            {
                cellCount = row.Cells.Count;
            }
            catch
            {
                cellCount = 0;
            }

            for (int cellIndex = 1; cellIndex <= cellCount; cellIndex++)
            {
                Word.Cell cell = null;
                Word.Range cellRange = null;
                try
                {
                    cell = row.Cells[cellIndex];
                    cellRange = cell?.Range;
                    cellTexts.Add(NormalizeParagraphText(cellRange?.Text));
                }
                catch
                {
                    cellTexts.Add(string.Empty);
                }
                finally
                {
                    ReleaseComObject(cellRange);
                    ReleaseComObject(cell);
                }
            }

            return cellTexts;
        }

        private static bool TryExtractRequirementId(
            IReadOnlyList<string> cellTexts,
            TableHeaderInfo headerInfo,
            string preferredPrefix,
            out string requirementId,
            out int idCellIndex)
        {
            requirementId = string.Empty;
            idCellIndex = -1;
            if (cellTexts == null)
            {
                return false;
            }

            if (headerInfo != null && headerInfo.IdColumnIndex >= 0)
            {
                if (TryExtractRequirementIdFromColumnRange(
                    cellTexts,
                    preferredPrefix,
                    headerInfo.IdColumnIndex,
                    headerInfo.IdColumnIndex,
                    out requirementId,
                    out idCellIndex))
                {
                    return true;
                }
            }

            if (TryExtractRequirementIdFromColumnRange(
                cellTexts,
                preferredPrefix,
                0,
                Math.Min(1, cellTexts.Count - 1),
                out requirementId,
                out idCellIndex))
            {
                return true;
            }

            return TryExtractRequirementIdFromColumnRange(
                cellTexts,
                preferredPrefix,
                0,
                cellTexts.Count - 1,
                out requirementId,
                out idCellIndex);
        }

        private static bool TryExtractRequirementIdFromColumnRange(
            IReadOnlyList<string> cellTexts,
            string preferredPrefix,
            int startCol,
            int endCol,
            out string requirementId,
            out int idCellIndex)
        {
            requirementId = string.Empty;
            idCellIndex = -1;
            if (cellTexts == null || cellTexts.Count == 0)
            {
                return false;
            }

            startCol = Math.Max(0, startCol);
            endCol = Math.Min(endCol, cellTexts.Count - 1);
            if (startCol > endCol)
            {
                return false;
            }

            for (int i = startCol; i <= endCol; i++)
            {
                MatchCollection matches = RequirementIdRegex.Matches(cellTexts[i] ?? string.Empty);
                for (int matchIndex = 0; matchIndex < matches.Count; matchIndex++)
                {
                    Match match = matches[matchIndex];
                    if (!match.Success)
                    {
                        continue;
                    }

                    string id = match.Groups["id"].Value.Trim();
                    if (!IsRequirementIdMatched(id, preferredPrefix))
                    {
                        continue;
                    }

                    requirementId = id;
                    idCellIndex = i;
                    return true;
                }
            }

            return false;
        }

        private static string PickRequirementNameFromRow(
            IReadOnlyList<string> cellTexts,
            TableHeaderInfo headerInfo,
            int idCellIndex,
            string requirementId)
        {
            string candidate = GetRequirementNameCandidate(cellTexts, headerInfo?.NameColumnIndex ?? -1, idCellIndex, requirementId);
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                return candidate;
            }

            if (headerInfo != null && headerInfo.IsSafetyOrReliabilityTable)
            {
                for (int i = 0; i < Math.Min(2, cellTexts.Count); i++)
                {
                    candidate = GetRequirementNameCandidate(cellTexts, i, idCellIndex, requirementId);
                    if (!string.IsNullOrWhiteSpace(candidate))
                    {
                        return candidate;
                    }
                }
            }

            for (int i = 0; i < cellTexts.Count; i++)
            {
                candidate = GetRequirementNameCandidate(cellTexts, i, idCellIndex, requirementId);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            return "未命名需求";
        }

        private static string GetRequirementNameCandidate(
            IReadOnlyList<string> cellTexts,
            int candidateIndex,
            int idCellIndex,
            string requirementId)
        {
            if (cellTexts == null || candidateIndex < 0 || candidateIndex >= cellTexts.Count || candidateIndex == idCellIndex)
            {
                return string.Empty;
            }

            string text = RemoveRequirementIds(cellTexts[candidateIndex]);
            if (string.IsNullOrWhiteSpace(text) || LooksLikeTableHeaderText(text))
            {
                return string.Empty;
            }

            if (!string.IsNullOrWhiteSpace(requirementId) &&
                text.IndexOf(requirementId, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                text = RemoveRequirementIds(text);
            }

            if (string.IsNullOrWhiteSpace(text) || text.Length > 80)
            {
                return string.Empty;
            }

            return TrimRequirementNameFragment(text);
        }

        private static bool LooksLikeTableHeaderText(string text)
        {
            string normalized = NormalizeParagraphText(text);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return true;
            }

            return string.Equals(normalized, "唯一标识", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "需求标识", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "标识", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "编号", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "能力名称", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "接口名称", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "需求名称", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "功能名称", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "用途", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "来源", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "接收者", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "描述", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "说明", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(normalized, "备注", StringComparison.OrdinalIgnoreCase);
        }

        private static string RemoveRequirementIds(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            return TrimRequirementNameFragment(RequirementIdRegex.Replace(text, " "));
        }

        private static string FindNearestSectionNumber(Word.Document doc, int chapterStart, int chapterEnd, int targetStart)
        {
            string currentSectionNumber = string.Empty;
            Word.Range chapterRange = null;
            Word.Paragraphs paragraphs = null;
            try
            {
                chapterRange = doc.Range(chapterStart, chapterEnd);
                paragraphs = chapterRange.Paragraphs;
                if (paragraphs == null)
                {
                    return string.Empty;
                }

                int paragraphCount = paragraphs.Count;
                for (int i = 1; i <= paragraphCount; i++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                        if (paragraphRange == null || paragraphRange.Start >= targetStart)
                        {
                            break;
                        }

                        if (IsRangeInsideTable(paragraphRange))
                        {
                            continue;
                        }

                        string paragraphText = NormalizeParagraphText(paragraphRange.Text);
                        string updatedSectionNumber;
                        if (TryExtractSectionNumber(paragraphText, out updatedSectionNumber))
                        {
                            currentSectionNumber = updatedSectionNumber;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                    }
                }
            }
            finally
            {
                ReleaseComObject(paragraphs);
                ReleaseComObject(chapterRange);
            }

            return currentSectionNumber;
        }

        private static bool TryLocateSrsThirdChapterBounds(
            Word.Document doc,
            Action<string> progressReporter,
            out int chapterStart,
            out int chapterEnd)
        {
            chapterStart = 0;
            chapterEnd = 0;

            Word.Paragraphs paragraphs = null;
            Word.Range contentRange = null;
            try
            {
                paragraphs = doc.Paragraphs;
                contentRange = doc.Content;
                if (paragraphs == null || contentRange == null)
                {
                    return false;
                }

                int paragraphCount = paragraphs.Count;
                int progressInterval = paragraphCount > 2000 ? 200 : paragraphCount > 600 ? 80 : 25;
                bool chapterStarted = false;
                for (int i = 1; i <= paragraphCount; i++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                        string paragraphText = NormalizeParagraphText(paragraphRange?.Text);
                        if (string.IsNullOrWhiteSpace(paragraphText))
                        {
                            continue;
                        }

                        int chapterNumber;
                        if (!TryExtractSrsBoundaryChapterNumber(paragraphText, paragraph, out chapterNumber))
                        {
                            continue;
                        }

                        if (!chapterStarted && chapterNumber == 3)
                        {
                            chapterStart = paragraphRange?.End ?? 0;
                            chapterStarted = chapterStart > 0;
                            continue;
                        }

                        if (chapterStarted && chapterNumber >= 4)
                        {
                            chapterEnd = Math.Max(chapterStart, (paragraphRange?.Start ?? chapterStart));
                            break;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraphRange);
                        ReleaseComObject(paragraph);
                        if (i == 1 || i % progressInterval == 0 || i == paragraphCount)
                        {
                            ReportProgress(progressReporter, $"正在定位第三章：{i}/{paragraphCount}");
                        }
                    }
                }

                if (chapterStart <= 0)
                {
                    return false;
                }

                if (chapterEnd <= chapterStart)
                {
                    chapterEnd = contentRange.End;
                }

                return chapterEnd > chapterStart;
            }
            finally
            {
                ReleaseComObject(contentRange);
                ReleaseComObject(paragraphs);
            }
        }

        private static bool TryExtractSrsBoundaryChapterNumber(string text, Word.Paragraph paragraph, out int chapterNumber)
        {
            chapterNumber = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            Match chapterMatch = ChapterHeadingRegex.Match(text);
            if (chapterMatch.Success)
            {
                chapterNumber = ParseChapterNumberToken(chapterMatch.Groups["number"].Value);
                return chapterNumber == 3 || chapterNumber == 4;
            }

            Match simpleMatch = SimpleTopLevelHeadingRegex.Match(text);
            if (!simpleMatch.Success)
            {
                return false;
            }

            int parsedNumber = ParseChapterNumberToken(simpleMatch.Groups["number"].Value);
            if (parsedNumber != 3 && parsedNumber != 4)
            {
                return false;
            }

            if (!IsHeadingLikeParagraph(paragraph) &&
                !ContainsAny(text, "需求", "说明", "规格", "矩阵", "设计", "接口", "安全性", "可靠性"))
            {
                return false;
            }

            chapterNumber = parsedNumber;
            return true;
        }

        private static bool IsHeadingLikeParagraph(Word.Paragraph paragraph)
        {
            if (paragraph == null)
            {
                return false;
            }

            try
            {
                Word.WdOutlineLevel level = paragraph.OutlineLevel;
                return level >= Word.WdOutlineLevel.wdOutlineLevel1 &&
                       level <= Word.WdOutlineLevel.wdOutlineLevel3;
            }
            catch
            {
                return false;
            }
        }

        private static int ParseChapterNumberToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return 0;
            }

            int numericValue;
            if (int.TryParse(token.Trim(), out numericValue))
            {
                return numericValue;
            }

            Dictionary<char, int> map = new Dictionary<char, int>
            {
                ['零'] = 0,
                ['〇'] = 0,
                ['一'] = 1,
                ['壹'] = 1,
                ['二'] = 2,
                ['贰'] = 2,
                ['两'] = 2,
                ['三'] = 3,
                ['叁'] = 3,
                ['四'] = 4,
                ['肆'] = 4,
                ['五'] = 5,
                ['伍'] = 5,
                ['六'] = 6,
                ['陆'] = 6,
                ['七'] = 7,
                ['柒'] = 7,
                ['八'] = 8,
                ['捌'] = 8,
                ['九'] = 9,
                ['玖'] = 9,
                ['十'] = 10,
                ['拾'] = 10
            };

            string normalized = token.Trim();
            if (normalized.Length == 1)
            {
                int singleValue;
                if (map.TryGetValue(normalized[0], out singleValue))
                {
                    return singleValue;
                }
            }

            int tenIndex = normalized.IndexOf('十');
            if (tenIndex < 0)
            {
                tenIndex = normalized.IndexOf('拾');
            }

            if (tenIndex >= 0)
            {
                string left = normalized.Substring(0, tenIndex);
                string right = normalized.Substring(tenIndex + 1);
                int tens = string.IsNullOrWhiteSpace(left) ? 1 : ParseChapterNumberToken(left);
                int units = string.IsNullOrWhiteSpace(right) ? 0 : ParseChapterNumberToken(right);
                return tens * 10 + units;
            }

            return 0;
        }

        private static bool TryExtractSectionHeading(string text, out string sectionNumber, out string sectionTitle)
        {
            sectionNumber = string.Empty;
            sectionTitle = string.Empty;
            if (string.IsNullOrWhiteSpace(text) || text.Length > 120)
            {
                return false;
            }

            Match numberedMatch = NumberedHeadingRegex.Match(text);
            if (numberedMatch.Success)
            {
                sectionNumber = numberedMatch.Groups["number"].Value.Trim().Replace('．', '.');
                sectionTitle = TrimRequirementNameFragment(numberedMatch.Groups["title"].Value);
                return !string.IsNullOrWhiteSpace(sectionNumber);
            }

            Match chapterMatch = ChapterHeadingRegex.Match(text);
            if (chapterMatch.Success)
            {
                sectionNumber = chapterMatch.Groups["number"].Value.Trim();
                string titlePart = text.Substring(chapterMatch.Length);
                sectionTitle = TrimRequirementNameFragment(titlePart);
                return !string.IsNullOrWhiteSpace(sectionNumber);
            }

            return false;
        }

        private static string ResolveParagraphRequirementName(
            Word.Paragraphs paragraphs,
            int paragraphIndex,
            string currentSectionTitle,
            string paragraphText,
            Match currentMatch,
            MatchCollection allMatches,
            string currentRequirementId)
        {
            string name = ExtractRequirementName(paragraphText, currentMatch, allMatches);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            if (!string.IsNullOrWhiteSpace(currentSectionTitle))
            {
                return currentSectionTitle;
            }

            name = TryGetPreviousMeaningfulParagraphText(paragraphs, paragraphIndex, "SRS", currentRequirementId);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            name = TryGetSupplementalRequirementName(paragraphs, paragraphIndex, "SRS", currentRequirementId);
            return name;
        }

        private static string TryGetPreviousMeaningfulParagraphText(
            Word.Paragraphs paragraphs,
            int paragraphIndex,
            string preferredPrefix,
            string currentRequirementId)
        {
            if (paragraphs == null)
            {
                return string.Empty;
            }

            for (int offset = 1; offset <= 2; offset++)
            {
                int previousIndex = paragraphIndex - offset;
                if (previousIndex < 1)
                {
                    break;
                }

                Word.Paragraph paragraph = null;
                Word.Range paragraphRange = null;
                try
                {
                    paragraph = paragraphs[previousIndex];
                    paragraphRange = paragraph?.Range;
                    string text = NormalizeParagraphText(paragraphRange?.Text);
                    if (string.IsNullOrWhiteSpace(text) || ContainsRequirementId(text, preferredPrefix))
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(currentRequirementId) &&
                        text.IndexOf(currentRequirementId, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        continue;
                    }

                    string sectionNumber;
                    string sectionTitle;
                    if (TryExtractSectionHeading(text, out sectionNumber, out sectionTitle) &&
                        !string.IsNullOrWhiteSpace(sectionTitle))
                    {
                        return sectionTitle;
                    }

                    string trimmed = TrimRequirementNameFragment(text);
                    if (!string.IsNullOrWhiteSpace(trimmed))
                    {
                        return trimmed;
                    }
                }
                finally
                {
                    ReleaseComObject(paragraphRange);
                    ReleaseComObject(paragraph);
                }
            }

            return string.Empty;
        }

        private static RequirementTrackingDocumentSnapshot CreateSnapshot(Word.Document doc)
        {
            return new RequirementTrackingDocumentSnapshot
            {
                DisplayName = TryGetDocumentDisplayName(doc),
                FullName = TryGetDocumentFullName(doc)
            };
        }

        private static RequirementTrackingDocumentKind DetectDocumentKind(Word.Document doc)
        {
            string sample = BuildFastDocumentKindSample(doc);
            if (ContainsAny(sample, "软件需求规格说明", "需求规格说明", "SRS"))
            {
                return RequirementTrackingDocumentKind.RequirementSpecification;
            }

            if (ContainsAny(sample, "系统规格说明", "系统/子系统规格说明", "系统子系统规格说明", "SSS"))
            {
                return RequirementTrackingDocumentKind.SystemSpecification;
            }

            if (ContainsAny(sample, "软件设计说明", "软件设计描述", "SDD", "SDS"))
            {
                return RequirementTrackingDocumentKind.SoftwareDesignDescription;
            }

            if (ContainsAny(sample, "软件测试说明", "软件测试描述", "测试说明", "STD", "STS"))
            {
                return RequirementTrackingDocumentKind.SoftwareTestDescription;
            }

            return RequirementTrackingDocumentKind.Unknown;
        }

        private static RequirementItem CreateRequirementItem(string id, string name, string sectionNumber, int start)
        {
            string normalizedId = NormalizeRequirementId(id);
            string normalizedName = string.IsNullOrWhiteSpace(name) ? "未命名需求" : TrimRequirementNameFragment(name);
            return new RequirementItem
            {
                Id = string.IsNullOrWhiteSpace(normalizedId) ? id : normalizedId,
                Name = string.IsNullOrWhiteSpace(normalizedName) ? "未命名需求" : normalizedName,
                SectionNumber = sectionNumber ?? string.Empty,
                BookmarkOrRange = start,
                Start = start
            };
        }

        private static List<RequirementItem> FinalizeRequirements(IEnumerable<RequirementItem> requirements)
        {
            return (requirements ?? Enumerable.Empty<RequirementItem>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                .GroupBy(item => new { Id = item.Id ?? string.Empty, Start = item.Start })
                .Select(group => group.First())
                .OrderBy(item => item.Start)
                .ThenBy(item => item.Id ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static bool IsRangeInsideTable(Word.Range range)
        {
            if (range == null)
            {
                return false;
            }

            try
            {
                return range.Tables.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        private static void ReportProgress(Action<string> progressReporter, string message)
        {
            if (progressReporter == null || string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            progressReporter(message);
        }

        private sealed class SectionHeadingSnapshot
        {
            internal int Start { get; set; }

            internal string SectionNumber { get; set; }

            internal string SectionTitle { get; set; }
        }

        private sealed class TextLineSnapshot
        {
            internal int StartIndex { get; set; }

            internal string Text { get; set; }

            internal bool IsTableCell { get; set; }

            internal string SectionNumber { get; set; }

            internal string SectionTitle { get; set; }
        }

        private sealed class TableHeaderInfo
        {
            internal int IdColumnIndex { get; set; } = -1;

            internal int NameColumnIndex { get; set; } = -1;

            internal bool IsCapabilityTable { get; set; }

            internal bool IsInterfaceTable { get; set; }

            internal bool IsSafetyOrReliabilityTable { get; set; }
        }

        private static bool TryExtractSectionNumber(string text, out string sectionNumber)
        {
            sectionNumber = string.Empty;
            if (string.IsNullOrWhiteSpace(text) || text.Length > 120)
            {
                return false;
            }

            Match chapterMatch = ChapterHeadingRegex.Match(text);
            if (chapterMatch.Success)
            {
                sectionNumber = chapterMatch.Groups["number"].Value.Trim();
                return !string.IsNullOrWhiteSpace(sectionNumber);
            }

            Match numberedMatch = NumberedHeadingRegex.Match(text);
            if (numberedMatch.Success)
            {
                sectionNumber = numberedMatch.Groups["number"].Value.Trim();
                return !string.IsNullOrWhiteSpace(sectionNumber);
            }

            return false;
        }

        private static string ExtractRequirementName(string paragraphText, Match currentMatch, MatchCollection allMatches)
        {
            if (string.IsNullOrWhiteSpace(paragraphText) || currentMatch == null)
            {
                return string.Empty;
            }

            int startIndex = currentMatch.Index + currentMatch.Length;
            int endIndex = paragraphText.Length;
            foreach (Match match in allMatches)
            {
                if (match == null || !match.Success || match.Index <= currentMatch.Index)
                {
                    continue;
                }

                endIndex = match.Index;
                break;
            }

            if (endIndex <= startIndex || startIndex >= paragraphText.Length)
            {
                return string.Empty;
            }

            string fragment = paragraphText.Substring(startIndex, endIndex - startIndex);
            return TrimRequirementNameFragment(fragment);
        }

        private static string TryGetSupplementalRequirementName(Word.Paragraphs paragraphs, int paragraphIndex, string preferredPrefix, string currentRequirementId)
        {
            if (paragraphs == null)
            {
                return string.Empty;
            }

            int paragraphCount = 0;
            try
            {
                paragraphCount = paragraphs.Count;
            }
            catch
            {
                paragraphCount = 0;
            }

            for (int offset = 1; offset <= 2; offset++)
            {
                int nextIndex = paragraphIndex + offset;
                if (nextIndex > paragraphCount)
                {
                    break;
                }

                Word.Paragraph nextParagraph = null;
                Word.Range nextRange = null;
                try
                {
                    nextParagraph = paragraphs[nextIndex];
                    nextRange = nextParagraph?.Range;
                    string nextText = NormalizeParagraphText(nextRange?.Text);
                    if (string.IsNullOrWhiteSpace(nextText))
                    {
                        continue;
                    }

                    if (ContainsRequirementId(nextText, preferredPrefix) || IsHeadingLikeText(nextText))
                    {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(currentRequirementId) &&
                        nextText.IndexOf(currentRequirementId, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        continue;
                    }

                    string trimmed = TrimRequirementNameFragment(nextText);
                    if (!string.IsNullOrWhiteSpace(trimmed))
                    {
                        return trimmed;
                    }
                }
                finally
                {
                    ReleaseComObject(nextRange);
                    ReleaseComObject(nextParagraph);
                }
            }

            return string.Empty;
        }

        private static bool ContainsRequirementId(string text, string preferredPrefix)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            MatchCollection matches = RequirementIdRegex.Matches(text);
            foreach (Match match in matches)
            {
                if (!match.Success)
                {
                    continue;
                }

                string id = match.Groups["id"].Value.Trim();
                if (IsRequirementIdMatched(id, preferredPrefix))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsHeadingLikeText(string text)
        {
            return TryExtractSectionNumber(text, out _);
        }

        private static string TrimRequirementNameFragment(string fragment)
        {
            if (string.IsNullOrWhiteSpace(fragment))
            {
                return string.Empty;
            }

            string trimmed = fragment.Trim();
            trimmed = trimmed.Trim('：', ':', '，', ',', '。', '.', '；', ';', '、', '-', '—', '－', ' ', '\t', '(', ')', '（', '）', '【', '】', '[', ']');
            trimmed = Regex.Replace(trimmed, @"\s+", " ");
            return trimmed.Trim();
        }

        private static string NormalizeParagraphText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            string normalized = text.Replace("\r", string.Empty).Replace("\a", string.Empty);
            normalized = normalized.Replace('\u0007'.ToString(), string.Empty);
            normalized = Regex.Replace(normalized, @"\s+", " ");
            return normalized.Trim();
        }

        private static string DetectPreferredRequirementPrefix(Word.Document doc)
        {
            string sample = BuildSampleText(doc);
            if (ContainsAny(sample, "软件需求规格说明", "需求规格说明", "SRS"))
            {
                return "SRS";
            }

            if (ContainsAny(sample, "软件研制任务书", "任务书", "系统/子系统规格说明", "系统子系统规格说明", "SSS"))
            {
                return "SSS";
            }

            return string.Empty;
        }

        private static bool IsRequirementIdMatched(string id, string preferredPrefix)
        {
            string normalizedId = NormalizeRequirementId(id);
            if (string.IsNullOrWhiteSpace(normalizedId))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(preferredPrefix))
            {
                return true;
            }

            return normalizedId.IndexOf($"-{preferredPrefix}-", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   normalizedId.StartsWith($"{preferredPrefix}-", StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeRequirementId(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return string.Empty;
            }

            string normalized = Regex.Replace(id.Trim(), @"\s*[-－–—]\s*", "-");
            normalized = Regex.Replace(normalized, @"\s+", string.Empty);
            return normalized.ToUpperInvariant();
        }

        private static string ExtractRequirementIdSuffix(string id, string preferredPrefix)
        {
            string normalizedId = NormalizeRequirementId(id);
            if (string.IsNullOrWhiteSpace(normalizedId) || string.IsNullOrWhiteSpace(preferredPrefix))
            {
                return string.Empty;
            }

            MatchCollection matches = Regex.Matches(
                normalizedId,
                $@"{Regex.Escape(preferredPrefix)}-\d+",
                RegexOptions.IgnoreCase);
            return matches.Count == 0 ? string.Empty : matches[matches.Count - 1].Value.ToUpperInvariant();
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

            Word.Paragraphs paragraphs = null;
            try
            {
                paragraphs = doc.Paragraphs;
                if (paragraphs == null)
                {
                    return builder.ToString();
                }

                int paragraphCount = paragraphs.Count;
                int maxParagraphs = Math.Min(paragraphCount, 80);
                for (int i = 1; i <= maxParagraphs && builder.Length < 10000; i++)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range range = null;
                    try
                    {
                        paragraph = paragraphs[i];
                        range = paragraph?.Range;
                        string text = NormalizeParagraphText(range?.Text);
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            continue;
                        }

                        builder.AppendLine(text);
                    }
                    finally
                    {
                        ReleaseComObject(range);
                        ReleaseComObject(paragraph);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(paragraphs);
            }

            return builder.ToString();
        }

        private static string BuildFastDocumentKindSample(Word.Document doc)
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

            Word.Range contentRange = null;
            try
            {
                contentRange = doc.Content;
                string text = contentRange?.Text ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    builder.AppendLine(text);
                }
            }
            catch
            {
                builder.AppendLine(BuildSampleText(doc));
            }
            finally
            {
                ReleaseComObject(contentRange);
            }

            return builder.ToString();
        }

        private static string TryGetDocumentDisplayName(Word.Document doc)
        {
            if (doc == null)
            {
                return "当前文档";
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(doc.Name))
                {
                    return doc.Name;
                }
            }
            catch
            {
            }

            return "当前文档";
        }

        private static string TryGetDocumentFullName(Word.Document doc)
        {
            if (doc == null)
            {
                return string.Empty;
            }

            try
            {
                return doc.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool ContainsAny(string text, params string[] keywords)
        {
            if (string.IsNullOrWhiteSpace(text) || keywords == null || keywords.Length == 0)
            {
                return false;
            }

            foreach (string keyword in keywords)
            {
                if (string.IsNullOrWhiteSpace(keyword))
                {
                    continue;
                }

                if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
            }
        }
    }
}
