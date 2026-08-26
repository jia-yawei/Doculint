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
        private const string CaptionReferenceBookmarkPrefix = "DocuLintCaptionRef_";

        internal enum CaptionReferenceKind
        {
            Number,
            FullCaption,
            PageNumber
        }

        private sealed class CaptionReferenceTarget
        {
            public int Start { get; set; }
            public Word.Paragraph Paragraph { get; set; }
            public string Text { get; set; }
            public string ReferenceType { get; set; }
            public int ReferenceItem { get; set; }
        }

        private void button28_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertNearestCaptionReference(searchForward: false);
        }

        private void button29_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertNearestCaptionReference(searchForward: true);
        }

        private void button31_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertSelectedCaptionReference();
        }

        private static void ExecuteInsertNearestCaptionReference(bool searchForward)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                Word.Selection selection = app?.Selection;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                if (selection?.Range == null)
                {
                    MessageBox.Show("请先把光标放到要插入引用的位置。", "文档不加班");
                    return;
                }

                List<CaptionReferenceTarget> targets = CollectCaptionReferenceTargets(doc);
                if (targets.Count == 0)
                {
                    MessageBox.Show("当前文档中没有可引用的题注。", "文档不加班");
                    return;
                }

                int cursorPosition = selection.Range.Start;
                CaptionReferenceTarget target = searchForward
                    ? targets.FirstOrDefault(item => item.Start > cursorPosition)
                    : targets.LastOrDefault(item => item.Start < cursorPosition);
                if (target == null)
                {
                    MessageBox.Show(
                        searchForward ? "光标后面没有找到可引用的题注。" : "光标前面没有找到可引用的题注。",
                        "文档不加班");
                    return;
                }

                InsertNativeCaptionReference(selection, target, CaptionReferenceKind.Number);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入题注引用失败: {ex.Message}", "文档不加班");
            }
        }

        private static void ExecuteInsertSelectedCaptionReference()
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                Word.Selection selection = app?.Selection;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                if (selection?.Range == null)
                {
                    MessageBox.Show("请先把光标放到要插入引用的位置。", "文档不加班");
                    return;
                }

                List<CaptionReferenceTarget> targets = CollectCaptionReferenceTargets(doc);
                if (targets.Count == 0)
                {
                    MessageBox.Show("当前文档中没有可引用的题注。", "文档不加班");
                    return;
                }

                using (CaptionReferencePickerForm form = new CaptionReferencePickerForm(
                    targets.Select(item => item.Text).ToList(),
                    CaptionReferenceKind.Number))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    CaptionReferenceTarget target = targets[form.SelectedIndex];
                    InsertNativeCaptionReference(selection, target, form.SelectedReferenceKind);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入题注引用失败: {ex.Message}", "文档不加班");
            }
        }

        private static List<CaptionReferenceTarget> CollectCaptionReferenceTargets(Word.Document doc)
        {
            List<CaptionReferenceTarget> targets = new List<CaptionReferenceTarget>();
            if (doc == null)
            {
                return targets;
            }

            Dictionary<string, int> referenceItems = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (CaptionListEntry entry in CollectCaptionListEntries(doc))
            {
                ThrowIfOperationCancelled();
                Word.Paragraph paragraph = GetParagraphAtStart(doc, entry?.Start ?? 0);
                if (paragraph?.Range == null)
                {
                    continue;
                }

                string referenceType = GetNativeCaptionReferenceType(entry.Text);
                if (string.IsNullOrWhiteSpace(referenceType))
                {
                    continue;
                }

                int referenceItem = referenceItems.TryGetValue(referenceType, out int currentItem)
                    ? currentItem + 1
                    : 1;
                referenceItems[referenceType] = referenceItem;

                targets.Add(new CaptionReferenceTarget
                {
                    Start = paragraph.Range.Start,
                    Paragraph = paragraph,
                    Text = entry.Text,
                    ReferenceType = referenceType,
                    ReferenceItem = referenceItem
                });
            }

            return targets
                .OrderBy(item => item.Start)
                .ToList();
        }

        private static Word.Paragraph GetParagraphAtStart(Word.Document doc, int start)
        {
            if (doc == null || start < 0)
            {
                return null;
            }

            try
            {
                int safeStart = Math.Min(start, doc.Content.End);
                int safeEnd = Math.Min(doc.Content.End, safeStart + 1);
                Word.Range range = doc.Range(safeStart, safeEnd);
                Word.Paragraphs paragraphs = range?.Paragraphs;
                if (paragraphs == null || paragraphs.Count < 1)
                {
                    return null;
                }

                return paragraphs[1];
            }
            catch
            {
                return null;
            }
        }

        private static void InsertNativeCaptionReference(
            Word.Selection selection,
            CaptionReferenceTarget target,
            CaptionReferenceKind referenceKind)
        {
            if (selection?.Range == null
                || target == null
                || string.IsNullOrWhiteSpace(target.ReferenceType)
                || target.ReferenceItem <= 0)
            {
                return;
            }

            Word.Range insertRange = selection.Range.Duplicate;
            object referenceType = target.ReferenceType;
            object referenceItem = target.ReferenceItem;
            object insertAsHyperlink = true;
            object includePosition = false;
            object separateNumbers = false;
            object separatorString = Type.Missing;

            try
            {
                insertRange.InsertCrossReference(
                    ref referenceType,
                    ToNativeReferenceKind(referenceKind),
                    ref referenceItem,
                    ref insertAsHyperlink,
                    ref includePosition,
                    ref separateNumbers,
                    ref separatorString);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Word 原生交叉引用插入失败：{ex.Message}", ex);
            }

            Word.Range insertedRange = insertRange;
            if (insertedRange.End <= insertedRange.Start)
            {
                insertedRange = selection.Range.Duplicate;
            }

            try
            {
                ApplyCaptionReferenceFormatting(insertedRange);
            }
            catch
            {
            }

            insertedRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertedRange.Select();
        }

        private static void ApplyCaptionReferenceFormatting(Word.Range targetRange)
        {
            if (targetRange?.Font == null)
            {
                return;
            }

            try
            {
                targetRange.Font.NameFarEast = "宋体";
            }
            catch
            {
            }

            try
            {
                targetRange.Font.NameAscii = "宋体";
            }
            catch
            {
            }

            try
            {
                targetRange.Font.NameOther = "宋体";
            }
            catch
            {
            }

            try
            {
                targetRange.Font.Name = "宋体";
            }
            catch
            {
            }

            try
            {
                targetRange.Font.Size = 12f;
            }
            catch
            {
            }
        }

        private static Word.WdReferenceKind ToNativeReferenceKind(CaptionReferenceKind referenceKind)
        {
            switch (referenceKind)
            {
                case CaptionReferenceKind.FullCaption:
                    return Word.WdReferenceKind.wdEntireCaption;
                case CaptionReferenceKind.PageNumber:
                    return Word.WdReferenceKind.wdPageNumber;
                default:
                    return Word.WdReferenceKind.wdOnlyLabelAndNumber;
            }
        }

        private static string GetNativeCaptionReferenceType(string captionText)
        {
            string normalized = NormalizeCaptionParagraphText(captionText);
            Match match = CaptionPrefixRegex.Match(normalized ?? string.Empty);
            if (!match.Success)
            {
                return null;
            }

            string prefix = match.Groups["prefix"].Value.Trim();
            if (string.Equals(prefix, "Figure", StringComparison.OrdinalIgnoreCase))
            {
                return "Figure";
            }

            if (string.Equals(prefix, "Table", StringComparison.OrdinalIgnoreCase))
            {
                return "Table";
            }

            return prefix;
        }

        private static void TryApplyCharacterFormatting(Word.Range sourceRange, Word.Range targetRange)
        {
            if (sourceRange?.Font == null || targetRange?.Font == null)
            {
                return;
            }

            try
            {
                targetRange.Font.NameFarEast = sourceRange.Font.NameFarEast;
            }
            catch
            {
            }

            try
            {
                targetRange.Font.NameAscii = sourceRange.Font.NameAscii;
            }
            catch
            {
            }

            try
            {
                targetRange.Font.NameOther = sourceRange.Font.NameOther;
            }
            catch
            {
            }

            try
            {
                targetRange.Font.Name = sourceRange.Font.Name;
            }
            catch
            {
            }

            try
            {
                targetRange.Font.Size = sourceRange.Font.Size;
            }
            catch
            {
            }

            try
            {
                targetRange.Font.Bold = sourceRange.Font.Bold;
            }
            catch
            {
            }

            try
            {
                targetRange.Font.Italic = sourceRange.Font.Italic;
            }
            catch
            {
            }

            try
            {
                targetRange.Font.Color = sourceRange.Font.Color;
            }
            catch
            {
            }

            try
            {
                targetRange.Font.Underline = sourceRange.Font.Underline;
            }
            catch
            {
            }
        }

        private static string EnsureCaptionReferenceBookmark(
            Word.Document doc,
            Word.Paragraph paragraph,
            CaptionReferenceKind referenceKind,
            string preferredName)
        {
            if (doc == null || paragraph?.Range == null)
            {
                return null;
            }

            Word.Range bookmarkRange = GetCaptionReferenceRange(paragraph, referenceKind);
            if (bookmarkRange == null)
            {
                return null;
            }

            List<string> existingNames = GetCaptionReferenceBookmarkNames(doc, paragraph, referenceKind);
            string bookmarkName = !string.IsNullOrWhiteSpace(preferredName)
                ? preferredName
                : existingNames.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(bookmarkName))
            {
                bookmarkName = GenerateCaptionReferenceBookmarkName(doc, referenceKind);
            }

            foreach (string existingName in existingNames)
            {
                ThrowIfOperationCancelled();
                if (!string.Equals(existingName, bookmarkName, StringComparison.OrdinalIgnoreCase))
                {
                    TryDeleteBookmark(doc, existingName);
                }
            }

            TryDeleteBookmark(doc, bookmarkName);
            try
            {
                doc.Bookmarks.Add(bookmarkName, bookmarkRange);
                return bookmarkName;
            }
            catch
            {
                return null;
            }
        }

        private static string FindCaptionReferenceBookmarkName(Word.Document doc, Word.Paragraph paragraph)
        {
            return GetCaptionReferenceBookmarkNames(doc, paragraph, CaptionReferenceKind.Number).FirstOrDefault();
        }

        private static List<string> GetCaptionReferenceBookmarkNames(
            Word.Document doc,
            Word.Paragraph paragraph,
            CaptionReferenceKind referenceKind)
        {
            List<string> names = new List<string>();
            if (doc == null || paragraph?.Range == null)
            {
                return names;
            }

            try
            {
                Word.Bookmarks bookmarks = paragraph.Range.Bookmarks;
                if (bookmarks == null)
                {
                    return names;
                }

                foreach (Word.Bookmark bookmark in bookmarks)
                {
                    ThrowIfOperationCancelled();
                    string name = bookmark?.Name ?? string.Empty;
                    if (!IsCaptionReferenceBookmarkForKind(name, referenceKind))
                    {
                        continue;
                    }

                    names.Add(name);
                }
            }
            catch
            {
            }

            return names
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static Word.Range GetCaptionReferenceRange(Word.Paragraph paragraph, CaptionReferenceKind referenceKind)
        {
            switch (referenceKind)
            {
                case CaptionReferenceKind.FullCaption:
                case CaptionReferenceKind.PageNumber:
                    return GetCaptionReferenceFullRange(paragraph);
                default:
                    return GetCaptionReferencePrefixRange(paragraph);
            }
        }

        private static Word.Range GetCaptionReferenceFullRange(Word.Paragraph paragraph)
        {
            if (paragraph?.Range == null)
            {
                return null;
            }

            Word.Range range = paragraph.Range.Duplicate;
            while (range.End > range.Start)
            {
                ThrowIfOperationCancelled();
                string tail = range.Text?.Substring(Math.Max(0, range.Text.Length - 1)) ?? string.Empty;
                if (tail != "\r" && tail != "\a")
                {
                    break;
                }

                range.End -= 1;
            }

            return range.End > range.Start ? range : null;
        }

        private static Word.Range GetCaptionReferencePrefixRange(Word.Paragraph paragraph)
        {
            if (paragraph?.Range == null)
            {
                return null;
            }

            string normalized = NormalizeCaptionParagraphText(paragraph.Range.Text);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return null;
            }

            Match match = CaptionPrefixRegex.Match(normalized);
            if (!match.Success || match.Length <= 0)
            {
                return null;
            }

            Word.Range range = paragraph.Range.Duplicate;
            if (range.End > range.Start)
            {
                range.End -= 1;
            }

            int targetEnd = Math.Min(range.End, range.Start + match.Length);
            if (targetEnd <= range.Start)
            {
                return null;
            }

            range.End = targetEnd;
            return range;
        }

        private static string GenerateCaptionReferenceBookmarkName(Word.Document doc, CaptionReferenceKind referenceKind)
        {
            string prefix = GetCaptionReferenceBookmarkPrefix(referenceKind);
            for (int i = 0; i < 200; i++)
            {
                ThrowIfOperationCancelled();
                string candidate = prefix + Guid.NewGuid().ToString("N");
                try
                {
                    if (doc?.Bookmarks == null || !doc.Bookmarks.Exists(candidate))
                    {
                        return candidate;
                    }
                }
                catch
                {
                    return candidate;
                }
            }

            return prefix + DateTime.UtcNow.Ticks.ToString();
        }

        private static string GetCaptionReferenceBookmarkPrefix(CaptionReferenceKind referenceKind)
        {
            switch (referenceKind)
            {
                case CaptionReferenceKind.FullCaption:
                    return CaptionReferenceBookmarkPrefix + "Full_";
                case CaptionReferenceKind.PageNumber:
                    return CaptionReferenceBookmarkPrefix + "Page_";
                default:
                    return CaptionReferenceBookmarkPrefix + "Number_";
            }
        }

        private static bool IsCaptionReferenceBookmarkForKind(string bookmarkName, CaptionReferenceKind referenceKind)
        {
            if (string.IsNullOrWhiteSpace(bookmarkName))
            {
                return false;
            }

            if (referenceKind != CaptionReferenceKind.Number)
            {
                return bookmarkName.StartsWith(GetCaptionReferenceBookmarkPrefix(referenceKind), StringComparison.OrdinalIgnoreCase);
            }

            if (!bookmarkName.StartsWith(CaptionReferenceBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return !bookmarkName.StartsWith(GetCaptionReferenceBookmarkPrefix(CaptionReferenceKind.FullCaption), StringComparison.OrdinalIgnoreCase)
                && !bookmarkName.StartsWith(GetCaptionReferenceBookmarkPrefix(CaptionReferenceKind.PageNumber), StringComparison.OrdinalIgnoreCase);
        }

        private static void TryDeleteBookmark(Word.Document doc, string bookmarkName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(bookmarkName))
            {
                return;
            }

            try
            {
                if (doc.Bookmarks.Exists(bookmarkName))
                {
                    doc.Bookmarks[bookmarkName].Delete();
                }
            }
            catch
            {
            }
        }

        private static void UpdateCaptionReferenceFields(Word.Document doc)
        {
            if (doc == null)
            {
                return;
            }

            CollectCaptionReferenceTargets(doc);

            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                ThrowIfOperationCancelled();
                UpdateCaptionReferenceFieldsInRange(storyRange);
            }
        }

        private static void UpdateCaptionReferenceFieldsInRange(Word.Range storyRange)
        {
            if (storyRange?.Fields == null)
            {
                return;
            }

            try
            {
                foreach (Word.Field field in storyRange.Fields)
                {
                    ThrowIfOperationCancelled();
                    if (field == null)
                    {
                        continue;
                    }

                    if (field.Type != Word.WdFieldType.wdFieldRef)
                    {
                        continue;
                    }

                    string codeText = null;
                    try
                    {
                        codeText = field.Code?.Text;
                    }
                    catch
                    {
                        codeText = null;
                    }

                    if (string.IsNullOrWhiteSpace(codeText)
                        || codeText.IndexOf(CaptionReferenceBookmarkPrefix, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    try
                    {
                        field.Update();
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
    }
}
