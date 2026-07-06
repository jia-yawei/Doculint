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

        private sealed class CaptionReferenceTarget
        {
            public int Start { get; set; }
            public Word.Paragraph Paragraph { get; set; }
            public string BookmarkName { get; set; }
        }

        private void button28_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertNearestCaptionReference(searchForward: false);
        }

        private void button29_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertNearestCaptionReference(searchForward: true);
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

                InsertCaptionReferenceField(selection, target.BookmarkName);
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

            foreach (CaptionListEntry entry in CollectCaptionListEntries(doc))
            {
                Word.Paragraph paragraph = GetParagraphAtStart(doc, entry?.Start ?? 0);
                if (paragraph?.Range == null)
                {
                    continue;
                }

                string bookmarkName = EnsureCaptionReferenceBookmark(doc, paragraph, null);
                if (string.IsNullOrWhiteSpace(bookmarkName))
                {
                    continue;
                }

                targets.Add(new CaptionReferenceTarget
                {
                    Start = paragraph.Range.Start,
                    Paragraph = paragraph,
                    BookmarkName = bookmarkName
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

        private static void InsertCaptionReferenceField(Word.Selection selection, string bookmarkName)
        {
            if (selection?.Range == null || string.IsNullOrWhiteSpace(bookmarkName))
            {
                return;
            }

            Word.Range insertRange = selection.Range.Duplicate;
            Word.Range formatSourceRange = selection.Range.Duplicate;
            Word.Field field = null;
            try
            {
                field = insertRange.Fields.Add(
                    insertRange,
                    Word.WdFieldType.wdFieldEmpty,
                    $" REF {bookmarkName} \\h ",
                    false);
            }
            catch
            {
                field = null;
            }

            if (field == null)
            {
                throw new InvalidOperationException("题注引用域插入失败。");
            }

            try
            {
                field.Update();
            }
            catch
            {
            }

            Word.Range endRange = field.Result?.Duplicate ?? insertRange;
            TryApplyCharacterFormatting(formatSourceRange, endRange);
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            endRange.Select();
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

        private static string EnsureCaptionReferenceBookmark(Word.Document doc, Word.Paragraph paragraph, string preferredName)
        {
            if (doc == null || paragraph?.Range == null)
            {
                return null;
            }

            Word.Range prefixRange = GetCaptionReferencePrefixRange(paragraph);
            if (prefixRange == null)
            {
                return null;
            }

            List<string> existingNames = GetCaptionReferenceBookmarkNames(doc, paragraph);
            string bookmarkName = !string.IsNullOrWhiteSpace(preferredName)
                ? preferredName
                : existingNames.FirstOrDefault();
            if (string.IsNullOrWhiteSpace(bookmarkName))
            {
                bookmarkName = GenerateCaptionReferenceBookmarkName(doc);
            }

            foreach (string existingName in existingNames)
            {
                if (!string.Equals(existingName, bookmarkName, StringComparison.OrdinalIgnoreCase))
                {
                    TryDeleteBookmark(doc, existingName);
                }
            }

            TryDeleteBookmark(doc, bookmarkName);
            try
            {
                doc.Bookmarks.Add(bookmarkName, prefixRange);
                return bookmarkName;
            }
            catch
            {
                return null;
            }
        }

        private static string FindCaptionReferenceBookmarkName(Word.Document doc, Word.Paragraph paragraph)
        {
            return GetCaptionReferenceBookmarkNames(doc, paragraph).FirstOrDefault();
        }

        private static List<string> GetCaptionReferenceBookmarkNames(Word.Document doc, Word.Paragraph paragraph)
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
                    string name = bookmark?.Name ?? string.Empty;
                    if (!name.StartsWith(CaptionReferenceBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
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

        private static string GenerateCaptionReferenceBookmarkName(Word.Document doc)
        {
            for (int i = 0; i < 200; i++)
            {
                string candidate = CaptionReferenceBookmarkPrefix + Guid.NewGuid().ToString("N");
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

            return CaptionReferenceBookmarkPrefix + DateTime.UtcNow.Ticks.ToString();
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
