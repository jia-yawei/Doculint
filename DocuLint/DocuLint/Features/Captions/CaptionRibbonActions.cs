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
        private static readonly Regex CaptionPrefixRegex = new Regex(
            @"^(?<prefix>(?:\u56FE|\u8868|Figure|Table))\s*[\.:：-]?\s*(?<number>[0-9\uFF10-\uFF19\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D\u5341\u767E\u5343]+(?:\s*[\.．\-—]\s*[0-9\uFF10-\uFF19\u4E00\u4E8C\u4E09\u56DB\u4E94\u516D\u4E03\u516B\u4E5D\u5341\u767E\u5343]+)*)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

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

        private void btnListCaptions_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("当前没有活动文档。", "文档不加班");
                return;
            }

            bool oldScreenUpdating = app.ScreenUpdating;
            app.ScreenUpdating = false;

            try
            {
                List<CaptionListEntry> entries = CollectCaptionListEntries(doc);
                Globals.ThisAddIn.ShowCaptionListPane(doc, entries);

                if (entries.Count == 0)
                {
                    MessageBox.Show("未找到任何图注或表注段落。", "文档不加班");
                }
            }
            finally
            {
                app.ScreenUpdating = oldScreenUpdating;
            }
        }

        internal static List<CaptionListEntry> CollectCaptionListEntries(Word.Document doc)
        {
            Dictionary<int, CaptionListEntry> entryMap = new Dictionary<int, CaptionListEntry>();
            if (doc == null)
            {
                return new List<CaptionListEntry>();
            }

            int scanStart = GetScanStartAfterToc(doc);
            try
            {
                foreach (Word.InlineShape inlineShape in doc.InlineShapes)
                {
                    if (!IsPictureInlineShape(inlineShape))
                    {
                        continue;
                    }

                    Word.Range range = inlineShape?.Range;
                    if (range == null || range.End <= scanStart)
                    {
                        continue;
                    }

                    TryAddCaptionEntry(entryMap, GetParagraphBelowRange(doc, range)?.Range, scanStart);
                }
            }
            catch
            {
            }

            try
            {
                foreach (Word.Shape shape in doc.Shapes)
                {
                    if (!IsPictureShape(shape))
                    {
                        continue;
                    }

                    Word.Range anchor = shape?.Anchor;
                    if (anchor == null || anchor.End <= scanStart)
                    {
                        continue;
                    }

                    TryAddCaptionEntry(entryMap, GetParagraphBelowRange(doc, anchor)?.Range, scanStart);
                }
            }
            catch
            {
            }

            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    Word.Range tableRange = table?.Range;
                    if (tableRange == null || tableRange.End <= scanStart)
                    {
                        continue;
                    }

                    TryAddCaptionEntry(entryMap, FindParagraphImmediatelyBeforeTable(table)?.Range, scanStart);
                }
            }
            catch
            {
            }

            // 兜底：如果文档没有被宿主正确暴露为图/表对象，再走一次段落扫描。
            if (entryMap.Count == 0)
            {
                CollectCaptionEntriesByParagraphScan(doc, scanStart, entryMap);
            }

            return entryMap.Values.OrderBy(item => item.Start).ToList();
        }

        private static void CollectCaptionEntriesByParagraphScan(
            Word.Document doc,
            int scanStart,
            Dictionary<int, CaptionListEntry> entryMap)
        {
            if (doc == null || entryMap == null)
            {
                return;
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

            for (int i = 1; i <= paragraphCount; i++)
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

                Word.Range range = paragraph?.Range;
                if (range == null || range.End <= scanStart)
                {
                    continue;
                }

                if (i % 300 == 0)
                {
                    Application.DoEvents();
                }

                TryAddCaptionEntry(entryMap, range, scanStart);
            }
        }

        private static void TryAddCaptionEntry(Dictionary<int, CaptionListEntry> entryMap, Word.Range paragraphRange, int scanStart)
        {
            if (entryMap == null || paragraphRange == null || paragraphRange.End <= scanStart)
            {
                return;
            }

            int start = paragraphRange.Start;
            if (entryMap.ContainsKey(start))
            {
                return;
            }

            string normalized = NormalizeParagraphText(paragraphRange.Text);
            if (!IsCaptionText(normalized))
            {
                return;
            }

            entryMap[start] = new CaptionListEntry
            {
                Start = start,
                Text = normalized
            };
        }

        private static string NormalizeParagraphText(string text)
        {
            return (text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Trim();
        }

        private static bool IsCaptionText(string normalized)
        {
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            return CaptionPrefixRegex.IsMatch(normalized);
        }
    }
}
