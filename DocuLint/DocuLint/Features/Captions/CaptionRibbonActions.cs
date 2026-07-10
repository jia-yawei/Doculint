using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
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

        private static bool IsPictureInlineShape(Word.InlineShape inlineShape)
        {
            if (inlineShape == null)
            {
                return false;
            }

            try
            {
                switch (inlineShape.Type)
                {
                    case Word.WdInlineShapeType.wdInlineShapePicture:
                    case Word.WdInlineShapeType.wdInlineShapeLinkedPicture:
                        return true;
                    case Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject:
                    case Word.WdInlineShapeType.wdInlineShapeLinkedOLEObject:
                        return IsVisioOleObject(inlineShape.OLEFormat);
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool IsPictureShape(Word.Shape shape)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                return shape.Type == Office.MsoShapeType.msoPicture
                    || shape.Type == Office.MsoShapeType.msoLinkedPicture
                    || ((shape.Type == Office.MsoShapeType.msoEmbeddedOLEObject
                        || shape.Type == Office.MsoShapeType.msoLinkedOLEObject)
                        && IsVisioOleObject(shape.OLEFormat));
            }
            catch
            {
                return false;
            }
        }

        private static bool IsVisioOleObject(Word.OLEFormat oleFormat)
        {
            if (oleFormat == null)
            {
                return false;
            }

            try
            {
                string progId = oleFormat.ProgID ?? string.Empty;
                if (progId.IndexOf("Visio", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                string classType = oleFormat.ClassType ?? string.Empty;
                return classType.IndexOf("Visio", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
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
            CollectCaptionEntriesBySequenceFields(doc, scanStart, entryMap);
            if (entryMap.Count == 0)
            {
                CollectCaptionEntriesFromFields(doc.Fields, scanStart, entryMap);
            }

            return entryMap.Values.OrderBy(item => item.Start).ToList();
        }

        private static void CollectCaptionEntriesBySequenceFields(
            Word.Document doc,
            int scanStart,
            Dictionary<int, CaptionListEntry> entryMap)
        {
            if (doc == null || entryMap == null)
            {
                return;
            }

            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                if (storyRange?.Fields == null)
                {
                    continue;
                }

                int storyScanStart = IsMainTextStory(storyRange) ? scanStart : 0;
                try
                {
                    foreach (Word.Field field in storyRange.Fields)
                    {
                        try
                        {
                            if (!IsCaptionSequenceField(field))
                            {
                                continue;
                            }

                            TryAddCaptionEntry(entryMap, GetCaptionParagraphRange(field), storyScanStart);
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

        private static void CollectCaptionEntriesFromFields(
            Word.Fields fields,
            int scanStart,
            Dictionary<int, CaptionListEntry> entryMap)
        {
            if (fields == null || entryMap == null)
            {
                return;
            }

            try
            {
                foreach (Word.Field field in fields)
                {
                    try
                    {
                        if (!IsCaptionSequenceField(field))
                        {
                            continue;
                        }

                        TryAddCaptionEntry(entryMap, GetCaptionParagraphRange(field), scanStart);
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

        private static Word.Range GetCaptionParagraphRange(Word.Field field)
        {
            try
            {
                return GetHostParagraph(field?.Result)?.Range;
            }
            catch
            {
                return null;
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
            string value = (text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Trim();

            return Regex.Replace(value, @"^[\u0000-\u001F]+", string.Empty).Trim();
        }

        private static bool IsCaptionText(string normalized)
        {
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            return CaptionPrefixRegex.IsMatch(normalized);
        }

        private static bool IsCaptionSequenceField(Word.Field field)
        {
            if (field == null)
            {
                return false;
            }

            string codeText;
            try
            {
                codeText = field.Code?.Text ?? string.Empty;
            }
            catch
            {
                return false;
            }

            Match match = Regex.Match(codeText, @"(?:^|\s)SEQ\s+(?<name>[^\s\\]+)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return false;
            }

            string sequenceName = match.Groups["name"].Value.Trim();
            return string.Equals(sequenceName, ImageCaptionSequenceIdentifier, StringComparison.OrdinalIgnoreCase)
                || string.Equals(sequenceName, TableCaptionSequenceIdentifier, StringComparison.OrdinalIgnoreCase)
                || string.Equals(sequenceName, "Figure", StringComparison.OrdinalIgnoreCase)
                || string.Equals(sequenceName, "Table", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsMainTextStory(Word.Range range)
        {
            try
            {
                return range != null && range.StoryType == Word.WdStoryType.wdMainTextStory;
            }
            catch
            {
                return true;
            }
        }
    }
}
