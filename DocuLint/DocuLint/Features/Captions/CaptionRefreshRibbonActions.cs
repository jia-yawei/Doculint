using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        internal static void ExecuteRefreshImageCaptions()
        {
            try
            {
                ResetOperationCancellation();
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                int updatedCount = RefreshImageCaptions(doc);
                UpdateCaptionReferenceFields(doc);
                MessageBox.Show(
                    updatedCount > 0
                        ? $"已更新 {updatedCount} 个图片题注。"
                        : "未找到可更新的图片题注。",
                    "文档不加班");
            }
            catch (OperationCanceledException)
            {
                TryUpdateStatusBar(Globals.ThisAddIn?.Application, "图片题注更新已停止");
                MessageBox.Show("图片题注更新已停止。", "文档不加班");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新图片题注失败: {ex.Message}", "文档不加班");
            }
        }

        internal static void ExecuteInsertImageCaption()
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
                    MessageBox.Show("请先把光标放到要插入图片题注的位置。", "文档不加班");
                    return;
                }

                InsertImageCaptionAtSelection(selection);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入图片题注失败: {ex.Message}", "文档不加班");
            }
        }

        internal static void ExecuteRefreshTableCaptions()
        {
            try
            {
                ResetOperationCancellation();
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                int updatedCount = RefreshTableCaptions(doc);
                UpdateCaptionReferenceFields(doc);
                MessageBox.Show(
                    updatedCount > 0
                        ? $"已更新 {updatedCount} 个表格题注。"
                        : "未找到可更新的表格题注。",
                    "文档不加班");
            }
            catch (OperationCanceledException)
            {
                TryUpdateStatusBar(Globals.ThisAddIn?.Application, "表格题注更新已停止");
                MessageBox.Show("表格题注更新已停止。", "文档不加班");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新表格题注失败: {ex.Message}", "文档不加班");
            }
        }

        internal static void ExecuteInsertTableCaption()
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
                    MessageBox.Show("请先把光标放到要插入表格题注的位置。", "文档不加班");
                    return;
                }

                InsertTableCaptionAtSelection(selection);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入表格题注失败: {ex.Message}", "文档不加班");
            }
        }

        private static int RefreshImageCaptions(Word.Document doc)
        {
            return UpdateImageCaptionSequenceFields(doc);
        }

        private static int RefreshTableCaptions(Word.Document doc)
        {
            int updatedCount = UpdateTableCaptionSequenceFields(doc);
            updatedCount += SyncContinuationTableCaptions(doc);
            return updatedCount;
        }

        private static Word.Paragraph GetParagraphBelowRange(Word.Document doc, Word.Range range)
        {
            if (doc == null || range == null)
            {
                return null;
            }

            Word.Paragraph hostParagraph = GetHostParagraph(range);
            if (hostParagraph?.Range == null)
            {
                return null;
            }

            return GetParagraphAfterParagraph(doc, hostParagraph);
        }

        private static Word.Paragraph GetHostParagraph(Word.Range range)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                Word.Paragraphs paragraphs = range.Paragraphs;
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

        private static Word.Paragraph GetParagraphAfterParagraph(Word.Document doc, Word.Paragraph paragraph)
        {
            if (doc == null || paragraph?.Range == null)
            {
                return null;
            }

            try
            {
                int start = Math.Min(doc.Content.End, paragraph.Range.End);
                Word.Range lookupRange = doc.Range(start, doc.Content.End);
                Word.Paragraphs paragraphs = lookupRange.Paragraphs;
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

        private static string NormalizeCaptionParagraphText(string text)
        {
            return (text ?? string.Empty).Replace("\r", string.Empty).Replace("\a", string.Empty).Trim();
        }

        private static void InsertImageCaptionAtSelection(Word.Selection selection)
        {
            if (selection?.Range == null)
            {
                throw new InvalidOperationException("当前光标位置无效。");
            }

            Word.Range insertRange = selection.Range.Duplicate;
            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            insertRange.Text = "图";

            Word.Range fieldInsertRange = insertRange.Duplicate;
            fieldInsertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Field sequenceField = fieldInsertRange.Fields.Add(
                fieldInsertRange,
                Word.WdFieldType.wdFieldEmpty,
                $" SEQ {ImageCaptionSequenceIdentifier} \\* ARABIC ",
                false);
            if (sequenceField == null)
            {
                throw new InvalidOperationException("图片题注编号域插入失败。");
            }

            try
            {
                sequenceField.Update();
            }
            catch
            {
            }

            Word.Range endRange = (sequenceField.Result ?? fieldInsertRange).Duplicate;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            ApplyInsertedCaptionFormatting(selection.Document.Range(insertRange.Start, endRange.End));
            endRange.Select();
        }

        private static void InsertTableCaptionAtSelection(Word.Selection selection)
        {
            if (selection?.Range == null)
            {
                throw new InvalidOperationException("当前光标位置无效。");
            }

            Word.Range insertRange = selection.Range.Duplicate;
            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            insertRange.Text = "表";

            Word.Range fieldInsertRange = insertRange.Duplicate;
            fieldInsertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Field sequenceField = fieldInsertRange.Fields.Add(
                fieldInsertRange,
                Word.WdFieldType.wdFieldEmpty,
                $" SEQ {TableCaptionSequenceIdentifier} \\* ARABIC ",
                false);
            if (sequenceField == null)
            {
                throw new InvalidOperationException("表格题注编号域插入失败。");
            }

            try
            {
                sequenceField.Update();
            }
            catch
            {
            }

            Word.Range endRange = (sequenceField.Result ?? fieldInsertRange).Duplicate;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            ApplyInsertedCaptionFormatting(selection.Document.Range(insertRange.Start, endRange.End));
            endRange.Select();
        }

        private static void ApplyInsertedCaptionFormatting(Word.Range captionRange)
        {
            if (captionRange == null)
            {
                return;
            }

            try
            {
                captionRange.Font.NameFarEast = "黑体";
                captionRange.Font.Name = "黑体";
                captionRange.Font.Size = 12f;
                Word.Range paragraphRange = GetHostParagraph(captionRange)?.Range;
                if (paragraphRange != null)
                {
                    paragraphRange.Font.NameFarEast = "黑体";
                    paragraphRange.Font.Name = "黑体";
                    paragraphRange.Font.Size = 12f;
                }
            }
            catch
            {
            }

            try
            {
                captionRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                Word.Range paragraphRange = GetHostParagraph(captionRange)?.Range;
                if (paragraphRange != null)
                {
                    paragraphRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }
            catch
            {
            }

            try
            {
                Word.Selection selection = Globals.ThisAddIn?.Application?.Selection;
                if (selection != null)
                {
                    selection.Font.NameFarEast = "黑体";
                    selection.Font.Name = "黑体";
                    selection.Font.Size = 12f;
                    selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
            }
            catch
            {
            }
        }

        private static int UpdateImageCaptionSequenceFields(Word.Document doc)
        {
            if (doc == null)
            {
                return 0;
            }

            int updatedCount = 0;
            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                ThrowIfOperationCancelled();
                updatedCount += UpdateImageCaptionSequenceFieldsInRange(storyRange);
            }

            return updatedCount;
        }

        private static int UpdateTableCaptionSequenceFields(Word.Document doc)
        {
            if (doc == null)
            {
                return 0;
            }

            int updatedCount = 0;
            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                ThrowIfOperationCancelled();
                updatedCount += UpdateTableCaptionSequenceFieldsInRange(storyRange);
            }

            return updatedCount;
        }

        private static int SyncContinuationTableCaptions(Word.Document doc)
        {
            if (doc == null)
            {
                return 0;
            }

            int updatedCount = 0;
            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                ThrowIfOperationCancelled();
                if (storyRange?.Paragraphs == null)
                {
                    continue;
                }

                string lastTableCaptionBase = string.Empty;
                try
                {
                    foreach (Word.Paragraph paragraph in storyRange.Paragraphs)
                    {
                        ThrowIfOperationCancelled();
                        Word.Range paragraphRange = paragraph?.Range;
                        string text = NormalizeCaptionParagraphText(paragraphRange?.Text);
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            continue;
                        }

                        if (IsContinuationTableCaptionText(text))
                        {
                            if (!string.IsNullOrWhiteSpace(lastTableCaptionBase))
                            {
                                string expected = lastTableCaptionBase + "（续）";
                                if (!string.Equals(text, expected, StringComparison.Ordinal))
                                {
                                    paragraphRange.Text = expected + "\r";
                                    ApplyContinuationTableCaptionFormatting(paragraphRange);
                                    updatedCount++;
                                }
                            }

                            continue;
                        }

                        string tableCaptionBase = ExtractTableCaptionBaseText(text);
                        if (!string.IsNullOrWhiteSpace(tableCaptionBase))
                        {
                            lastTableCaptionBase = tableCaptionBase;
                        }
                    }
                }
                catch
                {
                }
            }

            return updatedCount;
        }

        private static string ExtractTableCaptionBaseText(string text)
        {
            Match match = TableCaptionBaseRegex.Match(text ?? string.Empty);
            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }

        private static bool IsContinuationTableCaptionText(string text)
        {
            return ContinuationTableCaptionRegex.IsMatch(text ?? string.Empty);
        }

        private static void ApplyContinuationTableCaptionFormatting(Word.Range captionRange)
        {
            if (captionRange == null)
            {
                return;
            }

            try
            {
                captionRange.Font.NameFarEast = "黑体";
                captionRange.Font.Name = "黑体";
                captionRange.Font.Size = 12f;
                captionRange.Font.Bold = 0;
                captionRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                captionRange.ParagraphFormat.SpaceBefore = 0f;
                captionRange.ParagraphFormat.SpaceAfter = 0f;
                captionRange.ParagraphFormat.KeepWithNext = -1;
                captionRange.ParagraphFormat.KeepTogether = -1;
            }
            catch
            {
            }
        }

        private static int UpdateImageCaptionSequenceFieldsInRange(Word.Range range)
        {
            if (range?.Fields == null)
            {
                return 0;
            }

            int updatedCount = 0;
            try
            {
                foreach (Word.Field field in range.Fields)
                {
                    ThrowIfOperationCancelled();
                    if (!IsImageCaptionSequenceField(field))
                    {
                        continue;
                    }

                    try
                    {
                        field.Update();
                        updatedCount++;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return updatedCount;
        }

        private static int UpdateTableCaptionSequenceFieldsInRange(Word.Range range)
        {
            if (range?.Fields == null)
            {
                return 0;
            }

            int updatedCount = 0;
            try
            {
                foreach (Word.Field field in range.Fields)
                {
                    ThrowIfOperationCancelled();
                    if (!IsTableCaptionSequenceField(field))
                    {
                        continue;
                    }

                    try
                    {
                        field.Update();
                        updatedCount++;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return updatedCount;
        }

        private static bool IsImageCaptionSequenceField(Word.Field field)
        {
            if (field == null)
            {
                return false;
            }

            if (field.Type != Word.WdFieldType.wdFieldSequence)
            {
                return false;
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

            if (string.IsNullOrWhiteSpace(codeText))
            {
                return false;
            }

            return Regex.IsMatch(
                codeText,
                $@"\bSEQ\s+{Regex.Escape(ImageCaptionSequenceIdentifier)}\b",
                RegexOptions.IgnoreCase);
        }

        private static bool IsTableCaptionSequenceField(Word.Field field)
        {
            if (field == null)
            {
                return false;
            }

            if (field.Type != Word.WdFieldType.wdFieldSequence)
            {
                return false;
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

            if (string.IsNullOrWhiteSpace(codeText))
            {
                return false;
            }

            return Regex.IsMatch(
                codeText,
                $@"\bSEQ\s+{Regex.Escape(TableCaptionSequenceIdentifier)}\b",
                RegexOptions.IgnoreCase);
        }

        private const string ImageCaptionSequenceIdentifier = "图";
        private const string TableCaptionSequenceIdentifier = "表";
        private static readonly Regex TableCaptionBaseRegex = new Regex(
            @"^\s*(表\s*[0-9０-９一二三四五六七八九十百千]+(?:\s*[\.．\-—]\s*[0-9０-９一二三四五六七八九十百千]+)*)",
            RegexOptions.IgnoreCase);
        private static readonly Regex ContinuationTableCaptionRegex = new Regex(
            @"^\s*表\s*[0-9０-９一二三四五六七八九十百千]+(?:\s*[\.．\-—]\s*[0-9０-９一二三四五六七八九十百千]+)*\s*[（(]\s*续\s*[）)]\s*$",
            RegexOptions.IgnoreCase);
    }
}
