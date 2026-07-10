using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private const float DefaultOuterBorderWidthPt = 1.5f;
        private static TablesAndFiguresFormattingSettings currentTablesAndFiguresFormattingSettings =
            TablesAndFiguresFormattingSettings.CreateDefault();

        internal static void ShowTablesAndFiguresFormattingSettingsDialog()
        {
            TablesAndFiguresFormattingSettings dialogDefaults = TablesAndFiguresFormattingSettings.CreateDefault();
            using (TablesAndFiguresFormattingSettingsForm settingsForm =
                new TablesAndFiguresFormattingSettingsForm(dialogDefaults))
            {
                if (settingsForm.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                currentTablesAndFiguresFormattingSettings = settingsForm.Settings?.Clone()
                    ?? TablesAndFiguresFormattingSettings.CreateDefault();
            }
        }

        private static TableFormattingOptions GetCurrentTableFormattingOptions()
        {
            return (currentTablesAndFiguresFormattingSettings?.TableOptions ?? TableFormattingOptions.CreateDefault()).Clone();
        }


        private void button18_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                Word.Selection selection = app?.Selection;

                if (doc == null || selection == null)
                {
                    MessageBox.Show("请先把光标放到要拆分的表格中。", "文档不加班");
                    return;
                }

                Word.Table sourceTable = GetTargetTableFromSelection(selection);
                if (sourceTable?.Range == null)
                {
                    MessageBox.Show("请先把光标放到要拆分的表格中。", "文档不加班");
                    return;
                }

                Word.Paragraph captionParagraph = GetNearestNonEmptyParagraphBeforeTableParagraph(doc, sourceTable);
                string baseCaption = ExtractTableCaptionBase(captionParagraph == null
                    ? string.Empty
                    : NormalizeParagraphText(captionParagraph.Range.Text));
                if (string.IsNullOrWhiteSpace(baseCaption))
                {
                    MessageBox.Show("未识别到当前表格的表题，无法生成“表X（续）”。", "文档不加班");
                    return;
                }

                int sourceTableStart = sourceTable.Range.Start;
                int sourceTableEnd = sourceTable.Range.End;
                int tableCountBefore = doc.Tables?.Count ?? 0;

                using (new WordPerformanceScope(app))
                {
                    if (!TryExecuteNativeSplitTable(selection))
                    {
                        MessageBox.Show("拆分表格失败。请确认光标位于需要成为第二个表格首行的单元格中。", "文档不加班");
                        return;
                    }

                    int tableCountAfter = doc.Tables?.Count ?? 0;
                    if (tableCountAfter <= tableCountBefore)
                    {
                        MessageBox.Show("拆分表格命令未生成第二个表格。请把光标放到需要成为续表首行的单元格中。", "文档不加班");
                        return;
                    }

                    Word.Table continuationTable = FindSplitContinuationTable(doc, sourceTableStart, sourceTableEnd);
                    if (continuationTable?.Range == null)
                    {
                        MessageBox.Show("表格已执行拆分命令，但未找到拆分后的续表。", "文档不加班");
                        return;
                    }

                    Word.Table leadingTable = FindPreviousTableBefore(doc, continuationTable.Range.Start, sourceTableStart - 1)
                        ?? sourceTable;
                    ApplyLeadingBottomOuterBorder(leadingTable);
                    ApplyContinuationTopOuterBorder(continuationTable);

                    if (!InsertSimpleContinuationCaptionBeforeTable(doc, continuationTable, baseCaption + "（续）"))
                    {
                        MessageBox.Show("表格已拆分，但未能在续表前找到可写入题注的表外段落。", "文档不加班");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"拆分表格失败: {ex.Message}", "文档不加班");
            }
        }

        internal static void ExecuteNormalizeSelectedTableAction()
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

                Word.Table targetTable = GetTargetTableFromSelection(selection);
                if (targetTable?.Range == null)
                {
                    MessageBox.Show("请先选中一个表格，或把光标放到需要规范的表格中。", "文档不加班");
                    return;
                }

                TableFormattingOptions options = GetCurrentTableFormattingOptions();
                using (new WordPerformanceScope(app))
                {
                    float tableWidthPoints = ConvertCentimetersToPoints(app, options.TableWidthCentimeters);
                    FormatSingleTable(doc, targetTable, tableWidthPoints, options);
                }

                try
                {
                    app.ScreenRefresh();
                }
                catch
                {
                }

                MessageBox.Show("已应用快速表格样式。", "文档不加班");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"快速表格样式失败: {ex.Message}", "文档不加班");
            }
        }

        private static void ApplyContinuationTopOuterBorder(Word.Table table)
        {
            if (table?.Borders == null)
            {
                return;
            }

            Word.WdLineWidth outerWidth = GetTableBorderLineWidth(table, Word.WdBorderType.wdBorderBottom)
                ?? MapLineWidth(DefaultOuterBorderWidthPt);
            SetTableBorderLine(table.Borders, Word.WdBorderType.wdBorderTop, outerWidth);

            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell != null && GetCellRowIndex(cell) == 1)
                    {
                        SetCellBorderLine(cell, Word.WdBorderType.wdBorderTop, outerWidth);
                    }
                }
            }
            catch
            {
            }
        }

        private static void ApplyLeadingBottomOuterBorder(Word.Table table)
        {
            if (table?.Borders == null)
            {
                return;
            }

            Word.WdLineWidth outerWidth = GetTableBorderLineWidth(table, Word.WdBorderType.wdBorderTop)
                ?? MapLineWidth(DefaultOuterBorderWidthPt);
            SetTableBorderLine(table.Borders, Word.WdBorderType.wdBorderBottom, outerWidth);

            int lastRowIndex = GetTableRowCount(table);
            if (lastRowIndex < 1)
            {
                return;
            }

            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell != null && GetCellRowIndex(cell) == lastRowIndex)
                    {
                        SetCellBorderLine(cell, Word.WdBorderType.wdBorderBottom, outerWidth);
                    }
                }
            }
            catch
            {
            }
        }

        private static Word.WdLineWidth? GetTableBorderLineWidth(Word.Table table, Word.WdBorderType borderType)
        {
            if (table?.Borders == null)
            {
                return null;
            }

            try
            {
                Word.Border border = table.Borders[borderType];
                if (border != null && border.LineStyle != Word.WdLineStyle.wdLineStyleNone)
                {
                    return border.LineWidth;
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool TryExecuteNativeSplitTable(Word.Selection selection)
        {
            if (selection == null)
            {
                return false;
            }

            try
            {
                selection.SplitTable();
                return true;
            }
            catch
            {
                return false;
            }
        }

                                                                private static int GetCellRowIndex(Word.Cell cell)
        {
            if (cell == null)
            {
                return 0;
            }

            try
            {
                return cell.RowIndex;
            }
            catch
            {
                return 0;
            }
        }

        private static int GetCellColumnIndex(Word.Cell cell)
        {
            if (cell == null)
            {
                return 0;
            }

            try
            {
                return cell.ColumnIndex;
            }
            catch
            {
                return 0;
            }
        }

                private static Word.Cell GetFirstCellInRow(Word.Table table, int rowIndex)
        {
            if (table == null || rowIndex < 1)
            {
                return null;
            }

            try
            {
                return table.Cell(rowIndex, 1);
            }
            catch
            {
            }

            try
            {
                Word.Row row = table.Rows[rowIndex];
                if (row?.Cells != null && row.Cells.Count > 0)
                {
                    return row.Cells[1];
                }
            }
            catch
            {
            }

            try
            {
                Word.Cell bestCell = null;
                int bestColumnIndex = int.MaxValue;

                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell == null || GetCellRowIndex(cell) != rowIndex)
                    {
                        continue;
                    }

                    int columnIndex = GetCellColumnIndex(cell);
                    if (bestCell == null || (columnIndex > 0 && columnIndex < bestColumnIndex))
                    {
                        bestCell = cell;
                        bestColumnIndex = columnIndex > 0 ? columnIndex : bestColumnIndex;
                    }
                }

                if (bestCell != null)
                {
                    return bestCell;
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool IsSelectionInsideTable(Word.Selection selection)
        {
            if (selection == null)
            {
                return false;
            }

            try
            {
                return Convert.ToBoolean(selection.get_Information(Word.WdInformation.wdWithInTable));
            }
            catch
            {
                return false;
            }
        }

        private static int GetSelectionRowNumber(Word.Selection selection)
        {
            if (selection == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(selection.get_Information(Word.WdInformation.wdStartOfRangeRowNumber));
            }
            catch
            {
                return 0;
            }
        }

                                private static Word.Table GetTargetTableFromSelection(Word.Selection selection)
        {
            if (selection == null)
            {
                return null;
            }

            try
            {
                if (selection.Tables != null && selection.Tables.Count > 0)
                {
                    return selection.Tables[1];
                }
            }
            catch
            {
            }

            try
            {
                if (Convert.ToBoolean(selection.get_Information(Word.WdInformation.wdWithInTable))
                    && selection.Range?.Tables != null
                    && selection.Range.Tables.Count > 0)
                {
                    return selection.Range.Tables[1];
                }
            }
            catch
            {
            }

            return null;
        }

        private static Word.Table FindTableContainingPosition(Word.Document doc, int position)
        {
            if (doc == null || position < 0)
            {
                return null;
            }

            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    if (table?.Range == null)
                    {
                        continue;
                    }

                    if (table.Range.Start <= position && table.Range.End >= position)
                    {
                        return table;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static Word.Table FindNextTableAfter(Word.Document doc, int position)
        {
            if (doc == null)
            {
                return null;
            }

            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    if (table?.Range == null)
                    {
                        continue;
                    }

                    if (table.Range.Start > position)
                    {
                        return table;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static Word.Table FindSplitContinuationTable(Word.Document doc, int sourceTableStart, int sourceTableEnd)
        {
            if (doc?.Tables == null)
            {
                return null;
            }

            try
            {
                for (int i = 1; i <= doc.Tables.Count; i++)
                {
                    Word.Table table = doc.Tables[i];
                    if (table?.Range == null)
                    {
                        continue;
                    }

                    int start = table.Range.Start;
                    if (start > sourceTableStart && start <= sourceTableEnd + 64)
                    {
                        return table;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static Word.Table FindPreviousTableBefore(Word.Document doc, int position, int minStart)
        {
            if (doc?.Tables == null || position <= 0)
            {
                return null;
            }

            Word.Table previousTable = null;
            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    if (table?.Range == null)
                    {
                        continue;
                    }

                    int start = table.Range.Start;
                    if (start >= minStart && start < position)
                    {
                        previousTable = table;
                    }
                }
            }
            catch
            {
            }

            return previousTable;
        }

                private static int GetTableRowCount(Word.Table table)
        {
            try
            {
                return table?.Rows?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static bool TryGetTableSplitRowIndex(Word.Table table, out int splitRowIndex)
        {
            splitRowIndex = 0;
            if (table?.Range == null)
            {
                return false;
            }

            int firstPage = GetRangeContentStartPageNumber(table.Range);
            if (firstPage <= 0)
            {
                return false;
            }

            try
            {
                for (int i = 2; i <= table.Rows.Count; i++)
                {
                    Word.Row row = table.Rows[i];
                    if (row?.Range == null)
                    {
                        continue;
                    }

                    int rowStartPage = GetRangeContentStartPageNumber(row.Range);
                    int rowEndPage = GetRangeContentEndPageNumber(row.Range);
                    if (rowStartPage > firstPage || rowEndPage > firstPage)
                    {
                        splitRowIndex = i;
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static int GetRangeEndPageNumber(Word.Range range)
        {
            if (range == null)
            {
                return 0;
            }

            try
            {
                Word.Range collapsedRange = range.Duplicate;
                collapsedRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                collapsedRange.MoveStart(Word.WdUnits.wdCharacter, -1);
                return collapsedRange.Information[Word.WdInformation.wdActiveEndPageNumber];
            }
            catch
            {
                return 0;
            }
        }

        private static int GetRangeStartPageNumber(Word.Range range)
        {
            if (range == null)
            {
                return 0;
            }

            try
            {
                Word.Range collapsedRange = range.Duplicate;
                collapsedRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                return collapsedRange.Information[Word.WdInformation.wdActiveEndPageNumber];
            }
            catch
            {
                return 0;
            }
        }

        private static int GetRangeContentEndPageNumber(Word.Range range)
        {
            Word.Range contentRange = GetContentRangeWithoutTrailingMarkers(range);
            return GetRangeEndPageNumber(contentRange ?? range);
        }

        private static int GetRangeContentStartPageNumber(Word.Range range)
        {
            Word.Range contentRange = GetContentRangeWithoutTrailingMarkers(range);
            return GetRangeStartPageNumber(contentRange ?? range);
        }

        private static Word.Range GetContentRangeWithoutTrailingMarkers(Word.Range range)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                Word.Range contentRange = range.Duplicate;
                int removableCharacters = CountTrailingTableMarkers(contentRange.Text);
                if (removableCharacters > 0 && contentRange.End - removableCharacters > contentRange.Start)
                {
                    contentRange.End -= removableCharacters;
                }

                return contentRange;
            }
            catch
            {
                return range;
            }
        }

        private static int CountTrailingTableMarkers(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            int count = 0;
            for (int i = text.Length - 1; i >= 0; i--)
            {
                char c = text[i];
                if (c == '\r' || c == '\a' || c == '\v' || c == '\f')
                {
                    count++;
                    continue;
                }

                break;
            }

            return count;
        }

        private static Word.Paragraph GetNearestNonEmptyParagraphBeforeTableParagraph(Word.Document doc, Word.Table table)
        {
            if (doc == null || table?.Range == null)
            {
                return null;
            }

            int tableStart = table.Range.Start;
            int lookupStart = Math.Max(0, tableStart - 256);

            try
            {
                Word.Range nearbyRange = doc.Range(lookupStart, tableStart);
                Word.Paragraphs paragraphs = nearbyRange?.Paragraphs;
                if (paragraphs == null || paragraphs.Count < 1)
                {
                    return null;
                }

                for (int i = paragraphs.Count; i >= 1; i--)
                {
                    Word.Paragraph paragraph = null;
                    Word.Range paragraphRange = null;

                    try
                    {
                        paragraph = paragraphs[i];
                        paragraphRange = paragraph?.Range;
                    }
                    catch
                    {
                        continue;
                    }

                    if (paragraphRange == null || paragraphRange.End > tableStart)
                    {
                        continue;
                    }

                    string paragraphText = NormalizeParagraphText(paragraphRange.Text);
                    if (!string.IsNullOrWhiteSpace(paragraphText))
                    {
                        return paragraph;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static string ExtractTableCaptionBase(string paragraphText)
        {
            if (string.IsNullOrWhiteSpace(paragraphText))
            {
                return string.Empty;
            }

            Match match = Regex.Match(
                paragraphText,
                @"^(表\s*[0-9０-９一二三四五六七八九十百千]+(?:\s*[\.．\-—]\s*[0-9０-９一二三四五六七八九十百千]+)*)",
                RegexOptions.IgnoreCase);

            return match.Success
                ? (match.Groups[1].Value ?? string.Empty).Trim()
                : string.Empty;
        }

        private static bool InsertSimpleContinuationCaptionBeforeTable(
            Word.Document doc,
            Word.Table continuationTable,
            string continuationCaption)
        {
            if (doc == null
                || continuationTable?.Range == null
                || string.IsNullOrWhiteSpace(continuationCaption))
            {
                return false;
            }

            Word.Range paragraphRange = GetExistingExternalParagraphBeforeTable(continuationTable);
            if (paragraphRange == null)
            {
                return false;
            }

            int paragraphStart = paragraphRange.Start;
            Word.Range textRange = paragraphRange.Duplicate;
            if (textRange.End > textRange.Start)
            {
                textRange.End -= 1;
            }

            textRange.Text = continuationCaption;

            Word.Range formattedRange = doc.Range(
                paragraphStart,
                Math.Min(doc.Content.End, paragraphStart + continuationCaption.Length + 1));
            formattedRange.Font.NameFarEast = "黑体";
            formattedRange.Font.Name = "黑体";
            formattedRange.Font.Size = 12f;
            formattedRange.Font.Bold = 0;
            formattedRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            formattedRange.ParagraphFormat.SpaceBefore = 0f;
            formattedRange.ParagraphFormat.SpaceAfter = 0f;
            formattedRange.ParagraphFormat.KeepWithNext = -1;
            formattedRange.ParagraphFormat.KeepTogether = -1;
            EnsureCaptionOnSamePageAsContinuationTable(doc, formattedRange, continuationTable);
            return true;
        }

        private static Word.Range GetExistingExternalParagraphBeforeTable(Word.Table table)
        {
            if (table?.Range == null)
            {
                return null;
            }

            try
            {
                Word.Range lookupRange = table.Range.Duplicate;
                lookupRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                lookupRange.MoveStart(Word.WdUnits.wdCharacter, -1);

                Word.Paragraphs paragraphs = lookupRange.Paragraphs;
                if (paragraphs == null)
                {
                    return null;
                }

                for (int i = paragraphs.Count; i >= 1; i--)
                {
                    Word.Range paragraphRange = paragraphs[i]?.Range;
                    if (paragraphRange == null || IsRangeInsideTable(paragraphRange))
                    {
                        continue;
                    }

                    if (paragraphRange.End <= table.Range.Start)
                    {
                        return paragraphRange;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool IsRangeInsideTable(Word.Range range)
        {
            if (range == null)
            {
                return false;
            }

            try
            {
                return Convert.ToBoolean(range.Information[Word.WdInformation.wdWithInTable]);
            }
            catch
            {
                return false;
            }
        }

        private static void EnsureCaptionOnSamePageAsContinuationTable(
            Word.Document doc,
            Word.Range captionRange,
            Word.Table continuationTable)
        {
            if (doc == null || captionRange == null || continuationTable?.Range == null)
            {
                return;
            }

            try
            {
                doc.Repaginate();
                int captionPage = GetRangeStartPageNumber(captionRange);
                int tablePage = GetRangeContentStartPageNumber(continuationTable.Range);
                if (captionPage > 0 && tablePage > 0 && captionPage < tablePage)
                {
                    captionRange.ParagraphFormat.PageBreakBefore = -1;
                }
            }
            catch
            {
            }
        }

        private static Word.Paragraph FindParagraphImmediatelyBeforeTable(Word.Table table)
        {
            if (table?.Range == null)
            {
                return null;
            }

            try
            {
                Word.Range lookupRange = table.Range.Duplicate;
                lookupRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                lookupRange.MoveStart(Word.WdUnits.wdCharacter, -1);
                Word.Paragraphs paragraphs = lookupRange.Paragraphs;
                if (paragraphs != null && paragraphs.Count > 0)
                {
                    return paragraphs[paragraphs.Count];
                }
            }
            catch
            {
            }

            return null;
        }

        private static Word.Paragraph FindParagraphBeforeTable(Word.Table table, int offsetFromTable)
        {
            if (table?.Range == null || offsetFromTable < 1)
            {
                return null;
            }

            try
            {
                Word.Range lookupRange = table.Range.Duplicate;
                lookupRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                lookupRange.MoveStart(Word.WdUnits.wdCharacter, -1);
                Word.Paragraphs paragraphs = lookupRange.Paragraphs;
                if (paragraphs == null || paragraphs.Count < offsetFromTable)
                {
                    return null;
                }

                return paragraphs[paragraphs.Count - offsetFromTable + 1];
            }
            catch
            {
                return null;
            }
        }

                private static void EnsureContinuationCaptionStartsOnNextPage(Word.Table firstTable, Word.Range captionRange)
        {
            if (firstTable?.Range == null || captionRange == null)
            {
                return;
            }

            try
            {
                int firstTablePage = GetRangeContentStartPageNumber(firstTable.Range);
                int captionPage = GetRangeStartPageNumber(captionRange);
                if (firstTablePage > 0 && captionPage > 0 && captionPage <= firstTablePage)
                {
                    captionRange.ParagraphFormat.PageBreakBefore = -1;
                }
            }
            catch
            {
            }
        }

        private static bool NeedsPageBreakBeforeContinuationCaption(Word.Table firstTable, Word.Range captionInsertionRange)
        {
            if (firstTable?.Range == null || captionInsertionRange == null)
            {
                return true;
            }

            try
            {
                int firstTablePage = GetRangeContentStartPageNumber(firstTable?.Range);
                int insertionPage = GetRangeStartPageNumber(captionInsertionRange);
                return firstTablePage <= 0 || insertionPage <= 0 || insertionPage <= firstTablePage;
            }
            catch
            {
                return true;
            }
        }

        private static void ApplyCaptionFormatting(Word.Range sourceRange, Word.Range targetRange, string captionStyleName)
        {
            if (targetRange == null)
            {
                return;
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(captionStyleName))
                {
                    TrySetStyle(targetRange, captionStyleName);
                }

                if (sourceRange != null)
                {
                    targetRange.Font = sourceRange.Font;
                    targetRange.ParagraphFormat = sourceRange.ParagraphFormat;
                }
            }
            catch
            {
            }
        }

                private static Word.Cell GetTopLeftCell(Word.Table table)
        {
            if (table?.Range?.Cells == null)
            {
                return null;
            }

            try
            {
                Word.Cell topLeftCell = null;
                int bestRow = int.MaxValue;
                int bestColumn = int.MaxValue;

                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell == null)
                    {
                        continue;
                    }

                    int rowIndex = GetCellRowIndex(cell);
                    int columnIndex = GetCellColumnIndex(cell);
                    if (rowIndex <= 0 || columnIndex <= 0)
                    {
                        continue;
                    }

                    if (rowIndex < bestRow || (rowIndex == bestRow && columnIndex < bestColumn))
                    {
                        bestRow = rowIndex;
                        bestColumn = columnIndex;
                        topLeftCell = cell;
                    }
                }

                return topLeftCell;
            }
            catch
            {
                return null;
            }
        }

        private static int GetCellEndRowIndex(Word.Cell cell)
        {
            if (cell?.Range == null)
            {
                return 0;
            }

            try
            {
                object value = cell.Range.get_Information(Word.WdInformation.wdEndOfRangeRowNumber);
                return value is int rowNumber ? rowNumber : Convert.ToInt32(value);
            }
            catch
            {
                return GetCellRowIndex(cell);
            }
        }

                private static int GetNativeHeaderRowCount(Word.Table table)
        {
            if (table?.Rows == null)
            {
                return 0;
            }

            try
            {
                int nativeHeaderRows = 0;
                foreach (Word.Row row in table.Rows)
                {
                    if (row == null)
                    {
                        continue;
                    }

                    if (row.HeadingFormat != 0)
                    {
                        nativeHeaderRows = row.Index;
                        continue;
                    }

                    break;
                }

                if (nativeHeaderRows > 0)
                {
                    return nativeHeaderRows;
                }
            }
            catch
            {
            }

            return GetLogicalHeaderRowCountFallback(table);
        }

        private static int GetTableRowCountSafe(Word.Table table, int fallback = 0)
        {
            if (table == null)
            {
                return Math.Max(0, fallback);
            }

            try
            {
                if (table.Rows != null)
                {
                    int rowCount = table.Rows.Count;
                    if (rowCount > 0)
                    {
                        return rowCount;
                    }
                }
            }
            catch
            {
            }

            int maxObservedRow = 0;
            try
            {
                if (table.Range?.Cells != null)
                {
                    foreach (Word.Cell cell in table.Range.Cells)
                    {
                        if (cell == null)
                        {
                            continue;
                        }

                        int endRowIndex = GetCellEndRowIndex(cell);
                        if (endRowIndex > maxObservedRow)
                        {
                            maxObservedRow = endRowIndex;
                        }
                    }
                }
            }
            catch
            {
            }

            if (maxObservedRow > 0)
            {
                return maxObservedRow;
            }

            return Math.Max(0, fallback);
        }

        private static int GetNormalizedHeaderRowCount(Word.Table table)
        {
            if (table?.Rows == null || table.Rows.Count < 1)
            {
                return 0;
            }

            int nativeHeaderRowCount = GetNativeHeaderRowCount(table);
            int logicalHeaderRowCount = GetLogicalHeaderRowCountFallback(table);
            int headerRowCount = ResolveHeaderRowCount(nativeHeaderRowCount, logicalHeaderRowCount, table.Rows.Count);

            if (headerRowCount > table.Rows.Count)
            {
                headerRowCount = table.Rows.Count;
            }

            // 自动识别的表头不应吞掉整张多行表格；若出现整表命中，至少保留最后一行作为正文。
            if (table.Rows.Count > 1 && headerRowCount >= table.Rows.Count)
            {
                headerRowCount = table.Rows.Count - 1;
            }

            // 兼容“视觉一行 + 纵向合并”场景：若表头末行为空行，回退到非空行为止。
            while (headerRowCount > 1 && string.IsNullOrWhiteSpace(GetComparableRowSignature(table, headerRowCount)))
            {
                headerRowCount--;
            }

            return headerRowCount;
        }

        private static int GetEffectiveHeaderRowCount(Word.Document doc, Word.Table table)
        {
            return GetNormalizedHeaderRowCount(table);
        }

        private static void FormatSingleTable(
            Word.Document doc,
            Word.Table table,
            float tableWidthPoints,
            TableFormattingOptions options)
        {
            if (doc == null || table?.Range == null || options == null)
            {
                return;
            }

            int headerRowCount = GetEffectiveHeaderRowCount(doc, table);
            ApplyTableWidth(table, tableWidthPoints);
            ApplyTableBorders(table, headerRowCount, options);
            ApplyTableFonts(table, headerRowCount, options);
        }

        private static void ApplyTableWidth(Word.Table table, float tableWidthPoints)
        {
            if (table == null || tableWidthPoints <= 0f)
            {
                return;
            }

            try
            {
                table.AllowAutoFit = false;
            }
            catch
            {
            }

            try
            {
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed);
            }
            catch
            {
            }

            try
            {
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                table.PreferredWidth = tableWidthPoints;
            }
            catch
            {
            }
        }

        private static void ApplyTableBorders(Word.Table table, int headerRowCount, TableFormattingOptions options)
        {
            if (table?.Range == null || options == null)
            {
                return;
            }

            Word.WdLineWidth outerWidth = MapLineWidth(options.OuterBorderWidthPoints);
            Word.WdLineWidth innerWidth = MapLineWidth(options.InnerBorderWidthPoints);

            SetTableBorderLine(table.Borders, Word.WdBorderType.wdBorderHorizontal, innerWidth);
            SetTableBorderLine(table.Borders, Word.WdBorderType.wdBorderVertical, innerWidth);
            SetTableBorderLine(table.Borders, Word.WdBorderType.wdBorderTop, outerWidth);
            SetTableBorderLine(table.Borders, Word.WdBorderType.wdBorderBottom, outerWidth);
            SetTableBorderLine(table.Borders, Word.WdBorderType.wdBorderLeft, outerWidth);
            SetTableBorderLine(table.Borders, Word.WdBorderType.wdBorderRight, outerWidth);
            ClearTableBorderLine(table.Borders, Word.WdBorderType.wdBorderDiagonalDown);
            ClearTableBorderLine(table.Borders, Word.WdBorderType.wdBorderDiagonalUp);

            if (headerRowCount > 0)
            {
                ApplyHeaderOuterBorders(table, headerRowCount, outerWidth);
            }
        }

        private static void ApplyHeaderOuterBorders(Word.Table table, int headerRowCount, Word.WdLineWidth outerWidth)
        {
            if (table?.Range?.Cells == null || headerRowCount < 1)
            {
                return;
            }

            int safeHeaderRowCount = Math.Min(headerRowCount, GetTableRowCountSafe(table, headerRowCount));
            int maxColumnIndex = GetMaxColumnIndex(table);
            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell?.Borders == null)
                    {
                        continue;
                    }

                    int currentRowIndex = GetCellRowIndex(cell);
                    if (currentRowIndex <= 0 || currentRowIndex > safeHeaderRowCount)
                    {
                        continue;
                    }

                    int currentColumnIndex = GetCellColumnIndex(cell);
                    int endRowIndex = GetCellEndRowIndex(cell);
                    int endColumnIndex = GetCellEndColumnIndex(cell);

                    if (currentRowIndex == 1)
                    {
                        SetCellBorderLine(cell, Word.WdBorderType.wdBorderTop, outerWidth);
                    }

                    if (endRowIndex >= safeHeaderRowCount)
                    {
                        SetCellBorderLine(cell, Word.WdBorderType.wdBorderBottom, outerWidth);
                    }

                    if (currentColumnIndex == 1)
                    {
                        SetCellBorderLine(cell, Word.WdBorderType.wdBorderLeft, outerWidth);
                    }

                    if (maxColumnIndex > 0 && endColumnIndex >= maxColumnIndex)
                    {
                        SetCellBorderLine(cell, Word.WdBorderType.wdBorderRight, outerWidth);
                    }
                }
            }
            catch
            {
            }
        }

        private static void ApplyTableFonts(Word.Table table, int headerRowCount, TableFormattingOptions options)
        {
            if (table?.Range?.Cells == null || options == null)
            {
                return;
            }

            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell?.Range == null)
                    {
                        continue;
                    }

                    int rowIndex = GetCellRowIndex(cell);
                    bool isHeaderCell = rowIndex > 0 && rowIndex <= headerRowCount;
                    ApplyFontToRange(
                        cell.Range,
                        isHeaderCell ? options.HeaderFontName : options.BodyFontName,
                        isHeaderCell ? options.HeaderFontSizePoints : options.BodyFontSizePoints);
                }
            }
            catch
            {
            }
        }

        private static void ApplyFontToRange(Word.Range range, string fontName, float fontSize)
        {
            if (range == null)
            {
                return;
            }

            try
            {
                range.Font.NameFarEast = fontName;
            }
            catch
            {
            }

            try
            {
                range.Font.Name = fontName;
            }
            catch
            {
            }

            try
            {
                range.Font.Size = fontSize;
                range.Font.Color = Word.WdColor.wdColorBlack;
            }
            catch
            {
            }
        }

        private static float ConvertCentimetersToPoints(Word.Application app, float centimeters)
        {
            if (app == null)
            {
                return centimeters * 28.3464567f;
            }

            try
            {
                return app.CentimetersToPoints(centimeters);
            }
            catch
            {
                return centimeters * 28.3464567f;
            }
        }

        private static void SetTableBorderLine(Word.Borders borders, Word.WdBorderType borderType, Word.WdLineWidth lineWidth)
        {
            if (borders == null)
            {
                return;
            }

            try
            {
                Word.Border border = borders[borderType];
                if (border == null)
                {
                    return;
                }

                border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                border.LineWidth = lineWidth;
                border.Color = Word.WdColor.wdColorBlack;
            }
            catch
            {
            }
        }

        private static void ClearTableBorderLine(Word.Borders borders, Word.WdBorderType borderType)
        {
            if (borders == null)
            {
                return;
            }

            try
            {
                Word.Border border = borders[borderType];
                if (border == null)
                {
                    return;
                }

                border.LineStyle = Word.WdLineStyle.wdLineStyleNone;
            }
            catch
            {
            }
        }

        private static void SetCellBorderLine(Word.Cell cell, Word.WdBorderType borderType, Word.WdLineWidth lineWidth)
        {
            if (cell?.Borders == null)
            {
                return;
            }

            try
            {
                Word.Border border = cell.Borders[borderType];
                if (border == null)
                {
                    return;
                }

                border.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                border.LineWidth = lineWidth;
                border.Color = Word.WdColor.wdColorBlack;
            }
            catch
            {
            }
        }

        private static int GetMaxColumnIndex(Word.Table table)
        {
            if (table?.Range?.Cells == null)
            {
                return 0;
            }

            int maxColumnIndex = 0;
            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell == null)
                    {
                        continue;
                    }

                    int endColumnIndex = GetCellEndColumnIndex(cell);
                    if (endColumnIndex > maxColumnIndex)
                    {
                        maxColumnIndex = endColumnIndex;
                    }
                }
            }
            catch
            {
            }

            return maxColumnIndex;
        }

        private static int GetCellEndColumnIndex(Word.Cell cell)
        {
            if (cell?.Range == null)
            {
                return 0;
            }

            try
            {
                object value = cell.Range.get_Information(Word.WdInformation.wdEndOfRangeColumnNumber);
                return value is int columnNumber ? columnNumber : Convert.ToInt32(value);
            }
            catch
            {
                return GetCellColumnIndex(cell);
            }
        }

        private static Word.WdLineWidth MapLineWidth(float points)
        {
            KeyValuePair<float, Word.WdLineWidth>[] supportedLineWidths =
            {
                new KeyValuePair<float, Word.WdLineWidth>(0.25f, Word.WdLineWidth.wdLineWidth025pt),
                new KeyValuePair<float, Word.WdLineWidth>(0.5f, Word.WdLineWidth.wdLineWidth050pt),
                new KeyValuePair<float, Word.WdLineWidth>(0.75f, Word.WdLineWidth.wdLineWidth075pt),
                new KeyValuePair<float, Word.WdLineWidth>(1f, Word.WdLineWidth.wdLineWidth100pt),
                new KeyValuePair<float, Word.WdLineWidth>(1.5f, Word.WdLineWidth.wdLineWidth150pt),
                new KeyValuePair<float, Word.WdLineWidth>(2.25f, Word.WdLineWidth.wdLineWidth225pt),
                new KeyValuePair<float, Word.WdLineWidth>(3f, Word.WdLineWidth.wdLineWidth300pt),
            };

            float bestDistance = float.MaxValue;
            Word.WdLineWidth bestWidth = Word.WdLineWidth.wdLineWidth050pt;
            foreach (KeyValuePair<float, Word.WdLineWidth> candidate in supportedLineWidths)
            {
                float distance = Math.Abs(candidate.Key - points);
                if (distance < bestDistance)
                {
                    bestDistance = distance;
                    bestWidth = candidate.Value;
                }
            }

            return bestWidth;
        }

        private static int ResolveHeaderRowCount(int nativeCount, int logicalCount, int totalRows)
        {
            int safeNative = Math.Max(0, nativeCount);
            int safeLogical = Math.Max(0, logicalCount);
            int safeTotalRows = Math.Max(0, totalRows);

            if (safeTotalRows <= 0)
            {
                return 0;
            }

            // 场景约束：复杂表头通常是首行拆分/合并形成的 2~3 行。
            // 这里允许结构推断修正 native=1 的低估，但限制最大修正范围，避免被异常纵向合并放大。
            int inferredCap = Math.Min(safeTotalRows, 4);

            if (safeNative <= 0)
            {
                if (safeLogical <= 0)
                {
                    return 0;
                }

                return Math.Min(safeLogical, inferredCap);
            }

            int result = safeNative;
            if (safeLogical > safeNative && safeLogical <= inferredCap)
            {
                result = safeLogical;
            }

            return Math.Min(result, safeTotalRows);
        }

        private static int GetLogicalHeaderRowCountFallback(Word.Table table)
        {
            if (table?.Range?.Cells == null)
            {
                return 0;
            }

            try
            {
                int headerEndRow = 1;

                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell == null)
                    {
                        continue;
                    }

                    int rowIndex = GetCellRowIndex(cell);
                    if (rowIndex != 1)
                    {
                        continue;
                    }

                    int cellEndRow = GetCellEndRowIndex(cell);
                    if (cellEndRow > headerEndRow)
                    {
                        headerEndRow = cellEndRow;
                    }
                }

                return headerEndRow;
            }
            catch
            {
                return 0;
            }
        }

                        private static void DeleteOneLineBreakBeforeRange(Word.Range range)
        {
            DeleteLineBreaksBeforeRange(range, 1);
        }

        private static int DeleteLineBreaksBeforeRange(Word.Range range, int maxDeleteCount = 1, int minStart = 0)
        {
            if (range?.Document == null || range.Start <= 0 || maxDeleteCount <= 0)
            {
                return 0;
            }

            int deletedCount = 0;
            try
            {
                while (deletedCount < maxDeleteCount && range.Start > minStart)
                {
                    Word.Range previousChar = range.Document.Range(range.Start - 1, range.Start);
                    string text = previousChar.Text ?? string.Empty;
                    if (text.Length == 0)
                    {
                        break;
                    }

                    char ch = text[0];
                    if (!IsLineBreakChar(ch))
                    {
                        break;
                    }

                    previousChar.Delete();
                    deletedCount++;
                }
            }
            catch
            {
            }

            return deletedCount;
        }

        private static bool IsLineBreakChar(char ch)
        {
            return ch == '\r' || ch == '\n' || ch == '\v' || ch == '\f';
        }

                private static string GetComparableRowText(Word.Table table, int rowIndex)
        {
            if (table?.Range?.Cells == null || rowIndex < 1)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();

            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell == null || GetCellRowIndex(cell) != rowIndex)
                    {
                        continue;
                    }

                    string cellText = NormalizeParagraphText(cell.Range?.Text);
                    if (builder.Length > 0)
                    {
                        builder.Append("|");
                    }

                    builder.Append(cellText);
                }
            }
            catch
            {
                return string.Empty;
            }

            return builder.ToString();
        }

                                                                private static string GetComparableRowSignature(Word.Table table, int rowIndex)
        {
            string text = GetComparableRowText(table, rowIndex);
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder(text.Length);
            foreach (char ch in text)
            {
                if (!char.IsWhiteSpace(ch))
                {
                    builder.Append(ch);
                }
            }

            return builder.ToString();
        }

        private static bool TryDeleteSingleRow(Word.Application app, Word.Table table, int rowIndex)
        {
            if (app == null || table == null || rowIndex < 1)
            {
                return false;
            }

            try
            {
                Word.Row directRow = table.Rows[rowIndex];
                directRow.Delete();
                return true;
            }
            catch
            {
            }

            try
            {
                Word.Row selectedRow = table.Rows[rowIndex];
                selectedRow.Select();
                app.Selection.Rows.Delete();
                return true;
            }
            catch
            {
            }

            try
            {
                Word.Row rangeRow = table.Rows[rowIndex];
                if (rangeRow?.Range == null)
                {
                    return false;
                }

                Word.Range rowRange = rangeRow.Range.Duplicate;
                rowRange.Select();
                app.Selection.Rows.Delete();
                return true;
            }
            catch
            {
            }

            try
            {
                Word.Cell firstCell = GetFirstCellInRow(table, rowIndex);
                if (firstCell?.Range != null)
                {
                    firstCell.Range.Select();
                    app.Selection.SelectRow();
                    app.Selection.Rows.Delete();
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool TryDeleteWholeTable(Word.Table table)
        {
            if (table?.Range == null)
            {
                return false;
            }

            try
            {
                table.Delete();
                return true;
            }
            catch
            {
            }

            try
            {
                table.Range.Delete();
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void TrimTableToHeaderRowCount(Word.Table table, int headerRowCount)
        {
            if (table?.Rows == null || headerRowCount < 1)
            {
                return;
            }

            try
            {
                while (table.Rows.Count > headerRowCount)
                {
                    int lastRowIndex = table.Rows.Count;
                    if (!TryDeleteSingleRow(Globals.ThisAddIn?.Application, table, lastRowIndex))
                    {
                        break;
                    }
                }
            }
            catch
            {
            }
        }

        private static bool TryInsertRowsAbove(Word.Selection selection, int rowCount)
        {
            if (selection == null || rowCount < 1)
            {
                return false;
            }

            try
            {
                for (int i = 0; i < rowCount; i++)
                {
                    selection.InsertRowsAbove(1);
                }

                return true;
            }
            catch
            {
            }

            try
            {
                object wordBasic = Globals.ThisAddIn?.Application?.WordBasic;
                if (wordBasic == null)
                {
                    return false;
                }

                for (int i = 0; i < rowCount; i++)
                {
                    wordBasic.GetType().InvokeMember(
                        "TableInsertRow",
                        BindingFlags.InvokeMethod,
                        null,
                        wordBasic,
                        null);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryPasteRowsAtSelection(Word.Selection selection)
        {
            if (selection == null)
            {
                return false;
            }

            try
            {
                selection.PasteAndFormat(Word.WdRecoveryType.wdTableInsertAsRows);
                return true;
            }
            catch
            {
            }

            return TryPasteAtSelection(selection);
        }

        private static bool TryPasteKeepSourceFormatting(Word.Selection selection)
        {
            if (selection == null)
            {
                return false;
            }

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    selection.PasteAndFormat(Word.WdRecoveryType.wdFormatOriginalFormatting);
                    return true;
                }
                catch
                {
                }

                try
                {
                    object placement = Word.WdOLEPlacement.wdInLine;
                    object displayAsIcon = false;
                    object dataType = Word.WdPasteDataType.wdPasteRTF;
                    object link = false;
                    selection.Range.PasteSpecial(ref link, ref dataType, ref placement, ref displayAsIcon);
                    return true;
                }
                catch
                {
                }

                try
                {
                    object wordBasic = Globals.ThisAddIn?.Application?.WordBasic;
                    if (wordBasic != null)
                    {
                        wordBasic.GetType().InvokeMember(
                            "EditPaste",
                            BindingFlags.InvokeMethod,
                            null,
                            wordBasic,
                            null);
                        return true;
                    }
                }
                catch
                {
                }

                System.Threading.Thread.Sleep(20);
            }

            return TryPasteAtSelection(selection);
        }

        private static bool TrySelectCurrentRow(Word.Selection selection)
        {
            if (selection == null)
            {
                return false;
            }

            try
            {
                selection.SelectRow();
                return true;
            }
            catch
            {
            }

            try
            {
                object wordBasic = Globals.ThisAddIn?.Application?.WordBasic;
                if (wordBasic == null)
                {
                    return false;
                }

                wordBasic.GetType().InvokeMember(
                    "SelectRow",
                    BindingFlags.InvokeMethod,
                    null,
                    wordBasic,
                    null);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryExtendSelectionToNextRow(Word.Selection selection, bool selectCurrentRowFirst)
        {
            if (selection == null)
            {
                return false;
            }

            if (selectCurrentRowFirst)
            {
                return TrySelectCurrentRow(selection);
            }

            try
            {
                selection.MoveDown(Word.WdUnits.wdRow, 1, Word.WdMovementType.wdExtend);
                return true;
            }
            catch
            {
            }

            try
            {
                object wordBasic = Globals.ThisAddIn?.Application?.WordBasic;
                if (wordBasic == null)
                {
                    return false;
                }

                object[] args = { 1 };
                wordBasic.GetType().InvokeMember(
                    "NextRow",
                    BindingFlags.InvokeMethod,
                    null,
                    wordBasic,
                    args);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryPasteAtSelection(Word.Selection selection)
        {
            if (selection == null)
            {
                return false;
            }

            try
            {
                selection.Paste();
                return true;
            }
            catch
            {
            }

            try
            {
                object wordBasic = Globals.ThisAddIn?.Application?.WordBasic;
                if (wordBasic == null)
                {
                    return false;
                }

                wordBasic.GetType().InvokeMember(
                    "EditPaste",
                    BindingFlags.InvokeMethod,
                    null,
                    wordBasic,
                    null);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TrySelectionStartsInTable(Word.Selection selection, Word.Table expectedTable)
        {
            if (selection?.Range == null || expectedTable?.Range == null)
            {
                return false;
            }

            try
            {
                int selectionStart = selection.Range.Start;
                return selectionStart >= expectedTable.Range.Start && selectionStart <= expectedTable.Range.End;
            }
            catch
            {
                return false;
            }
        }

        private static void MarkRowAsHeading(Word.Row row)
        {
            if (row == null)
            {
                return;
            }

            try
            {
                row.HeadingFormat = -1;
            }
            catch
            {
            }
        }

        private static void MarkTopRowsAsHeading(Word.Table table, int rowCount)
        {
            if (table?.Rows == null || rowCount < 1)
            {
                return;
            }

            try
            {
                int safeRowCount = Math.Min(rowCount, table.Rows.Count);
                for (int i = 1; i <= safeRowCount; i++)
                {
                    MarkRowAsHeading(table.Rows[i]);
                }
            }
            catch
            {
            }
        }

        private static void MarkFirstRowsAsHeading(Word.Table firstTable, Word.Table continuationTable)
        {
            try
            {
                MarkTopRowsAsHeading(firstTable, 1);
                MarkTopRowsAsHeading(continuationTable, 1);
            }
            catch
            {
            }
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                Word.Selection selection = app?.Selection;
                Word.Range cursorRange = selection?.Range;

                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                if (cursorRange == null)
                {
                    MessageBox.Show("请先把光标放到 Word 文档正文中。", "文档不加班");
                    return;
                }

                SelectNextCrossPageTableAfterCursor(doc, cursorRange);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择下个跨页表格失败: {ex.Message}", "文档不加班");
            }
        }

        private static void SelectNextCrossPageTableAfterCursor(Word.Document doc, Word.Range cursorRange)
        {
            if (doc == null || cursorRange == null)
            {
                MessageBox.Show("请先把光标放到 Word 文档正文中。", "文档不加班");
                return;
            }

            int cursorPosition = cursorRange.End;
            Word.Table nextCrossPageTable = null;

            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    if (table?.Range == null)
                    {
                        continue;
                    }

                    if (table.Range.Start < cursorPosition)
                    {
                        continue;
                    }

                    if (IsCrossPageTable(table))
                    {
                        nextCrossPageTable = table;
                        break;
                    }
                }
            }
            catch
            {
            }

            if (nextCrossPageTable == null)
            {
                MessageBox.Show("光标后面没有找到跨页表格。", "文档不加班");
                return;
            }

            nextCrossPageTable.Select();
        }

        private static bool IsCrossPageTable(Word.Table table)
        {
            if (table?.Range == null)
            {
                return false;
            }

            try
            {
                int startPage = GetRangeContentStartPageNumber(table.Range);
                int endPage = GetRangeContentEndPageNumber(table.Range);
                return startPage != endPage;
            }
            catch
            {
                return false;
            }
        }

    }
}
