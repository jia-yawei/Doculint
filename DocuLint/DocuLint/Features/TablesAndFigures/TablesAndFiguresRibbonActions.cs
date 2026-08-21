using System;
using System.Collections.Generic;
using System.Drawing;
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
        private static DocumentCheckSettings currentDocumentCheckSettings =
            DocumentCheckSettings.CreateDefault();

        internal static void ShowTablesAndFiguresFormattingSettingsDialog()
        {
            TablesAndFiguresFormattingSettings dialogDefaults = TablesAndFiguresFormattingSettings.CreateDefault();
            using (TablesAndFiguresFormattingSettingsForm settingsForm =
                new TablesAndFiguresFormattingSettingsForm(
                    dialogDefaults,
                    CommonPhraseLibrary.ConfiguredPath,
                    currentDocumentCheckSettings))
            {
                if (settingsForm.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                currentTablesAndFiguresFormattingSettings = settingsForm.Settings?.Clone()
                    ?? TablesAndFiguresFormattingSettings.CreateDefault();
                currentDocumentCheckSettings = settingsForm.DocumentCheckSettings?.Clone()
                    ?? DocumentCheckSettings.CreateDefault();
                CommonPhraseLibrary.SaveConfiguredPath(settingsForm.CommonPhraseLibraryPath);
                Globals.ThisAddIn?.RefreshCommonPhrasesPane();
            }
        }

        private static TableFormattingOptions GetCurrentTableFormattingOptions()
        {
            return (currentTablesAndFiguresFormattingSettings?.TableOptions ?? TableFormattingOptions.CreateDefault()).Clone();
        }

        internal static DocumentCheckSettings GetCurrentDocumentCheckSettings()
        {
            return (currentDocumentCheckSettings ?? DocumentCheckSettings.CreateDefault()).Clone();
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

                int splitRowIndex = GetSelectionRowNumber(selection);
                int totalRows = GetTableRowCount(sourceTable);
                if (splitRowIndex <= 1 || splitRowIndex > totalRows)
                {
                    MessageBox.Show(
                        "请先把光标放在要成为续表首行的单元格中，再点击“按续表拆分”。",
                        "按续表拆分",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
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
                ContinuationHeaderSelectionForm selector = new ContinuationHeaderSelectionForm(
                    splitRowIndex - 1,
                    mode => CompleteContinuationTableSplit(doc, sourceTableStart, splitRowIndex, baseCaption, mode),
                    () =>
                    {
                        try
                        {
                            app.ActiveWindow?.Activate();
                        }
                        catch
                        {
                        }
                    });
                selector.Show();
                try
                {
                    // Keep Word active so the user can immediately drag-select the header.
                    app.ActiveWindow?.Activate();
                }
                catch
                {
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"拆分表格失败: {ex.Message}", "文档不加班");
            }
        }

        private bool CompleteContinuationTableSplit(
            Word.Document doc,
            int sourceTableStart,
            int splitRowIndex,
            string baseCaption,
            ContinuationHeaderMode headerMode)
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            Word.Table sourceTable = FindTableContainingPosition(doc, sourceTableStart);
            if (doc == null || sourceTable?.Range == null)
            {
                MessageBox.Show("当前表格已不可用，请重新开始按续表拆分。", "按续表拆分", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            int headerRowCount = 1;
            if (headerMode == ContinuationHeaderMode.Custom)
            {
                Word.Selection headerSelection = app?.Selection;
                if (headerSelection?.Range == null ||
                    !TryGetSelectedHeaderRowCount(headerSelection, sourceTable, splitRowIndex, out headerRowCount))
                {
                    MessageBox.Show(
                        "请在原表格中从第一行开始选中完整表头，且表头行不能包含拆分点所在行。",
                        "按续表拆分",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return false;
                }
            }

            int sourceTableEnd = sourceTable.Range.End;
            int tableCountBefore = doc.Tables?.Count ?? 0;
            Word.Cell splitCell = GetFirstCellInRow(sourceTable, splitRowIndex);
            if (splitCell?.Range == null)
            {
                MessageBox.Show("未找到拆分行，请重新开始按续表拆分。", "按续表拆分", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            try
            {
                using (new WordPerformanceScope(app))
                {
                    app.Selection.SetRange(splitCell.Range.Start, splitCell.Range.Start);
                    if (!TryExecuteNativeSplitTable(app.Selection) || (doc.Tables?.Count ?? 0) <= tableCountBefore)
                    {
                        throw new InvalidOperationException("未能在指定行拆分表格。");
                    }

                    Word.Table continuationTable = FindSplitContinuationTable(doc, sourceTableStart, sourceTableEnd);
                    if (continuationTable?.Range == null)
                    {
                        throw new InvalidOperationException("表格已拆分，但未找到拆分后的续表。");
                    }

                    // Keep using the table object captured before SplitTable. Looking it up
                    // again by character position can resolve to the new continuation table
                    // after Word reflows the document ranges.
                    Word.Table leadingTable = sourceTable;
                    if (leadingTable == null || leadingTable.Range == null ||
                        leadingTable.Range.Start >= continuationTable.Range.Start)
                    {
                        leadingTable = FindPreviousTableBefore(doc, continuationTable.Range.Start, sourceTableStart - 1)
                            ?? sourceTable;
                    }
                    if (!CopyHeaderRowsToContinuationTable(doc, leadingTable, continuationTable, headerRowCount))
                    {
                        throw new InvalidOperationException("续表已生成，但复制表头失败。");
                    }

                    CopyTableGeometry(leadingTable, continuationTable);
                    ApplyLeadingBottomOuterBorder(leadingTable);
                    ApplyContinuationTopOuterBorder(continuationTable);
                    if (!InsertSimpleContinuationCaptionBeforeTable(doc, continuationTable, baseCaption + "（续）"))
                    {
                        throw new InvalidOperationException("续表已生成，但未能在续表前写入题注。");
                    }
                }

                app?.ScreenRefresh();
                MessageBox.Show("已完成表格拆分，并为续表补充表头和续表题注。", "按续表拆分", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("拆分表格失败：" + ex.Message, "按续表拆分", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private static bool TryGetSelectedHeaderRowCount(
            Word.Selection selection,
            Word.Table sourceTable,
            int splitRowIndex,
            out int headerRowCount)
        {
            headerRowCount = 0;
            if (selection?.Range == null || sourceTable?.Range == null || splitRowIndex <= 1)
            {
                return false;
            }

            Word.Range selectedRange = selection.Range;
            if (selectedRange.Start < sourceTable.Range.Start || selectedRange.End > sourceTable.Range.End)
            {
                return false;
            }

            int firstRow = int.MaxValue;
            int lastRow = 0;
            try
            {
                foreach (Word.Cell cell in sourceTable.Range.Cells)
                {
                    if (cell?.Range == null || cell.Range.End <= selectedRange.Start || cell.Range.Start >= selectedRange.End)
                    {
                        continue;
                    }

                    int row = GetCellRowIndex(cell);
                    if (row > 0)
                    {
                        firstRow = Math.Min(firstRow, row);
                        lastRow = Math.Max(lastRow, row);
                    }
                }
            }
            catch
            {
                return false;
            }

            if (firstRow != 1 || lastRow <= 0 || lastRow >= splitRowIndex)
            {
                return false;
            }

            headerRowCount = lastRow;
            return true;
        }

        private static bool CopyHeaderRowsToContinuationTable(
            Word.Document doc,
            Word.Table sourceTable,
            Word.Table continuationTable,
            int headerRowCount)
        {
            if (doc == null || sourceTable?.Range == null || continuationTable?.Range == null || headerRowCount < 1)
            {
                return false;
            }

            try
            {
                if (GetTableRowCount(sourceTable) < headerRowCount ||
                    GetTableRowCount(continuationTable) < headerRowCount)
                {
                    return false;
                }

                for (int row = 0; row < headerRowCount; row++)
                {
                    continuationTable.Rows.Add(BeforeRow: continuationTable.Rows[1]);
                }

                // A custom header may contain vertical or horizontal merged cells. Rebuild
                // its merge structure before copying content, rather than assuming every
                // header row has the same number of cells as the continuation data rows.
                if (!MirrorHeaderMergedCells(sourceTable, continuationTable, headerRowCount) ||
                    !CopyHeaderCells(sourceTable, continuationTable, headerRowCount))
                {
                    return false;
                }

                ClearHeaderAutomaticNumbering(continuationTable, headerRowCount);
                int safeHeaderRows = Math.Min(headerRowCount, GetTableRowCount(continuationTable));
                for (int row = 1; row <= safeHeaderRows; row++)
                {
                    try
                    {
                        // Rows.Add can inherit the first data row's automatic numbering
                        // (for example 9). Remove that numbering from the copied header;
                        // the data row below must retain its own sequence number.
                        continuationTable.Rows[row].Range.ListFormat.RemoveNumbers();
                        continuationTable.Rows[row].HeadingFormat = -1;
                    }
                    catch
                    {
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void ClearHeaderAutomaticNumbering(Word.Table table, int headerRowCount)
        {
            if (table?.Range?.Cells == null)
            {
                return;
            }

            HashSet<int> processedCellStarts = new HashSet<int>();
            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell?.Range == null || GetCellRowIndex(cell) > headerRowCount ||
                        !processedCellStarts.Add(cell.Range.Start))
                    {
                        continue;
                    }

                    try
                    {
                        cell.Range.ListFormat.RemoveNumbers();
                    }
                    catch
                    {
                    }

                    try
                    {
                        foreach (Word.Paragraph paragraph in cell.Range.Paragraphs)
                        {
                            paragraph?.Range?.ListFormat?.RemoveNumbers();
                        }
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

        private static void CopyTableGeometry(Word.Table sourceTable, Word.Table continuationTable)
        {
            if (sourceTable == null || continuationTable == null)
            {
                return;
            }

            try
            {
                continuationTable.AllowAutoFit = false;
            }
            catch
            {
            }

            try
            {
                continuationTable.PreferredWidthType = sourceTable.PreferredWidthType;
                continuationTable.PreferredWidth = sourceTable.PreferredWidth;
            }
            catch
            {
            }

            try
            {
                continuationTable.Rows.Alignment = sourceTable.Rows.Alignment;
                continuationTable.Rows.SetLeftIndent(sourceTable.Rows.LeftIndent, Word.WdRulerStyle.wdAdjustNone);
            }
            catch
            {
            }

            try
            {
                int columnCount = Math.Min(sourceTable.Columns.Count, continuationTable.Columns.Count);
                for (int column = 1; column <= columnCount; column++)
                {
                    continuationTable.Columns[column].SetWidth(
                        sourceTable.Columns[column].Width,
                        Word.WdRulerStyle.wdAdjustNone);
                }
            }
            catch
            {
            }

            try
            {
                continuationTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed);
            }
            catch
            {
            }
        }

        private static string RemoveCellEndMarkers(string text)
        {
            return (text ?? string.Empty).TrimEnd('\a', '\r');
        }

        private static bool MirrorHeaderMergedCells(
            Word.Table sourceTable,
            Word.Table destinationTable,
            int headerRowCount)
        {
            try
            {
                int columnCount = GetTableColumnCount(destinationTable);
                if (columnCount < 1)
                {
                    return false;
                }

                List<Word.Cell> headerCells = new List<Word.Cell>();
                Dictionary<int, List<int>> sourceColumnsByRow = new Dictionary<int, List<int>>();
                HashSet<int> processedStarts = new HashSet<int>();
                foreach (Word.Cell sourceCell in sourceTable.Range.Cells)
                {
                    int firstRow = GetCellRowIndex(sourceCell);
                    int firstColumn = GetCellColumnIndex(sourceCell);
                    if (sourceCell?.Range == null || firstRow < 1 || firstRow > headerRowCount || firstColumn < 1 ||
                        !processedStarts.Add(sourceCell.Range.Start))
                    {
                        continue;
                    }

                    headerCells.Add(sourceCell);
                    if (!sourceColumnsByRow.TryGetValue(firstRow, out List<int> columns))
                    {
                        columns = new List<int>();
                        sourceColumnsByRow.Add(firstRow, columns);
                    }

                    columns.Add(firstColumn);
                }

                foreach (List<int> columns in sourceColumnsByRow.Values)
                {
                    columns.Sort();
                }

                foreach (Word.Cell sourceCell in headerCells)
                {
                    int firstRow = GetCellRowIndex(sourceCell);
                    int firstColumn = GetCellColumnIndex(sourceCell);
                    List<int> rowColumns = sourceColumnsByRow[firstRow];
                    int nextColumn = rowColumns.FirstOrDefault(column => column > firstColumn);
                    int lastColumn = nextColumn > 0 ? nextColumn - 1 : columnCount;
                    int lastRow = firstRow;

                    // Word exposes a vertically merged cell only on its first row. A
                    // missing cell origin at the same column in the following header row
                    // therefore means that the source cell continues into that row.
                    while (lastRow < headerRowCount &&
                        (!sourceColumnsByRow.TryGetValue(lastRow + 1, out List<int> nextRowColumns) ||
                         !nextRowColumns.Contains(firstColumn)))
                    {
                        lastRow++;
                    }

                    if (lastRow > firstRow || lastColumn > firstColumn)
                    {
                        Word.Cell firstDestinationCell = TryGetTableCell(destinationTable, firstRow, firstColumn);
                        Word.Cell lastDestinationCell = TryGetTableCell(destinationTable, lastRow, lastColumn);
                        if (firstDestinationCell == null || lastDestinationCell == null)
                        {
                            return false;
                        }

                        firstDestinationCell.Merge(lastDestinationCell);
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool CopyHeaderCells(Word.Table sourceTable, Word.Table destinationTable, int headerRowCount)
        {
            try
            {
                HashSet<int> processedStarts = new HashSet<int>();
                foreach (Word.Cell sourceCell in sourceTable.Range.Cells)
                {
                    int row = GetCellRowIndex(sourceCell);
                    int column = GetCellColumnIndex(sourceCell);
                    if (sourceCell?.Range == null || row < 1 || row > headerRowCount || column < 1 ||
                        !processedStarts.Add(sourceCell.Range.Start))
                    {
                        continue;
                    }

                    Word.Cell destinationCell = TryGetTableCell(destinationTable, row, column);
                    if (destinationCell?.Range == null)
                    {
                        return false;
                    }

                    Word.Range sourceRange = sourceCell.Range.Duplicate;
                    Word.Range destinationRange = destinationCell.Range.Duplicate;
                    sourceRange.End = Math.Max(sourceRange.Start, sourceRange.End - 1);
                    destinationRange.End = Math.Max(destinationRange.Start, destinationRange.End - 1);
                    destinationRange.Text = RemoveCellEndMarkers(sourceRange.Text);
                    destinationRange = destinationCell.Range.Duplicate;
                    destinationRange.End = Math.Max(destinationRange.Start, destinationRange.End - 1);
                    destinationRange.FormattedText = sourceRange.FormattedText;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Word.Cell TryGetTableCell(Word.Table table, int row, int column)
        {
            if (table == null || row < 1 || column < 1)
            {
                return null;
            }

            try
            {
                return table.Cell(row, column);
            }
            catch
            {
                return null;
            }
        }

        private static int GetTableColumnCount(Word.Table table)
        {
            try
            {
                return table?.Rows?[1]?.Cells?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static int GetHeaderEndPosition(Word.Table table, int headerRowCount)
        {
            if (table?.Range == null || headerRowCount < 1)
            {
                return 0;
            }

            try
            {
                return table.Rows[Math.Min(headerRowCount, table.Rows.Count)].Range.End;
            }
            catch
            {
                int end = table.Range.Start;
                try
                {
                    foreach (Word.Cell cell in table.Range.Cells)
                    {
                        if (GetCellRowIndex(cell) <= headerRowCount && cell?.Range != null)
                        {
                            end = Math.Max(end, cell.Range.End);
                        }
                    }
                }
                catch
                {
                }

                return end;
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

        private enum ContinuationHeaderMode
        {
            DefaultFirstRow,
            Custom
        }

        private sealed class ContinuationHeaderSelectionForm : Form
        {
            private readonly Func<ContinuationHeaderMode, bool> completeAction;
            private readonly Action activateWordAction;
            private readonly Label hint;
            private readonly Label status;
            private readonly RadioButton customHeaderRadioButton;

            internal ContinuationHeaderSelectionForm(
                int maximumHeaderRows,
                Func<ContinuationHeaderMode, bool> completeAction,
                Action activateWordAction)
            {
                this.completeAction = completeAction;
                this.activateWordAction = activateWordAction;
                Text = "指定续表表头";
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                StartPosition = FormStartPosition.CenterScreen;
                ShowInTaskbar = false;
                TopMost = true;
                ControlBox = true;
                AutoScaleMode = AutoScaleMode.Dpi;
                Font = new Font("Microsoft YaHei UI", 9F);
                ClientSize = new Size(680, 340);
                MinimumSize = new Size(620, 320);
                MinimizeBox = false;
                MaximizeBox = false;

                TableLayoutPanel layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = 5,
                    Padding = new Padding(16)
                };
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

                Label title = new Label
                {
                    AutoSize = true,
                    Text = "选择续表表头方式",
                    Font = new Font(Font, FontStyle.Bold),
                    Margin = new Padding(0, 0, 0, 8)
                };
                FlowLayoutPanel modes = new FlowLayoutPanel
                {
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.TopDown,
                    WrapContents = false,
                    Margin = new Padding(0, 0, 0, 6)
                };
                RadioButton defaultHeaderRadioButton = new RadioButton
                {
                    AutoSize = true,
                    Text = "默认表头（当前表格的第一行）",
                    Checked = true,
                    Margin = new Padding(0, 0, 0, 5)
                };
                customHeaderRadioButton = new RadioButton
                {
                    AutoSize = true,
                    Text = "自定义表头（在 Word 中手动选中表头）",
                    Margin = new Padding(0)
                };
                modes.Controls.Add(defaultHeaderRadioButton);
                modes.Controls.Add(customHeaderRadioButton);

                hint = new Label
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    TextAlign = ContentAlignment.MiddleLeft,
                    Padding = new Padding(0, 3, 0, 3),
                    ForeColor = Color.FromArgb(55, 65, 80)
                };
                status = new Label
                {
                    AutoSize = true,
                    ForeColor = Color.FromArgb(85, 95, 110),
                    Margin = new Padding(0, 4, 0, 10)
                };
                FlowLayoutPanel buttons = new FlowLayoutPanel
                {
                    AutoSize = true,
                    FlowDirection = FlowDirection.RightToLeft,
                    Dock = DockStyle.Fill,
                    WrapContents = false,
                    Padding = new Padding(0, 2, 0, 0),
                    Margin = new Padding(0),
                    MinimumSize = new Size(0, 42)
                };
                Button completeButton = new Button
                {
                    Text = "完成并拆分",
                    AutoSize = true,
                    Height = 34,
                    MinimumSize = new Size(104, 32),
                    Margin = new Padding(8, 0, 0, 0)
                };
                Button cancelButton = new Button
                {
                    Text = "取消",
                    AutoSize = true,
                    Height = 34,
                    MinimumSize = new Size(78, 32)
                };
                completeButton.Click += (_, __) =>
                {
                    ContinuationHeaderMode mode = customHeaderRadioButton.Checked
                        ? ContinuationHeaderMode.Custom
                        : ContinuationHeaderMode.DefaultFirstRow;
                    if (completeAction?.Invoke(mode) == true)
                    {
                        Close();
                    }
                };
                cancelButton.Click += (_, __) => Close();
                defaultHeaderRadioButton.CheckedChanged += (_, __) => UpdateModeHint(maximumHeaderRows);
                customHeaderRadioButton.CheckedChanged += (_, __) => UpdateModeHint(maximumHeaderRows);
                buttons.Controls.Add(completeButton);
                buttons.Controls.Add(cancelButton);
                layout.Controls.Add(title, 0, 0);
                layout.Controls.Add(modes, 0, 1);
                layout.Controls.Add(hint, 0, 2);
                layout.Controls.Add(status, 0, 3);
                layout.Controls.Add(buttons, 0, 4);
                Controls.Add(layout);
                UpdateModeHint(maximumHeaderRows);
            }

            private void UpdateModeHint(int maximumHeaderRows)
            {
                if (customHeaderRadioButton.Checked)
                {
                    hint.Text = "请回到 Word，在原表格中从第一行开始连续选中表头。\r\n" +
                        "表头最多可选择 " + maximumHeaderRows + " 行，不能包含续表的首行。";
                    status.Text = "选中后点击“完成并拆分”，插件会将所选表头复制到续表。";
                    BeginInvoke(new Action(() => activateWordAction?.Invoke()));
                    return;
                }

                hint.Text = "将使用当前表格的第一行作为续表表头，无需额外选择。";
                status.Text = "点击“完成并拆分”后，插件会自动拆分表格、复制表头并补充续表题注。";
            }

            // Do not activate this tool window when it is shown. Word remains the active
            // window, so drag-selection of the table header is uninterrupted.
            protected override bool ShowWithoutActivation => true;

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams parameters = base.CreateParams;
                    parameters.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
                    return parameters;
                }
            }
        }

    }
}
