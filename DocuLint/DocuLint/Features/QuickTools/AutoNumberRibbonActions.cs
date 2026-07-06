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
        private static readonly Regex SequenceFieldRegex =
            new Regex(@"SEQ\s+([A-Za-z0-9_]+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex FormulaFieldRegex =
            new Regex(@"^\s*=.*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private sealed class ListContinuationContext
        {
            public Word.ListTemplate ListTemplate { get; set; }
            public int ListLevelNumber { get; set; }
        }

        private void button8_Click(object sender, RibbonControlEventArgs e)
        {
            RememberInsertItemAction(button8, button8_Click);
            try
            {
                int filledCellCount = ExecuteAutoNumberInsertion();
                TryUpdateStatusBar(Globals.ThisAddIn.Application, $"DocuLint 已按垂直方向填充 {filledCellCount} 个自动序号");
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show($"自动序号失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private static int ExecuteAutoNumberInsertion()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app?.Selection;
            Word.Document doc = app?.ActiveDocument;
            if (selection == null || doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            Word.Cell currentCell = TryGetCurrentCell(selection);
            if (currentCell == null)
            {
                throw new InvalidOperationException("请先将光标放到表格单元格内。");
            }

            Word.Table table = TryGetOwningTable(currentCell);
            if (table == null)
            {
                throw new InvalidOperationException("未能识别当前所在表格。");
            }

            List<Word.Cell> targetCells = CollectTargetCells(currentCell);
            if (targetCells.Count == 0)
            {
                throw new InvalidOperationException("未找到可填充的目标单元格。");
            }

            List<Word.Cell> occupiedCells = targetCells
                .Where(cell => CellContainsManualContent(cell))
                .ToList();

            if (occupiedCells.Count > 0)
            {
                DialogResult result = MessageBox.Show(
                    "目标单元格中已有内容，继续后会改写这些单元格。是否继续？",
                    "文档不加班 快速工具",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result != DialogResult.Yes)
                {
                    throw new OperationCanceledException("已取消自动序号。");
                }
            }

            Word.Range lastInsertedRange = null;
            bool canContinueFromNeighbor = CanContinueFromNeighbor(currentCell);
            ListContinuationContext listContext = null;

            for (int index = 0; index < targetCells.Count; index++)
            {
                Word.Cell cell = targetCells[index];
                lastInsertedRange = ReplaceCellWithAutoNumberList(
                    cell,
                    isFirstCell: index == 0,
                    canContinueFromNeighbor: canContinueFromNeighbor,
                    listContext: ref listContext);
            }

            try
            {
                doc.Fields.Update();
            }
            catch
            {
            }

            try
            {
                app.ScreenRefresh();
            }
            catch
            {
            }

            Application.DoEvents();

            Word.Range selectionRange = lastInsertedRange ?? currentCell.Range;
            selectionRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            selectionRange.Select();

            return targetCells.Count;
        }

        private static Word.Cell TryGetCurrentCell(Word.Selection selection)
        {
            if (selection == null)
            {
                return null;
            }

            try
            {
                if (selection.Cells != null && selection.Cells.Count > 0)
                {
                    return selection.Cells[1];
                }
            }
            catch
            {
            }

            try
            {
                if (Convert.ToBoolean(selection.get_Information(Word.WdInformation.wdWithInTable)))
                {
                    return selection.Range?.Cells?[1];
                }
            }
            catch
            {
            }

            return null;
        }

        private static List<Word.Cell> CollectTargetCells(Word.Cell currentCell)
        {
            List<Word.Cell> cells = new List<Word.Cell>();
            Word.Table table = TryGetOwningTable(currentCell);
            if (table == null)
            {
                return cells;
            }

            int rowIndex = currentCell.RowIndex;
            int columnIndex = currentCell.ColumnIndex;

            for (int row = rowIndex; row <= table.Rows.Count; row++)
            {
                Word.Cell cell = TryGetTableCell(table, row, columnIndex);
                if (cell != null)
                {
                    cells.Add(cell);
                }
            }

            return cells;
        }

        private static Word.Cell TryGetTableCell(Word.Table table, int rowIndex, int columnIndex)
        {
            if (table == null)
            {
                return null;
            }

            try
            {
                return table.Cell(rowIndex, columnIndex);
            }
            catch
            {
                return null;
            }
        }

        private static bool CellContainsManualContent(Word.Cell cell)
        {
            if (cell == null)
            {
                return false;
            }

            string text = GetCellText(cell);
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return !IsNumericOrAutoNumberCell(cell);
        }

        private static string GetCellText(Word.Cell cell)
        {
            if (cell?.Range == null)
            {
                return string.Empty;
            }

            string text = cell.Range.Text ?? string.Empty;
            return text.Replace("\r", string.Empty)
                       .Replace("\a", string.Empty)
                       .Trim();
        }

        private static bool CanContinueFromNeighbor(Word.Cell currentCell)
        {
            Word.Table table = TryGetOwningTable(currentCell);
            if (table == null)
            {
                return false;
            }

            if (currentCell.RowIndex <= 1)
            {
                return false;
            }

            Word.Cell upperCell = TryGetTableCell(table, currentCell.RowIndex - 1, currentCell.ColumnIndex);
            return IsNumericOrAutoNumberCell(upperCell);
        }

        private static bool IsNumericOrAutoNumberCell(Word.Cell cell)
        {
            if (cell == null)
            {
                return false;
            }

            string text = GetCellText(cell);
            if (int.TryParse(text, out _))
            {
                return true;
            }

            try
            {
                if (cell.Range?.ListFormat != null &&
                    cell.Range.ListFormat.ListType != Word.WdListType.wdListNoNumbering)
                {
                    return true;
                }
            }
            catch
            {
            }

            if (cell?.Range?.Fields == null)
            {
                return false;
            }

            try
            {
                foreach (Word.Field field in cell.Range.Fields)
                {
                    string fieldCode = field?.Code?.Text ?? string.Empty;
                    if (SequenceFieldRegex.IsMatch(fieldCode) || FormulaFieldRegex.IsMatch(fieldCode))
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static Word.Table TryGetOwningTable(Word.Cell cell)
        {
            if (cell?.Range?.Tables == null)
            {
                return null;
            }

            try
            {
                return cell.Range.Tables.Count > 0 ? cell.Range.Tables[1] : null;
            }
            catch
            {
                return null;
            }
        }

        private static Word.Range ReplaceCellWithAutoNumberList(
            Word.Cell cell,
            bool isFirstCell,
            bool canContinueFromNeighbor,
            ref ListContinuationContext listContext)
        {
            if (cell?.Range == null)
            {
                return null;
            }

            Word.Range contentRange = cell.Range.Duplicate;
            if (contentRange.End > contentRange.Start)
            {
                contentRange.End -= 1;
            }

            contentRange.Text = string.Empty;

            contentRange.Text = " ";
            contentRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            Word.Range applyRange = cell.Range.Duplicate;
            if (applyRange.End > applyRange.Start)
            {
                applyRange.End -= 1;
            }

            Word.ListTemplate templateToUse = null;
            int levelToUse = 1;
            bool continuePrevious = false;

            if (isFirstCell && canContinueFromNeighbor)
            {
                Word.Cell upperCell = TryGetTableCell(TryGetOwningTable(cell), cell.RowIndex - 1, cell.ColumnIndex);
                if (upperCell?.Range?.ListFormat?.ListTemplate != null)
                {
                    templateToUse = upperCell.Range.ListFormat.ListTemplate;
                    levelToUse = Math.Max(1, upperCell.Range.ListFormat.ListLevelNumber);
                    continuePrevious = true;
                }
            }

            if (templateToUse == null && listContext?.ListTemplate != null)
            {
                templateToUse = listContext.ListTemplate;
                levelToUse = Math.Max(1, listContext.ListLevelNumber);
                continuePrevious = true;
            }

            if (templateToUse != null)
            {
                applyRange.ListFormat.ApplyListTemplateWithLevel(
                    templateToUse,
                    continuePrevious,
                    Word.WdListApplyTo.wdListApplyToSelection,
                    Word.WdDefaultListBehavior.wdWord10ListBehavior,
                    levelToUse);
            }
            else
            {
                applyRange.ListFormat.ApplyNumberDefault();
            }

            ApplyAutoNumberFormat(applyRange, levelToUse);

            if (applyRange?.ListFormat?.ListTemplate != null)
            {
                if (listContext == null)
                {
                    listContext = new ListContinuationContext();
                }

                listContext.ListTemplate = applyRange.ListFormat.ListTemplate;
                listContext.ListLevelNumber = Math.Max(1, applyRange.ListFormat.ListLevelNumber);
            }

            return applyRange.Duplicate;
        }

        private static void ApplyAutoNumberFormat(Word.Range range, int listLevelNumber)
        {
            if (range?.ListFormat?.ListTemplate == null)
            {
                return;
            }

            int safeLevel = Math.Max(1, listLevelNumber);

            try
            {
                Word.ListLevel level = range.ListFormat.ListTemplate.ListLevels[safeLevel];
                level.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                level.NumberFormat = "%" + safeLevel;
                level.TrailingCharacter = Word.WdTrailingCharacter.wdTrailingNone;
                level.NumberPosition = 0f;
                level.TextPosition = 0f;
                level.TabPosition = 0f;
            }
            catch
            {
            }

            try
            {
                range.ParagraphFormat.LeftIndent = 0f;
                range.ParagraphFormat.FirstLineIndent = 0f;
                range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            catch
            {
            }
        }
    }
}
