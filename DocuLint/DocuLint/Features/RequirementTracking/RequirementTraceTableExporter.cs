using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal sealed class RequirementTraceExportRow
    {
        internal RequirementItem Source { get; set; }

        internal RequirementItem Target { get; set; }

        internal int SourceSpan { get; set; }
    }

    internal static class RequirementTraceTableExporter
    {
        internal static void InsertTraceTable(
            Word.Document document,
            Word.Range range,
            IList<RequirementTraceExportRow> rows,
            string sourceTitle,
            string targetTitle,
            bool includeTargetHeaders = true)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            if (range == null)
            {
                throw new ArgumentNullException(nameof(range));
            }

            IList<RequirementTraceExportRow> safeRows = rows ?? Array.Empty<RequirementTraceExportRow>();
            int rowCount = Math.Max(1, safeRows.Count) + 2;
            Word.Table table = document.Tables.Add(range, rowCount, 6);
            table.Borders.Enable = 1;
            table.Range.Font.Name = "宋体";
            table.Range.Font.Size = 12;
            table.Range.Font.Color = Word.WdColor.wdColorBlack;
            table.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);

            SetCellText(table, 1, 1, sourceTitle);
            SetCellText(table, 1, 4, targetTitle);
            SetCellText(table, 2, 1, "要求名称");
            SetCellText(table, 2, 2, "标识");
            SetCellText(table, 2, 3, "章节号");
            SetCellText(table, 2, 4, includeTargetHeaders ? "需求名称" : string.Empty);
            SetCellText(table, 2, 5, includeTargetHeaders ? "标识" : string.Empty);
            SetCellText(table, 2, 6, includeTargetHeaders ? "章节号" : string.Empty);
            table.Rows[1].Range.Bold = 1;
            table.Rows[2].Range.Bold = 1;
            table.Rows[1].Range.Font.Color = Word.WdColor.wdColorBlack;
            table.Rows[2].Range.Font.Color = Word.WdColor.wdColorBlack;

            for (int i = 0; i < safeRows.Count; i++)
            {
                RequirementTraceExportRow row = safeRows[i];
                int tableRow = i + 3;
                if (row.SourceSpan > 0)
                {
                    SetCellText(table, tableRow, 1, row.Source?.Name);
                    SetCellText(table, tableRow, 2, RequirementItem.GetDisplayRequirementId(row.Source?.Id));
                    SetCellText(table, tableRow, 3, row.Source?.SectionNumber);
                }

                SetCellText(table, tableRow, 4, row.Target?.Name);
                SetCellText(table, tableRow, 5, RequirementItem.GetDisplayRequirementId(row.Target?.Id));
                SetCellText(table, tableRow, 6, row.Target?.SectionNumber);
            }

            MergeCellRange(table, 1, 4, 1, 6);
            MergeCellRange(table, 1, 1, 1, 3);
            MergeSourceCells(table, safeRows);

            table.Range.Font.Color = Word.WdColor.wdColorBlack;
        }

        internal static IList<RequirementTraceExportRow> BuildSourceOnlyRows(IEnumerable<RequirementItem> requirements)
        {
            return (requirements ?? Enumerable.Empty<RequirementItem>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                .Select(item => new RequirementTraceExportRow
                {
                    Source = item,
                    SourceSpan = 1
                })
                .ToList();
        }

        private static void MergeSourceCells(Word.Table table, IList<RequirementTraceExportRow> rows)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                int span = rows[i].SourceSpan;
                if (span <= 1)
                {
                    continue;
                }

                int startRow = i + 3;
                int endRow = startRow + span - 1;
                MergeCellRange(table, startRow, 1, endRow, 1);
                MergeCellRange(table, startRow, 2, endRow, 2);
                MergeCellRange(table, startRow, 3, endRow, 3);
            }
        }

        private static void SetCellText(Word.Table table, int row, int column, string text)
        {
            table.Cell(row, column).Range.Text = text ?? string.Empty;
        }

        private static void MergeCellRange(Word.Table table, int startRow, int startColumn, int endRow, int endColumn)
        {
            try
            {
                table.Cell(startRow, startColumn).Merge(table.Cell(endRow, endColumn));
            }
            catch
            {
            }
        }
    }
}
