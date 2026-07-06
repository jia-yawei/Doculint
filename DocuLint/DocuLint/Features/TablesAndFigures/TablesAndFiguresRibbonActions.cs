using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private const float DefaultTableWidthCm = 17.4f;
        private const float DefaultOuterBorderWidthPt = 1.5f;
        private const float DefaultInnerBorderWidthPt = 0.5f;
        private const float DefaultHeaderFontSizePt = 12f;
        private const float DefaultBodyFontSizePt = 10.5f;
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

        private static ImageFormattingOptions GetCurrentImageFormattingOptions()
        {
            return (currentTablesAndFiguresFormattingSettings?.ImageOptions ?? ImageFormattingOptions.CreateDefault()).Clone();
        }

        private void button18_Click(object sender, RibbonControlEventArgs e)
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

                if (selection == null)
                {
                    MessageBox.Show("请先选中跨页表格，或把光标放到跨页表格中。", "文档不加班");
                    return;
                }

                using (TableSplitProgressForm progressForm = new TableSplitProgressForm())
                {
                    progressForm.Show();
                    progressForm.ReportProgress(3, "正在准备拆分...", "正在读取当前选中的跨页表格。");

                    try
                    {
                        int continuationCount;
                        using (new WordPerformanceScope(app))
                        {
                            continuationCount = SplitTableAtAutoCrossPagePosition(app, doc, selection, progressForm);
                        }

                        if (continuationCount > 0 && !progressForm.IsFinalized)
                        {
                            try
                            {
                                app.ScreenRefresh();
                            }
                            catch
                            {
                            }

                            Application.DoEvents();
                            progressForm.Complete(
                                $"已成功拆分为{continuationCount}个续表。",
                                "表格拆分、表头补齐和续表题注插入已完成。",
                                true);
                        }
                        else if (!progressForm.IsFinalized)
                        {
                            progressForm.Complete(
                                "当前表格没有生成新的续表。",
                                "如果这是跨页表格，请确认光标位于需要处理的表格内，或先选中目标表格。",
                                false);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!progressForm.IsFinalized)
                        {
                            progressForm.Complete(
                                "按续表拆分失败。",
                                ex.Message,
                                false);
                        }
                    }

                    progressForm.WaitForUserClose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"按续表拆分失败: {ex.Message}", "文档不加班");
            }
        }

        private static void SplitTableAtCurrentCursor(Word.Application app, Word.Document doc, Word.Selection selection)
        {
            if (!TrySplitTableAtSelection(app, doc, selection, out string errorMessage))
            {
                MessageBox.Show(errorMessage, "文档不加班");
                return;
            }

            MessageBox.Show("表格已按当前光标位置拆分。", "文档不加班");
        }

        internal static void ExecuteNormalizeTablesAction()
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                TableFormattingOptions options = GetCurrentTableFormattingOptions();

                using (TableFormattingProgressForm progressForm = new TableFormattingProgressForm())
                {
                    progressForm.Show();
                    progressForm.ReportProgress(3, "正在准备规范表格...", "正在读取目录后的表格范围和当前参数。");

                    try
                    {
                        int formattedCount;
                        using (new WordPerformanceScope(app))
                        {
                            formattedCount = NormalizeTables(
                                app,
                                doc,
                                CollectTablesAfterToc(doc),
                                options,
                                progressForm,
                                "目录后的表格");
                        }

                        if (formattedCount > 0 && !progressForm.IsFinalized)
                        {
                            try
                            {
                                app.ScreenRefresh();
                            }
                            catch
                            {
                            }

                            Application.DoEvents();
                            progressForm.Complete(
                                $"已完成 {formattedCount} 个表格的规范处理。",
                                "表头、正文字体、表格宽度和边框已按当前参数统一设置。",
                                true);
                        }
                        else if (!progressForm.IsFinalized)
                        {
                            progressForm.Complete(
                                "未找到可规范的表格。",
                                "如果文档包含目录，则只会处理目录之后的表格；如果没有目录，则处理全部表格。",
                                false);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!progressForm.IsFinalized)
                        {
                            progressForm.Complete(
                                "一键规范表格失败。",
                                ex.Message,
                                false);
                        }
                    }

                    progressForm.WaitForUserClose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"一键规范表格失败: {ex.Message}", "文档不加班");
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

                using (TableFormattingProgressForm progressForm = new TableFormattingProgressForm())
                {
                    progressForm.Show();
                    progressForm.ReportProgress(5, "正在准备规范当前表格...", "正在读取当前选中的表格和规范参数。");

                    try
                    {
                        int formattedCount;
                        using (new WordPerformanceScope(app))
                        {
                            formattedCount = NormalizeTables(
                                app,
                                doc,
                                new List<Word.Table> { targetTable },
                                options,
                                progressForm,
                                "当前表格");
                        }

                        if (formattedCount > 0 && !progressForm.IsFinalized)
                        {
                            try
                            {
                                app.ScreenRefresh();
                            }
                            catch
                            {
                            }

                            Application.DoEvents();
                            progressForm.Complete(
                                "已完成当前表格的规范处理。",
                                "表头、正文字体、表格宽度和边框已按当前参数统一设置。",
                                true);
                        }
                        else if (!progressForm.IsFinalized)
                        {
                            progressForm.Complete(
                                "未找到可规范的当前表格。",
                                "请确认光标位于表格中，或已选中目标表格。",
                                false);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!progressForm.IsFinalized)
                        {
                            progressForm.Complete(
                                "规范当前表格失败。",
                                ex.Message,
                                false);
                        }
                    }

                    progressForm.WaitForUserClose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"规范当前表格失败: {ex.Message}", "文档不加班");
            }
        }

        internal static void ExecuteNormalizeAllImagesAction()
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                ImageFormattingOptions options = GetCurrentImageFormattingOptions();

                using (ImageFormattingProgressForm progressForm = new ImageFormattingProgressForm())
                {
                    progressForm.Show();
                    progressForm.ReportProgress(3, "正在准备规范图片...", "正在读取主文档中的全部图片和规范参数。");

                    try
                    {
                        int formattedCount;
                        using (new WordPerformanceScope(app))
                        {
                            formattedCount = NormalizeAllImages(doc, options, progressForm);
                        }

                        if (formattedCount > 0 && !progressForm.IsFinalized)
                        {
                            try
                            {
                                app.ScreenRefresh();
                            }
                            catch
                            {
                            }

                            Application.DoEvents();
                            progressForm.Complete(
                                $"已完成 {formattedCount} 张图片的规范处理。",
                                "图片缩放比例和所属段落格式已按当前参数统一设置。",
                                true);
                        }
                        else if (!progressForm.IsFinalized)
                        {
                            progressForm.Complete(
                                "未找到可规范的图片。",
                                "当前只处理主文档正文中的图片对象。",
                                false);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (!progressForm.IsFinalized)
                        {
                            progressForm.Complete(
                                "规范全部图片失败。",
                                ex.Message,
                                false);
                        }
                    }

                    progressForm.WaitForUserClose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"规范全部图片失败: {ex.Message}", "文档不加班");
            }
        }

        internal static void ExecuteMergeContinuationTableAction()
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

                using (new WordPerformanceScope(app))
                {
                    int mergedCount = MergeSelectedContinuationTable(app, doc, selection);
                    if (mergedCount > 0)
                    {
                        try
                        {
                            app.ScreenRefresh();
                        }
                        catch
                        {
                        }

                        MessageBox.Show("已合并当前续表。", "文档不加班");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"合并续表失败: {ex.Message}", "文档不加班");
            }
        }

        private static int MergeSelectedContinuationTable(
            Word.Application app,
            Word.Document doc,
            Word.Selection selection)
        {
            if (app == null || doc == null || selection == null)
            {
                MessageBox.Show("请先把光标放到要合并的续表中，或选中续表。", "文档不加班");
                return 0;
            }

            Word.Table selectedTable = GetTargetTableFromSelection(selection);
            if (selectedTable?.Range == null)
            {
                MessageBox.Show("请先把光标放到要合并的续表中，或选中续表。", "文档不加班");
                return 0;
            }

            if (!TryResolveContinuationMergeTables(doc, selectedTable, out Word.Table previousTable, out Word.Table continuationTable))
            {
                MessageBox.Show("未找到可合并的续表。请把光标放到续表中，或把光标放到正表中再执行合并。", "文档不加班");
                return 0;
            }

            Word.Paragraph continuationCaptionParagraph = GetNearestNonEmptyParagraphBeforeTableParagraph(doc, continuationTable);
            string continuationCaptionText = continuationCaptionParagraph == null
                ? string.Empty
                : NormalizeParagraphText(continuationCaptionParagraph.Range.Text);
            if (!IsContinuationTableCaption(continuationCaptionText))
            {
                MessageBox.Show("当前表格不是续表，无法执行合并续表。请先选中续表，或把光标放到续表中。", "文档不加班");
                return 0;
            }

            int rowCount = GetTableRowCount(continuationTable);
            if (rowCount <= 0)
            {
                MessageBox.Show("当前续表为空，无法合并。", "文档不加班");
                return 0;
            }

            if (ShouldDeleteContinuationHeaderBeforeMerge(doc, previousTable, continuationTable))
            {
                if (rowCount <= 1)
                {
                    MessageBox.Show("当前续表没有可保留的正文行，无法合并。", "文档不加班");
                    return 0;
                }

                int headerRowCount = GetMergeContinuationHeaderRowCount(doc, previousTable, continuationTable);
                if (headerRowCount <= 0)
                {
                    headerRowCount = 1;
                }

                headerRowCount = Math.Min(headerRowCount, rowCount - 1);
                if (!TryDeleteRowBlock(app, continuationTable, 1, headerRowCount))
                {
                    MessageBox.Show("删除续表表头失败，未执行合并。", "文档不加班");
                    return 0;
                }
            }

            List<TableHeaderMemoryItem> trailingTableMemories = CaptureTrailingTableHeaderMemories(doc, continuationTable.Range.Start);
            DeleteParagraphRange(continuationCaptionParagraph);
            DeleteParagraphBetweenTables(previousTable, continuationTable);
            RestoreTrailingTableHeaderMemories(doc, trailingTableMemories);
            Word.Table mergedTable = FindTableContainingPosition(doc, previousTable.Range.Start) ?? previousTable;
            try
            {
                mergedTable.Range.Select();
            }
            catch
            {
                try
                {
                    previousTable.Range.Select();
                }
                catch
                {
                }
            }

            return 1;
        }

        private static bool TryResolveContinuationMergeTables(
            Word.Document doc,
            Word.Table selectedTable,
            out Word.Table previousTable,
            out Word.Table continuationTable)
        {
            previousTable = null;
            continuationTable = null;

            if (doc == null || selectedTable?.Range == null)
            {
                return false;
            }

            Word.Paragraph selectedCaptionParagraph = GetNearestNonEmptyParagraphBeforeTableParagraph(doc, selectedTable);
            string selectedCaptionText = selectedCaptionParagraph == null
                ? string.Empty
                : NormalizeParagraphText(selectedCaptionParagraph.Range.Text);

            if (IsContinuationTableCaption(selectedCaptionText))
            {
                continuationTable = selectedTable;
                previousTable = FindPreviousTableBefore(doc, continuationTable.Range.Start);
                return previousTable?.Range != null;
            }

            previousTable = selectedTable;
            continuationTable = FindNextContinuationTable(doc, selectedTable.Range.End);
            return continuationTable?.Range != null;
        }

        private static bool TrySplitTableAtSelection(
            Word.Application app,
            Word.Document doc,
            Word.Selection selection,
            out string errorMessage)
        {
            errorMessage = string.Empty;

            if (app == null || doc == null || selection == null)
            {
                errorMessage = "请先把光标放到要拆分的表格中。";
                return false;
            }

            if (!IsSelectionInsideTable(selection))
            {
                errorMessage = "请先把光标放到要拆分的表格单元格中。";
                return false;
            }

            int rowNumber = GetSelectionRowNumber(selection);
            if (rowNumber <= 1)
            {
                errorMessage = "当前光标位于表格第一行，无法按当前位置拆成两个表格。";
                return false;
            }

            int tableCountBefore = GetTableCount(doc);
            int selectionStart = selection.Range?.Start ?? 0;

            try
            {
                app.Selection.SplitTable();
            }
            catch
            {
                if (!TryExecuteNativeSplitTableThroughWordBasic(app))
                {
                    throw;
                }
            }

            int tableCountAfter = GetTableCount(doc);
            if (tableCountAfter <= tableCountBefore)
            {
                TryExecuteNativeSplitTableThroughWordBasic(app);
                tableCountAfter = GetTableCount(doc);
            }

            if (tableCountAfter <= tableCountBefore)
            {
                errorMessage = $"拆分表格命令已执行，但文档中的表格数量没有增加。请确认光标在需要成为第二个表格首行的单元格中。光标位置：{selectionStart}";
                return false;
            }

            return true;
        }

        private static int SplitTableAtAutoCrossPagePosition(
            Word.Application app,
            Word.Document doc,
            Word.Selection selection,
            TableSplitProgressForm progressForm = null)
        {
            if (app == null || doc == null || selection == null)
            {
                ReportTableSplitIssue(progressForm, "请先把光标放到要拆分的表格中。");
                return 0;
            }

            Word.Table targetTable = GetTargetTableFromSelection(selection);
            if (targetTable?.Range == null)
            {
                ReportTableSplitIssue(progressForm, "请先选中跨页表格，或把光标放到跨页表格中。");
                return 0;
            }

            progressForm?.ReportProgress(8, "正在检查表格...", "正在确认表格行数和跨页拆分条件。");

            try
            {
                if (targetTable.Rows != null && targetTable.Rows.Count <= 5)
                {
                    DialogResult result = MessageBox.Show(
                        "当前选中表格的总行数不大于5行，是否继续进行按续表拆分？",
                        "文档不加班",
                        MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Question);
                    
                    if (result != DialogResult.OK)
                    {
                        return 0;
                    }
                }
            }
            catch
            {
            }

            Word.Paragraph captionParagraph = GetNearestNonEmptyParagraphBeforeTableParagraph(doc, targetTable);
            string originalCaptionText = captionParagraph == null
                ? string.Empty
                : NormalizeParagraphText(captionParagraph.Range.Text);
            progressForm?.ReportProgress(12, "正在识别表题...", "正在读取原表题，用于生成“表X（续）”。");
            string baseCaption = ExtractTableCaptionBase(originalCaptionText);
            if (string.IsNullOrWhiteSpace(baseCaption))
            {
                ReportTableSplitIssue(progressForm, "未识别到当前表格的表题，无法生成“表X（续）”。");
                return 0;
            }

            int workingTableStart = targetTable.Range.Start;
            string continuationCaption = baseCaption + "（续）";
            string captionStyleName = captionParagraph == null
                ? string.Empty
                : ResolveStyleName(TryGetParagraphStyle(captionParagraph.Range), doc);
            Word.Range captionStyleSourceRange = captionParagraph?.Range?.Duplicate;
            int continuationCount = 0;
            const int maxSplitRounds = 64;

            for (int round = 0; round < maxSplitRounds; round++)
            {
                int basePercent = Math.Min(90, 16 + (round * 10));
                progressForm?.ReportProgress(
                    basePercent,
                    $"正在分析第{round + 1}处跨页位置...",
                    $"已生成 {continuationCount} 个续表，正在判断当前表格的分页位置。");

                Word.Table currentTable = FindTableContainingPosition(doc, workingTableStart);
                if (currentTable?.Range == null)
                {
                    if (continuationCount == 0)
                    {
                        ReportTableSplitIssue(progressForm, "未找到当前需要拆分的表格。");
                    }
                    break;
                }

                if (!TryGetCrossPageSplitCell(currentTable, out Word.Cell splitCell, out string reason))
                {
                    if (continuationCount == 0)
                    {
                        ReportTableSplitIssue(progressForm, GetCrossPageTableRequiredMessage(reason));
                    }
                    break;
                }

                progressForm?.ReportProgress(
                    Math.Min(92, basePercent + 3),
                    $"正在定位第{round + 1}处拆分位置...",
                    "已找到跨页位置，正在把光标移动到续表首行。");

                try
                {
                    Word.Range splitRange = splitCell.Range;
                    splitRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    splitRange.Select();
                }
                catch (Exception ex)
                {
                    ReportTableSplitIssue(progressForm, $"定位拆分位置失败: {ex.Message}");
                    return continuationCount;
                }

                progressForm?.ReportProgress(
                    Math.Min(94, basePercent + 5),
                    $"正在拆分第{round + 1}个续表...",
                    "正在执行表格拆分命令。");

                if (!TrySplitTableAtSelection(app, doc, app.Selection, out string splitErrorMessage))
                {
                    ReportTableSplitIssue(progressForm, splitErrorMessage);
                    return continuationCount;
                }

                Word.Table upperTable = FindTableContainingPosition(doc, workingTableStart);
                Word.Table continuationTable = FindNextTableAfter(doc, upperTable?.Range?.End ?? workingTableStart);
                if (upperTable?.Range == null || continuationTable?.Range == null)
                {
                    ReportTableSplitIssue(progressForm, "表格已拆分，但未找到续表部分。");
                    return continuationCount;
                }

                progressForm?.ReportProgress(
                    Math.Min(96, basePercent + 7),
                    $"正在补齐第{round + 1}个续表表头...",
                    "正在复制原表表头到续表。");

                int effectiveHeaderCount = GetEffectiveHeaderRowCount(doc, currentTable);
                bool isFirstRound = (round == 0);
                if (!CopyHeaderRowToContinuationTable(app, upperTable, continuationTable, effectiveHeaderCount, isFirstRound, out string headerCopyMessage))
                {
                    ReportTableSplitIssue(progressForm, headerCopyMessage);
                    return continuationCount;
                }

                continuationTable = FindNextTableAfter(doc, upperTable.Range.End);
                if (continuationTable?.Range == null)
                {
                    ReportTableSplitIssue(progressForm, "表格已拆分并补齐表头，但未重新找到续表。");
                    return continuationCount;
                }

                progressForm?.ReportProgress(
                    Math.Min(98, basePercent + 9),
                    $"正在插入第{round + 1}个续表题注...",
                    "正在写入续表题注并调整分页位置。");

                Word.Range continuationCaptionRange = InsertContinuationCaptionAfterSplit(
                    app,
                    continuationTable,
                    continuationCaption,
                    captionStyleName,
                    captionStyleSourceRange);
                EnsureContinuationCaptionStartsOnNextPage(upperTable, continuationCaptionRange);
                RememberHeaderRowCount(doc, upperTable, effectiveHeaderCount);
                RememberHeaderRowCount(doc, continuationTable, effectiveHeaderCount);

                continuationCount++;
                workingTableStart = continuationTable.Range.Start;
            }

            return continuationCount;
        }

        private static void ReportTableSplitIssue(TableSplitProgressForm progressForm, string message)
        {
            if (progressForm != null)
            {
                progressForm.Complete(message, "请根据提示调整表格或光标位置后重试。", false);
                return;
            }

            MessageBox.Show(message, "文档不加班");
        }

        private static string GetCrossPageTableRequiredMessage(string fallbackReason)
        {
            return string.IsNullOrWhiteSpace(fallbackReason) || fallbackReason.Contains("跨页") || fallbackReason.Contains("拆分位置")
                ? "请确认当前表格是跨页表格。只有跨页表格才能按续表拆分。"
                : fallbackReason;
        }

        private static bool TryGetCrossPageSplitCell(Word.Table table, out Word.Cell splitCell, out string reason)
        {
            splitCell = null;
            reason = string.Empty;

            if (table?.Range == null)
            {
                reason = "请先选中跨页表格，或把光标放到跨页表格中。";
                return false;
            }

            int firstPage = GetRangeContentStartPageNumber(table.Range);
            if (firstPage <= 0)
            {
                reason = "无法识别当前表格所在页码。";
                return false;
            }

            int rowCount;
            try
            {
                rowCount = table.Rows.Count;
            }
            catch
            {
                rowCount = 0;
            }

            if (rowCount <= 1)
            {
                reason = "当前表格没有足够的正文行用于续表拆分。";
                return false;
            }

            int tableEndPage = GetRangeContentEndPageNumber(table.Range);
            if (tableEndPage <= firstPage)
            {
                reason = "当前表格不是跨页表格。";
                return false;
            }

            // 优先通过逐个单元格扫描获取准确的跨页单元格，无论是普通单元格还是纵向合并单元格，都能被精确识别。
            if (TryGetCrossPageCellByScanningCells(table, firstPage, out Word.Cell exactCrossPageCell))
            {
                splitCell = exactCrossPageCell;
                return true;
            }

            int firstNextPageRowIndex = TryFindFirstRowStartingAfterPage(table, firstPage, rowCount);
            if (firstNextPageRowIndex > 0)
            {
                int scanStart = Math.Max(2, firstNextPageRowIndex - 2);
                for (int i = scanStart; i < firstNextPageRowIndex; i++)
                {
                    try
                    {
                        Word.Row row = table.Rows[i];
                        if (TryGetCrossPageCellInRow(row, firstPage, out Word.Cell crossPageCell))
                        {
                            splitCell = crossPageCell;
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }

                splitCell = GetFirstCellInRow(table, firstNextPageRowIndex);
                if (splitCell != null)
                {
                    return true;
                }
            }

            if (TryGetCrossPageCellByScanningRows(table, firstPage, rowCount, out Word.Cell rowFallbackCell))
            {
                splitCell = rowFallbackCell;
                return true;
            }

            reason = "当前表格没有找到跨页拆分位置。";
            return false;
        }

        private static int TryFindFirstRowStartingAfterPage(Word.Table table, int firstPage, int rowCount)
        {
            if (table?.Rows == null || firstPage <= 0 || rowCount <= 1)
            {
                return 0;
            }

            int low = 2;
            int high = rowCount;
            int result = 0;

            try
            {
                while (low <= high)
                {
                    int mid = low + ((high - low) / 2);
                    Word.Row row = table.Rows[mid];
                    int rowStartPage = GetRangeContentStartPageNumber(row?.Range);
                    if (rowStartPage <= 0)
                    {
                        return 0;
                    }

                    if (rowStartPage > firstPage)
                    {
                        result = mid;
                        high = mid - 1;
                    }
                    else
                    {
                        low = mid + 1;
                    }
                }
            }
            catch
            {
                return 0;
            }

            return result;
        }

        private static bool TryGetCrossPageCellByScanningRows(
            Word.Table table,
            int firstPage,
            int rowCount,
            out Word.Cell splitCell)
        {
            splitCell = null;
            if (table?.Rows == null || firstPage <= 0 || rowCount <= 1)
            {
                return false;
            }

            try
            {
                for (int i = 2; i <= rowCount; i++)
                {
                    Word.Row row = table.Rows[i];
                    if (row?.Range == null)
                    {
                        continue;
                    }

                    int rowStartPage = GetRangeContentStartPageNumber(row.Range);
                    if (rowStartPage > firstPage)
                    {
                        splitCell = GetFirstCellInRow(table, i);
                        return splitCell != null;
                    }

                    if (TryGetCrossPageCellInRow(row, firstPage, out Word.Cell crossPageCell))
                    {
                        splitCell = crossPageCell;
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool TryGetCrossPageCellByScanningCells(Word.Table table, int firstPage, out Word.Cell splitCell)
        {
            splitCell = null;
            if (table?.Range == null || firstPage <= 0)
            {
                return false;
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
                    if (rowIndex <= 1)
                    {
                        continue;
                    }

                    int cellStartPage = GetRangeContentStartPageNumber(cell.Range);
                    int cellEndPage = GetRangeContentEndPageNumber(cell.Range);
                    if (cellStartPage > firstPage || (cellStartPage <= firstPage && cellEndPage > firstPage))
                    {
                        splitCell = cellStartPage > firstPage
                            ? GetFirstCellInRow(table, rowIndex) ?? cell
                            : cell;
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
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

        private static bool TryGetCrossPageCellInRow(Word.Row row, int firstPage, out Word.Cell crossPageCell)
        {
            crossPageCell = null;
            if (row?.Range == null || firstPage <= 0)
            {
                return false;
            }

            try
            {
                int rowStartPage = GetRangeContentStartPageNumber(row.Range);
                int rowEndPage = GetRangeContentEndPageNumber(row.Range);

                if (rowStartPage <= firstPage && rowEndPage > firstPage)
                {
                    try
                    {
                        if (row.Cells != null && row.Cells.Count > 0)
                        {
                            crossPageCell = row.Cells[1];
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }

                if (row.Cells != null)
                {
                    for (int i = 1; i <= row.Cells.Count; i++)
                    {
                        Word.Cell cell = row.Cells[i];
                        if (cell?.Range == null)
                        {
                            continue;
                        }

                        int cellStartPage = GetRangeContentStartPageNumber(cell.Range);
                        int cellEndPage = GetRangeContentEndPageNumber(cell.Range);
                        if (cellStartPage <= firstPage && cellEndPage > firstPage)
                        {
                            crossPageCell = cell;
                            return true;
                        }
                    }
                }
            }
            catch
            {
            }

            return false;
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

        private static int GetTableCount(Word.Document doc)
        {
            if (doc == null)
            {
                return 0;
            }

            try
            {
                return doc.Tables?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static bool TryExecuteNativeSplitTableThroughWordBasic(Word.Application app)
        {
            if (app == null)
            {
                return false;
            }

            try
            {
                object wordBasic = app.WordBasic;
                wordBasic.GetType().InvokeMember(
                    "TableSplitTable",
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

        private static void SplitSelectedCrossPageTable(Word.Application app, Word.Document doc, Word.Selection selection)
        {
            if (app == null)
            {
                MessageBox.Show("当前没有活动的 Word 应用。", "文档不加班");
                return;
            }

            Word.Table targetTable = GetTargetTableFromSelection(selection);
            if (targetTable?.Range == null)
            {
                MessageBox.Show("请先选中跨页表格，或把光标放到跨页表格中。", "文档不加班");
                return;
            }

            if (!IsCrossPageTable(targetTable))
            {
                MessageBox.Show("当前选中的表格不是跨页表格。", "文档不加班");
                return;
            }

            if (!TryGetTableSplitRowIndex(targetTable, out int splitRowIndex))
            {
                MessageBox.Show("当前跨页表格的分页发生在单行内部，暂不支持按续表拆分。", "文档不加班");
                return;
            }

            Word.Paragraph captionParagraph = GetNearestNonEmptyParagraphBeforeTableParagraph(doc, targetTable);
            string originalCaptionText = captionParagraph == null
                ? string.Empty
                : NormalizeParagraphText(captionParagraph.Range.Text);
            string baseCaption = ExtractTableCaptionBase(originalCaptionText);
            if (string.IsNullOrWhiteSpace(baseCaption))
            {
                MessageBox.Show("未识别到当前表格的表题，无法生成“表X（续）”。", "文档不加班");
                return;
            }

            string continuationCaption = baseCaption + "（续）";
            string captionStyleName = captionParagraph == null
                ? string.Empty
                : ResolveStyleName(TryGetParagraphStyle(captionParagraph.Range), doc);

            int originalTableStart = targetTable.Range.Start;
            int splitRowStart = 0;
            try
            {
                Word.Row splitRow = targetTable.Rows[splitRowIndex];
                splitRowStart = splitRow.Range.Start;
                Word.Range splitRange = splitRow.Cells[1].Range;
                splitRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                splitRange.Select();
                app.Selection.SplitTable();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"拆分表格失败: {ex.Message}", "文档不加班");
                return;
            }

            Word.Table firstTable = FindTableContainingPosition(doc, originalTableStart);
            Word.Table continuationTable = FindNextTableAfter(doc, firstTable?.Range?.End ?? splitRowStart);
            if (firstTable?.Range == null || continuationTable?.Range == null)
            {
                MessageBox.Show("表格拆分后未找到续表部分。", "文档不加班");
                return;
            }

            InsertContinuationCaptionBetweenTables(doc, firstTable, continuationTable, continuationCaption, captionStyleName);
            int manualHeaderCount = GetEffectiveHeaderRowCount(doc, targetTable);
            CopyHeaderRowToContinuationTable(app, firstTable, continuationTable, manualHeaderCount, true, out _);
            RememberHeaderRowCount(doc, firstTable, manualHeaderCount);
            RememberHeaderRowCount(doc, continuationTable, manualHeaderCount);
            MarkTopRowsAsHeading(firstTable, manualHeaderCount);
            MarkTopRowsAsHeading(continuationTable, manualHeaderCount);

            continuationTable.Select();
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

        private static Word.Table FindPreviousTableBefore(Word.Document doc, int position)
        {
            if (doc == null)
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

                    if (table.Range.Start >= position)
                    {
                        break;
                    }

                    previousTable = table;
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

        private static void InsertContinuationCaptionBetweenTables(
            Word.Document doc,
            Word.Table firstTable,
            Word.Table continuationTable,
            string continuationCaption,
            string captionStyleName)
        {
            if (doc == null
                || firstTable?.Range == null
                || continuationTable?.Range == null
                || string.IsNullOrWhiteSpace(continuationCaption))
            {
                return;
            }

            try
            {
                Word.Range betweenRange = doc.Range(firstTable.Range.End, continuationTable.Range.Start);
                if (betweenRange == null)
                {
                    return;
                }

                betweenRange.Text = "\f" + continuationCaption + "\r";
                Word.Range captionRange = doc.Range(firstTable.Range.End + 1, continuationTable.Range.Start);

                if (!string.IsNullOrWhiteSpace(captionStyleName)
                    && captionRange.Paragraphs != null
                    && captionRange.Paragraphs.Count > 0)
                {
                    TrySetStyle(captionRange.Paragraphs[1].Range, captionStyleName);
                }
            }
            catch
            {
            }
        }

        private static void InsertContinuationCaptionBeforeTable(
            Word.Application app,
            Word.Table firstTable,
            Word.Table continuationTable,
            string continuationCaption,
            string captionStyleName,
            Word.Range captionStyleSourceRange)
        {
            if (app == null || continuationTable?.Range == null || string.IsNullOrWhiteSpace(continuationCaption))
            {
                return;
            }

            try
            {
                continuationTable.Range.InsertParagraphBefore();
                Word.Paragraph captionParagraph = FindParagraphImmediatelyBeforeTable(continuationTable);
                Word.Range captionRange = captionParagraph?.Range;
                if (captionRange == null)
                {
                    return;
                }

                captionRange.Text = continuationCaption + "\r";
                Word.Range textRange = captionRange.Duplicate;
                textRange.End = Math.Max(textRange.Start, textRange.End - 1);

                ApplyCaptionFormatting(captionStyleSourceRange, textRange, captionStyleName);
                textRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                textRange.ParagraphFormat.PageBreakBefore = NeedsPageBreakBeforeContinuationCaption(firstTable, textRange) ? -1 : 0;
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

        private void button20_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                Word.Selection selection = app?.Selection;
                if (doc == null || selection?.Range == null)
                {
                    MessageBox.Show("请先在表格中选中表头区域。", "文档不加班");
                    return;
                }

                Word.Table targetTable = GetTargetTableFromSelection(selection);
                if (targetTable?.Range == null)
                {
                    MessageBox.Show("请先在需要拆分的表格中选中表头区域。", "文档不加班");
                    return;
                }

                int manualHeaderRowCount = GetSelectedHeaderRowCount(selection, targetTable);
                if (manualHeaderRowCount < 1)
                {
                    MessageBox.Show("未识别到有效的表头选区，请重新选择。", "文档不加班");
                    return;
                }

                RememberHeaderRowCount(doc, targetTable, manualHeaderRowCount);
                MessageBox.Show($"已将当前表格表头设置为前 {manualHeaderRowCount} 行。", "文档不加班");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置表头失败: {ex.Message}", "文档不加班");
            }
        }

        private static Word.Range InsertContinuationCaptionAfterSplit(
            Word.Application app,
            Word.Table continuationTable,
            string continuationCaption,
            string captionStyleName,
            Word.Range captionStyleSourceRange)
        {
            if (app == null || continuationTable?.Range == null || string.IsNullOrWhiteSpace(continuationCaption))
            {
                return null;
            }

            try
            {
                continuationTable.Select();
                Word.Selection selection = app.Selection;
                if (selection == null)
                {
                    return null;
                }

                selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.MoveLeft(Word.WdUnits.wdCharacter, 1);
                selection.TypeText(continuationCaption);

                Word.Range captionRange = selection.Range.Duplicate;
                if (captionRange.Start >= continuationCaption.Length)
                {
                    captionRange.Start -= continuationCaption.Length;
                }

                DeleteOneLineBreakBeforeRange(captionRange);
                ApplyCaptionFormatting(captionStyleSourceRange, captionRange, captionStyleName);
                captionRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                return captionRange;
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

        private static bool CopyHeaderRowToContinuationTable(
            Word.Application app,
            Word.Table sourceTable,
            Word.Table continuationTable,
            int headerRowCount,
            bool isFirstRound,
            out string message)
        {
            message = string.Empty;
            if (sourceTable?.Rows == null
                || sourceTable.Rows.Count < 1
                || continuationTable?.Rows == null
                || continuationTable.Rows.Count < 1)
            {
                message = "未找到可复制的标题行或续表。";
                return false;
            }

            try
            {
                if (app == null)
                {
                    message = "当前没有活动的 Word 应用，无法复制标题行。";
                    return false;
                }

                Word.Document doc = app.ActiveDocument;
                if (doc == null)
                {
                    message = "当前没有活动文档。";
                    return false;
                }

                Word.Selection selection = app.Selection;
                if (selection == null)
                {
                    message = "无法获取当前 Word 选区。";
                    return false;
                }

                if (headerRowCount < 1)
                {
                    message = "无法识别正表标题行数。";
                    return false;
                }

                int continuationTableStart = continuationTable.Range.Start;

                // 核心前提：在继续处理第一页之前，必须先在续表前面插入两个换行符作为隔离保护
                Word.Paragraph paragraphBeforeContinuation = FindParagraphImmediatelyBeforeTable(continuationTable);
                Word.Range separatorAnchor = paragraphBeforeContinuation?.Range?.Duplicate;
                if (separatorAnchor == null)
                {
                    message = "无法定位续表前的插入位置。";
                    return false;
                }
                separatorAnchor.InsertParagraphBefore();
                separatorAnchor.InsertParagraphBefore();

                if (isFirstRound)
                {
                    int bodyStartRow = headerRowCount + 1;
                    bool needsFirstPageSplit = sourceTable.Rows.Count >= bodyStartRow;
                    int sourceTableStart = sourceTable.Range.Start;

                    Word.Table headerOnlyTable = null;
                    Word.Table bodyPage1Table = null;

                    if (needsFirstPageSplit)
                    {
                        Word.Cell splitCell = GetFirstCellInRow(sourceTable, bodyStartRow);
                        if (splitCell == null)
                        {
                            message = "无法定位第一页正文的首个单元格进行拆分。";
                            return false;
                        }

                        Word.Range splitRange = splitCell.Range;
                        splitRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        splitRange.Select();

                        if (!TrySplitTableAtSelection(app, doc, selection, out string splitError))
                        {
                            message = "拆分表头和正文失败：" + splitError;
                            return false;
                        }

                        headerOnlyTable = FindTableContainingPosition(doc, sourceTableStart);
                        bodyPage1Table = FindNextTableAfter(doc, headerOnlyTable?.Range?.End ?? sourceTableStart);

                        if (headerOnlyTable == null || bodyPage1Table == null)
                        {
                            message = "拆分表头和正文后无法重新获取表格。";
                            return false;
                        }
                    }
                    else
                    {
                        headerOnlyTable = sourceTable;
                    }

                    // 复制表头（此时表头独占一个表格）
                    headerOnlyTable.Select();
                    selection.Copy();

                    // 还原第一页（如果发生过拆分）：删除第一页表头和正文之间的空白行
                    if (needsFirstPageSplit && headerOnlyTable != null && bodyPage1Table != null)
                    {
                        DeleteParagraphBetweenTables(headerOnlyTable, bodyPage1Table);
                    }
                }

                // 重新获取续表（前面的文档结构变动可能导致原 COM 对象或范围失效）
                Word.Table restoredContinuationTable = FindTableContainingPosition(doc, continuationTableStart) ?? FindNextTableAfter(doc, sourceTable.Range.End);
                if (restoredContinuationTable == null)
                {
                    // 若定位不到，可以回退到原本的 continuationTable
                    restoredContinuationTable = continuationTable;
                }

                Word.Paragraph headerPasteParagraph = FindParagraphBeforeTable(restoredContinuationTable, 1);
                if (headerPasteParagraph?.Range == null)
                {
                    message = "无法定位靠近续表的粘贴段落。";
                    return false;
                }

                Word.Range headerPasteRange = headerPasteParagraph.Range.Duplicate;
                headerPasteRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                selection.SetRange(headerPasteRange.Start, headerPasteRange.Start);
                
                if (!TryPasteKeepSourceFormatting(selection))
                {
                    message = "无法将表头粘贴到第二页。";
                    return false;
                }

                Word.Table pastedHeaderTable = GetTargetTableFromSelection(selection);
                if (pastedHeaderTable?.Range == null)
                {
                    message = "已复制表头，但未找到新插入的表格。";
                    return false;
                }

                Word.Table continuationTableAfter = FindNextTableAfter(doc, pastedHeaderTable.Range.End);
                if (continuationTableAfter?.Range == null)
                {
                    message = "粘贴表头后未找到原始续表。";
                    return false;
                }

                MarkTopRowsAsHeading(pastedHeaderTable, headerRowCount);
                MarkTopRowsAsHeading(continuationTableAfter, headerRowCount);
                DeleteParagraphBetweenTables(pastedHeaderTable, continuationTableAfter);

                return true;
            }
            catch (Exception ex)
            {
                message = "复制标题行失败: " + ex.Message;
                return false;
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

        private static bool TrySelectTopRows(Word.Table table, Word.Selection selection, int rowCount)
        {
            if (table?.Range?.Cells == null || selection == null || rowCount < 1)
            {
                return false;
            }

            try
            {
                Word.Cell topLeftCell = GetTopLeftCell(table);
                if (topLeftCell?.Range == null)
                {
                    return false;
                }

                Word.Range anchor = topLeftCell.Range.Duplicate;
                anchor.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                anchor.Select();

                for (int i = 0; i < rowCount; i++)
                {
                    if (!TryExtendSelectionToNextRow(selection, i == 0))
                    {
                        return false;
                    }
                }

                return true;
            }
            catch
            {
                return false;
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
            if (doc != null && table?.Range != null && TryGetManualHeaderOverride(doc, table, out int manualRowCount))
            {
                return manualRowCount;
            }

            return GetNormalizedHeaderRowCount(table);
        }

        private static int GetSelectedHeaderRowCount(Word.Selection selection, Word.Table table)
        {
            if (selection?.Range == null || table?.Range?.Cells == null)
            {
                return 0;
            }

            int selectionStart = selection.Range.Start;
            int selectionEnd = selection.Range.End;
            int maxRow = 0;
            bool hasAny = false;

            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    if (cell?.Range == null)
                    {
                        continue;
                    }

                    int cellStart = cell.Range.Start;
                    int cellEnd = cell.Range.End;
                    bool intersects = cellStart < selectionEnd && cellEnd > selectionStart;
                    if (!intersects)
                    {
                        continue;
                    }

                    int rowIndex = GetCellRowIndex(cell);
                    if (rowIndex <= 0)
                    {
                        continue;
                    }

                    hasAny = true;
                    if (rowIndex > maxRow)
                    {
                        maxRow = rowIndex;
                    }
                }
            }
            catch
            {
                return 0;
            }

            if (!hasAny || maxRow < 1)
            {
                return 0;
            }

            int safeTotalRows = table.Rows?.Count ?? maxRow;
            if (maxRow > safeTotalRows)
            {
                maxRow = safeTotalRows;
            }

            return maxRow;
        }

        private static void SaveManualHeaderOverride(Word.Document doc, Word.Table table, int headerRowCount)
        {
            RememberHeaderRowCount(doc, table, headerRowCount);
        }

        private sealed class TableHeaderMemoryItem
        {
            public Word.Table Table { get; set; }
            public int HeaderRowCount { get; set; }
        }

        private static void RememberHeaderRowCount(Word.Document doc, Word.Table table, int headerRowCount)
        {
            if (doc == null || table?.Range == null || headerRowCount < 1)
            {
                return;
            }

            string docKey = GetDocumentOverrideKey(doc);
            if (string.IsNullOrWhiteSpace(docKey))
            {
                return;
            }

            if (!ManualHeaderRowOverridesByDocument.TryGetValue(docKey, out Dictionary<int, int> tableMap))
            {
                tableMap = new Dictionary<int, int>();
                ManualHeaderRowOverridesByDocument[docKey] = tableMap;
            }

            tableMap[table.Range.Start] = headerRowCount;
        }

        private static bool TryGetManualHeaderOverride(Word.Document doc, Word.Table table, out int headerRowCount)
        {
            headerRowCount = 0;
            if (doc == null || table?.Range == null)
            {
                return false;
            }

            string docKey = GetDocumentOverrideKey(doc);
            if (string.IsNullOrWhiteSpace(docKey))
            {
                return false;
            }

            if (!ManualHeaderRowOverridesByDocument.TryGetValue(docKey, out Dictionary<int, int> tableMap) || tableMap == null)
            {
                return false;
            }

            if (!tableMap.TryGetValue(table.Range.Start, out int manualRowCount) || manualRowCount < 1)
            {
                return false;
            }

            int totalRows = GetTableRowCountSafe(table, manualRowCount);
            headerRowCount = Math.Min(manualRowCount, Math.Max(1, totalRows));
            return true;
        }

        private static int NormalizeTables(
            Word.Application app,
            Word.Document doc,
            IList<Word.Table> tables,
            TableFormattingOptions options,
            TableFormattingProgressForm progressForm,
            string scopeLabel)
        {
            if (app == null || doc == null || options == null || tables == null)
            {
                return 0;
            }

            List<Word.Table> validTables = tables
                .Where(table => table?.Range != null)
                .ToList();
            if (validTables.Count == 0)
            {
                return 0;
            }

            float tableWidthPoints = ConvertCentimetersToPoints(app, options.TableWidthCentimeters);
            for (int i = 0; i < validTables.Count; i++)
            {
                Word.Table table = validTables[i];

                int percent = Math.Min(96, 5 + (int)Math.Round(((i + 1) * 90d) / Math.Max(1, validTables.Count)));
                progressForm?.ReportProgress(
                    percent,
                    $"正在规范第 {i + 1} / {validTables.Count} 个表格...",
                    $"正在统一{scopeLabel}的表头、正文、宽度和边框。");

                FormatSingleTable(doc, table, tableWidthPoints, options);
            }

            return validTables.Count;
        }

        private static List<Word.Table> CollectTablesAfterToc(Word.Document doc)
        {
            List<Word.Table> tables = new List<Word.Table>();
            if (doc == null)
            {
                return tables;
            }

            int scanStart = GetScanStartAfterToc(doc);
            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    if (table?.Range == null)
                    {
                        continue;
                    }

                    if (scanStart > 0 && table.Range.End <= scanStart)
                    {
                        continue;
                    }

                    tables.Add(table);
                }
            }
            catch
            {
            }

            return tables;
        }

        private static int NormalizeAllImages(
            Word.Document doc,
            ImageFormattingOptions options,
            ImageFormattingProgressForm progressForm)
        {
            if (doc == null || options == null)
            {
                return 0;
            }

            List<Action> imageActions = CollectImageFormattingActions(doc, options);
            if (imageActions.Count == 0)
            {
                return 0;
            }

            for (int i = 0; i < imageActions.Count; i++)
            {
                int percent = Math.Min(96, 5 + (int)Math.Round(((i + 1) * 90d) / Math.Max(1, imageActions.Count)));
                progressForm?.ReportProgress(
                    percent,
                    $"正在规范第 {i + 1} / {imageActions.Count} 张图片...",
                    "正在统一图片缩放比例，并整理所属段落格式。");

                imageActions[i]?.Invoke();
            }

            return imageActions.Count;
        }

        private static List<Action> CollectImageFormattingActions(Word.Document doc, ImageFormattingOptions options)
        {
            List<Action> actions = new List<Action>();
            if (doc == null || options == null)
            {
                return actions;
            }

            try
            {
                foreach (Word.InlineShape inlineShape in doc.InlineShapes)
                {
                    if (!IsPictureInlineShape(inlineShape) || IsRangeOnFirstPage(inlineShape.Range))
                    {
                        continue;
                    }

                    actions.Add(() => FormatInlineShape(inlineShape, options));
                }
            }
            catch
            {
            }

            try
            {
                foreach (Word.Shape shape in doc.Shapes)
                {
                    if (!IsPictureShape(shape) || IsRangeOnFirstPage(shape.Anchor))
                    {
                        continue;
                    }

                    actions.Add(() => FormatShape(shape, options));
                }
            }
            catch
            {
            }

            return actions;
        }

        private static bool IsRangeOnFirstPage(Word.Range range)
        {
            if (range == null)
            {
                return false;
            }

            try
            {
                int startPage = GetRangeContentStartPageNumber(range);
                return startPage == 1;
            }
            catch
            {
                return false;
            }
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

        private static void FormatInlineShape(Word.InlineShape inlineShape, ImageFormattingOptions options)
        {
            if (inlineShape == null || options == null)
            {
                return;
            }

            try
            {
                inlineShape.LockAspectRatio = Office.MsoTriState.msoTrue;
            }
            catch
            {
            }

            try
            {
                inlineShape.ScaleHeight = options.ScalePercent;
                inlineShape.ScaleWidth = options.ScalePercent;
            }
            catch
            {
            }

            try
            {
                ApplyParagraphFormatting(inlineShape.Range?.ParagraphFormat);
            }
            catch
            {
            }
        }

        private static void FormatShape(Word.Shape shape, ImageFormattingOptions options)
        {
            if (shape == null || options == null)
            {
                return;
            }

            try
            {
                shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            }
            catch
            {
            }

            try
            {
                shape.ScaleHeight(options.ScalePercent, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
                shape.ScaleWidth(options.ScalePercent, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromTopLeft);
            }
            catch
            {
            }

            try
            {
                shape.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
                shape.Left = (float)Word.WdShapePosition.wdShapeCenter;
            }
            catch
            {
            }

            try
            {
                ApplyParagraphFormatting(shape.Anchor?.ParagraphFormat);
            }
            catch
            {
            }
        }

        private static void ApplyParagraphFormatting(Word.ParagraphFormat paragraphFormat)
        {
            if (paragraphFormat == null)
            {
                return;
            }

            try
            {
                paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            catch
            {
            }

            try
            {
                paragraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            }
            catch
            {
            }

            try
            {
                paragraphFormat.LeftIndent = 0f;
                paragraphFormat.RightIndent = 0f;
                paragraphFormat.FirstLineIndent = 0f;
            }
            catch
            {
            }

            try
            {
                paragraphFormat.CharacterUnitLeftIndent = 0f;
                paragraphFormat.CharacterUnitRightIndent = 0f;
                paragraphFormat.CharacterUnitFirstLineIndent = 0f;
            }
            catch
            {
            }
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

        private static List<TableHeaderMemoryItem> CaptureTrailingTableHeaderMemories(Word.Document doc, int position)
        {
            List<TableHeaderMemoryItem> memories = new List<TableHeaderMemoryItem>();
            if (doc == null)
            {
                return memories;
            }

            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    if (table?.Range == null || table.Range.Start <= position)
                    {
                        continue;
                    }

                    int headerRowCount = GetEffectiveHeaderRowCount(doc, table);
                    if (headerRowCount < 1)
                    {
                        continue;
                    }

                    memories.Add(new TableHeaderMemoryItem
                    {
                        Table = table,
                        HeaderRowCount = headerRowCount
                    });
                }
            }
            catch
            {
            }

            return memories;
        }

        private static void RestoreTrailingTableHeaderMemories(Word.Document doc, List<TableHeaderMemoryItem> memories)
        {
            if (doc == null || memories == null || memories.Count == 0)
            {
                return;
            }

            foreach (TableHeaderMemoryItem memory in memories)
            {
                if (memory?.Table?.Range == null || memory.HeaderRowCount < 1)
                {
                    continue;
                }

                RememberHeaderRowCount(doc, memory.Table, memory.HeaderRowCount);
                MarkTopRowsAsHeading(memory.Table, memory.HeaderRowCount);
            }
        }

        private static string GetDocumentOverrideKey(Word.Document doc)
        {
            if (doc == null)
            {
                return string.Empty;
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(doc.FullName))
                {
                    return doc.FullName;
                }
            }
            catch
            {
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

            return string.Empty;
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

        private static void DeleteParagraphBetweenTables(Word.Table upperTable, Word.Table lowerTable)
        {
            if (upperTable?.Range == null || lowerTable?.Range == null)
            {
                return;
            }

            try
            {
                if (lowerTable.Range.Start <= upperTable.Range.End)
                {
                    return;
                }

                // 优先从下表前向前清理换行，避免 Word 表格边界残留空段落。
                DeleteLineBreaksBeforeRange(lowerTable.Range, 32, upperTable.Range.End);
                if (lowerTable.Range.Start <= upperTable.Range.End)
                {
                    return;
                }

                Word.Range gapRange = upperTable.Range.Duplicate;
                gapRange.SetRange(upperTable.Range.End, lowerTable.Range.Start);
                try
                {
                    gapRange.Text = string.Empty;
                }
                catch
                {
                    gapRange.Delete();
                }
                return;
            }
            catch
            {
            }

            try
            {
                Word.Range fallbackRange = upperTable.Range.Duplicate;
                fallbackRange.SetRange(upperTable.Range.End, lowerTable.Range.Start);
                fallbackRange.Select();
                Word.Selection selection = Globals.ThisAddIn?.Application?.Selection;
                if (selection != null)
                {
                    selection.Delete();
                }
            }
            catch
            {
            }
        }

        private static void DeleteParagraphRange(Word.Paragraph paragraph)
        {
            if (paragraph?.Range == null)
            {
                return;
            }

            try
            {
                Word.Range paragraphRange = paragraph.Range.Duplicate;
                int documentStart = paragraphRange.Document?.Content?.Start ?? 0;

                // 题注删掉后，顺手清理它前面紧邻的空行，避免上下两个表之间残留空白段落。
                DeleteLineBreaksBeforeRange(paragraphRange, 8, documentStart);

                try
                {
                    paragraphRange.Text = string.Empty;
                }
                catch
                {
                    paragraphRange.Delete();
                }
            }
            catch
            {
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

        private static bool TryDeleteRowBlock(Word.Application app, Word.Table table, int startRow, int endRow)
        {
            if (app == null || table == null || startRow < 1 || endRow < startRow)
            {
                return false;
            }

            try
            {
                int deleteCount = endRow - startRow + 1;
                bool deletedAny = false;
                for (int i = 0; i < deleteCount; i++)
                {
                    if (!TryDeleteSingleRow(app, table, startRow))
                    {
                        return false;
                    }

                    deletedAny = true;
                }

                return deletedAny;
            }
            catch
            {
                return false;
            }
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

        private static bool ShouldDeleteContinuationHeaderBeforeMerge(
            Word.Document doc,
            Word.Table previousTable,
            Word.Table continuationTable)
        {
            if (previousTable?.Range == null || continuationTable?.Range == null)
            {
                return false;
            }

            int headerRowCount = GetMergeContinuationHeaderRowCount(doc, previousTable, continuationTable);
            if (headerRowCount <= 0)
            {
                return false;
            }

            if (HasManualHeaderOverrideForMerge(doc, previousTable, continuationTable))
            {
                return true;
            }

            if (DoHeaderRowsMatchForMerge(previousTable, continuationTable, headerRowCount))
            {
                return true;
            }

            if (!TryGetContinuationComparisonCell(continuationTable, out int compareColumnIndex, out string continuationCellText))
            {
                return false;
            }

            if (!DoesPreviousTableHeaderMatch(doc, previousTable, continuationTable, compareColumnIndex, continuationCellText))
            {
                return false;
            }

            return headerRowCount > 0;
        }

        private static bool DoHeaderRowsMatchForMerge(
            Word.Table previousTable,
            Word.Table continuationTable,
            int headerRowCount)
        {
            if (previousTable?.Range == null || continuationTable?.Range == null || headerRowCount < 1)
            {
                return false;
            }

            int previousRowCount = GetTableRowCountSafe(previousTable, headerRowCount);
            int continuationRowCount = GetTableRowCountSafe(continuationTable, headerRowCount);
            int safeHeaderRowCount = Math.Min(
                Math.Max(1, headerRowCount),
                Math.Min(Math.Max(1, previousRowCount), Math.Max(1, continuationRowCount)));

            bool matchedAnyNonEmptyRow = false;
            for (int rowIndex = 1; rowIndex <= safeHeaderRowCount; rowIndex++)
            {
                string previousSignature = GetComparableRowSignature(previousTable, rowIndex);
                string continuationSignature = GetComparableRowSignature(continuationTable, rowIndex);

                if (string.IsNullOrWhiteSpace(previousSignature) && string.IsNullOrWhiteSpace(continuationSignature))
                {
                    continue;
                }

                matchedAnyNonEmptyRow = true;
                if (!string.Equals(previousSignature, continuationSignature, StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return matchedAnyNonEmptyRow;
        }

        private static bool HasManualHeaderOverrideForMerge(
            Word.Document doc,
            Word.Table previousTable,
            Word.Table continuationTable)
        {
            if (doc == null)
            {
                return false;
            }

            if (continuationTable?.Range != null && TryGetManualHeaderOverride(doc, continuationTable, out _))
            {
                return true;
            }

            if (previousTable?.Range != null && TryGetManualHeaderOverride(doc, previousTable, out _))
            {
                return true;
            }

            return false;
        }

        private static bool DoesPreviousTableHeaderMatch(
            Word.Document doc,
            Word.Table previousTable,
            Word.Table continuationTable,
            int compareColumnIndex,
            string continuationCellText)
        {
            if (previousTable?.Range == null
                || continuationTable?.Range == null
                || compareColumnIndex < 1
                || string.IsNullOrWhiteSpace(continuationCellText))
            {
                return false;
            }

            string firstRowCellText = GetNormalizedCellText(previousTable, 1, compareColumnIndex);
            if (string.Equals(continuationCellText, firstRowCellText, StringComparison.Ordinal))
            {
                return true;
            }

            int headerRowCount = GetMergeContinuationHeaderRowCount(doc, previousTable, continuationTable);
            int safeHeaderRowCount = Math.Min(
                Math.Max(1, headerRowCount),
                Math.Max(1, GetTableRowCountSafe(previousTable, headerRowCount)));

            for (int rowIndex = 1; rowIndex <= safeHeaderRowCount; rowIndex++)
            {
                string previousCellText = GetNormalizedCellText(previousTable, rowIndex, compareColumnIndex);
                if (string.Equals(continuationCellText, previousCellText, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static int GetMergeContinuationHeaderRowCount(
            Word.Document doc,
            Word.Table previousTable,
            Word.Table continuationTable)
        {
            if (doc != null)
            {
                if (continuationTable?.Range != null && TryGetManualHeaderOverride(doc, continuationTable, out int continuationManualHeaderCount))
                {
                    return continuationManualHeaderCount;
                }

                if (previousTable?.Range != null && TryGetManualHeaderOverride(doc, previousTable, out int previousManualHeaderCount))
                {
                    return previousManualHeaderCount;
                }
            }

            return GetEffectiveHeaderRowCount(doc, continuationTable);
        }

        private static bool TryGetContinuationComparisonCell(
            Word.Table table,
            out int columnIndex,
            out string cellText)
        {
            columnIndex = 0;
            cellText = string.Empty;

            if (table?.Range == null)
            {
                return false;
            }

            int maxColumnsToInspect = 12;
            for (int candidateColumnIndex = 1; candidateColumnIndex <= maxColumnsToInspect; candidateColumnIndex++)
            {
                string candidateCellText = GetNormalizedCellText(table, 1, candidateColumnIndex);
                if (string.IsNullOrWhiteSpace(candidateCellText))
                {
                    continue;
                }

                columnIndex = candidateColumnIndex;
                cellText = candidateCellText;
                return true;
            }

            return false;
        }

        private static string GetNormalizedCellText(Word.Table table, int rowIndex, int columnIndex)
        {
            if (table == null || rowIndex < 1 || columnIndex < 1)
            {
                return string.Empty;
            }

            try
            {
                Word.Cell cell = table.Cell(rowIndex, columnIndex);
                return NormalizeParagraphText(cell?.Range?.Text);
            }
            catch
            {
                return string.Empty;
            }
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
            RememberTableSelectionAction(button5, button5_Click);
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

        private void button22_Click(object sender, RibbonControlEventArgs e)
        {
            RememberTableSelectionAction(button22, button22_Click);
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

                SelectNextContinuationTableAfterCursor(doc, cursorRange);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择下个续表失败: {ex.Message}", "文档不加班");
            }
        }

        private static void SelectNextContinuationTableAfterCursor(Word.Document doc, Word.Range cursorRange)
        {
            if (doc == null || cursorRange == null)
            {
                MessageBox.Show("请先把光标放到 Word 文档正文中。", "文档不加班");
                return;
            }

            int cursorPosition = cursorRange.Start;
            Word.Table continuationTable = FindNextContinuationTable(doc, cursorPosition);
            if (continuationTable == null)
            {
                MessageBox.Show("从当前光标位置开始，未找到下个续表。", "文档不加班");
                return;
            }

            continuationTable.Select();
        }

        private static Word.Table FindNextContinuationTable(Word.Document doc, int cursorPosition)
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

                    if (table.Range.Start < cursorPosition)
                    {
                        continue;
                    }

                    string captionText = GetNearestNonEmptyParagraphBeforeTable(doc, table);
                    if (IsContinuationTableCaption(captionText))
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

        private static string GetNearestNonEmptyParagraphBeforeTable(Word.Document doc, Word.Table table)
        {
            Word.Paragraph paragraph = GetNearestNonEmptyParagraphBeforeTableParagraph(doc, table);
            return paragraph == null ? string.Empty : NormalizeParagraphText(paragraph.Range.Text);
        }

        private static bool IsContinuationTableCaption(string paragraphText)
        {
            if (string.IsNullOrWhiteSpace(paragraphText))
            {
                return false;
            }

            return Regex.IsMatch(
                paragraphText,
                @"^表\s*[0-9０-９一二三四五六七八九十百千]+(?:\s*[\.．\-—]\s*[0-9０-９一二三四五六七八九十百千]+)*\s*[（(]\s*续\s*[）)]\s*$",
                RegexOptions.IgnoreCase);
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

        private void button19_Click(object sender, RibbonControlEventArgs e)
        {
            RememberTableSelectionAction(button19, button19_Click);
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                SelectAllTablesAfterToc(doc);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择全部表格失败: {ex.Message}", "文档不加班");
            }
        }

        private static void SelectAllTablesAfterToc(Word.Document doc)
        {
            if (doc == null)
            {
                return;
            }

            int scanStart = GetScanStartAfterToc(doc);
            List<Word.Table> matchedTables = new List<Word.Table>();

            try
            {
                foreach (Word.Table table in doc.Tables)
                {
                    if (table?.Range == null)
                    {
                        continue;
                    }

                    if (scanStart > 0 && table.Range.End <= scanStart)
                    {
                        continue;
                    }

                    matchedTables.Add(table);
                }
            }
            catch
            {
            }

            if (matchedTables.Count == 0)
            {
                string message = scanStart > 0
                    ? "目录后面没有找到可选取的表格。"
                    : "当前文档中没有找到可选取的表格。";
                MessageBox.Show(message, "文档不加班");
                return;
            }

            SelectEditableTableRanges(doc, matchedTables);

            string successMessage = scanStart > 0
                ? $"已选中目录后的 {matchedTables.Count} 个表格。"
                : $"已选中文档中的 {matchedTables.Count} 个表格。";
            MessageBox.Show(successMessage, "文档不加班");
        }

        private static void SelectEditableTableRanges(Word.Document doc, IList<Word.Table> tables)
        {
            if (doc == null || tables == null || tables.Count == 0)
            {
                return;
            }

            object editor = Word.WdEditorType.wdEditorEveryone;

            DeleteAllEditableRanges(doc, editor);
            try
            {
                foreach (Word.Table table in tables)
                {
                    Word.Range range = table?.Range;
                    if (range == null)
                    {
                        continue;
                    }

                    AddEditableRange(range, editor);
                }

                SelectAllEditableRanges(doc, editor);
            }
            finally
            {
                DeleteAllEditableRanges(doc, editor);
            }
        }

        private static void AddEditableRange(Word.Range range, object editor)
        {
            object[] args = { editor };
            range.Editors.GetType().InvokeMember(
                "Add",
                BindingFlags.InvokeMethod,
                null,
                range.Editors,
                args);
        }

        private static void SelectAllEditableRanges(Word.Document doc, object editor)
        {
            object[] args = { editor };
            doc.GetType().InvokeMember(
                "SelectAllEditableRanges",
                BindingFlags.InvokeMethod,
                null,
                doc,
                args);
        }

        private static void DeleteAllEditableRanges(Word.Document doc, object editor)
        {
            object[] args = { editor };
            doc.GetType().InvokeMember(
                "DeleteAllEditableRanges",
                BindingFlags.InvokeMethod,
                null,
                doc,
                args);
        }
    }
}
