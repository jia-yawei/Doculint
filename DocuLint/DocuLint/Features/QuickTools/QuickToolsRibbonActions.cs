using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    public partial class Ribbon1
    {
        private sealed class DocumentUndoScope : IDisposable
        {
            private readonly object undoRecord;
            private readonly bool customRecordStarted;

            public DocumentUndoScope(Word.Application app, string actionName)
            {
                if (app == null || string.IsNullOrWhiteSpace(actionName))
                {
                    return;
                }

                try
                {
                    undoRecord = TryGetComProperty(app, "UndoRecord");
                    if (undoRecord == null)
                    {
                        return;
                    }

                    undoRecord.GetType().InvokeMember(
                        "StartCustomRecord",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null,
                        undoRecord,
                        new object[] { actionName });
                    customRecordStarted = true;
                }
                catch
                {
                    customRecordStarted = false;
                }
            }

            public void Dispose()
            {
                if (!customRecordStarted || undoRecord == null)
                {
                    return;
                }

                try
                {
                    undoRecord.GetType().InvokeMember(
                        "EndCustomRecord",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null,
                        undoRecord,
                        null);
                }
                catch
                {
                }
            }
        }

        private void btnInsertTotalPages_Click(object sender, RibbonControlEventArgs e)
        {
            RememberInsertItemAction(btnInsertTotalPages, btnInsertTotalPages_Click);
            try
            {
                ExecuteDocumentCommand("插入总页码", InsertTotalPagesField);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入总页码失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void button7_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int updatedCount = ExecuteDocumentCommand("更新总页码", UpdateTotalPagesFields);
                if (updatedCount > 0)
                {
                    MessageBox.Show($"已更新 {updatedCount} 个总页码域。", "文档不加班 快速工具");
                }
                else
                {
                    MessageBox.Show("当前文档中没有找到可更新的总页码域。", "文档不加班 快速工具");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新总页码失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void button26_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                int refreshedCount = ExecuteDocumentCommand("刷新目录", RefreshTableOfContents);
                if (refreshedCount > 0)
                {
                    MessageBox.Show($"已刷新 {refreshedCount} 个目录。", "文档不加班 快速工具");
                }
                else
                {
                    MessageBox.Show("当前文档中没有找到可刷新的目录。", "文档不加班 快速工具");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刷新目录失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void btnTogglePageWhitespace_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                bool isVisible = TogglePageWhitespace();
                btnTogglePageWhitespace.Checked = isVisible;
            }
            catch (Exception ex)
            {
                btnTogglePageWhitespace.Checked = IsPageWhitespaceVisible();
                MessageBox.Show($"切换页面间空白失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private void button23_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!TryUndoLastCommand())
                {
                    MessageBox.Show("当前没有可撤销的文档操作。", "文档不加班 快速工具");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"撤销上次命令失败: {ex.Message}", "文档不加班 快速工具");
            }
        }

        private static void ExecuteDocumentCommand(string actionName, Action action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            Word.Application app = Globals.ThisAddIn.Application;
            using (new DocumentUndoScope(app, actionName))
            {
                action();
            }
        }

        private static T ExecuteDocumentCommand<T>(string actionName, Func<T> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            Word.Application app = Globals.ThisAddIn.Application;
            using (new DocumentUndoScope(app, actionName))
            {
                return action();
            }
        }

        private static bool TryUndoLastCommand()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            if (app == null)
            {
                return false;
            }

            try
            {
                object result = app.GetType().InvokeMember(
                    "Undo",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    app,
                    null);

                if (result is bool boolResult)
                {
                    return boolResult;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void InsertTotalPagesField()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Selection selection = app?.Selection;
            Word.Document doc = app?.ActiveDocument;
            if (selection?.Range == null)
            {
                throw new InvalidOperationException("当前没有可插入的位置。");
            }

            if (doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            PreparePaginationForTotalPages(app, doc);

            Word.Range insertRange = selection.Range.Duplicate;
            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            Word.Field field = insertRange.Fields.Add(
                insertRange,
                Word.WdFieldType.wdFieldNumPages,
                Type.Missing,
                false);
            field?.Update();

            try
            {
                doc.Fields.Update();
            }
            catch
            {
            }

            Word.Range endRange = field?.Result?.Duplicate ?? insertRange;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            endRange.Select();
        }

        private static int UpdateTotalPagesFields()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            PreparePaginationForTotalPages(app, doc);

            int updatedCount = 0;
            foreach (Word.Range storyRange in EnumerateStoryRanges(doc))
            {
                updatedCount += UpdateTotalPagesFieldsInRange(storyRange);
            }

            try
            {
                app.ScreenRefresh();
            }
            catch
            {
            }

            return updatedCount;
        }

        private static int RefreshTableOfContents()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                throw new InvalidOperationException("当前没有活动文档。");
            }

            int refreshedCount = 0;
            using (new WordPerformanceScope(app))
            {
                try
                {
                    doc.Repaginate();
                }
                catch
                {
                }

                try
                {
                    if (doc.TablesOfContents != null)
                    {
                        foreach (Word.TableOfContents toc in doc.TablesOfContents)
                        {
                            if (toc == null)
                            {
                                continue;
                            }

                            toc.Update();
                            refreshedCount++;
                        }
                    }
                }
                catch
                {
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
            }

            return refreshedCount;
        }

        private static void PreparePaginationForTotalPages(Word.Application app, Word.Document doc)
        {
            if (doc == null)
            {
                return;
            }

            try
            {
                doc.Repaginate();
            }
            catch
            {
            }

            try
            {
                doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, false);
            }
            catch
            {
            }

            try
            {
                app?.ScreenRefresh();
            }
            catch
            {
            }

            Application.DoEvents();
        }

        private static IEnumerable<Word.Range> EnumerateStoryRanges(Word.Document doc)
        {
            if (doc == null)
            {
                yield break;
            }

            Word.Range storyRange = null;
            try
            {
                storyRange = doc.StoryRanges?[Word.WdStoryType.wdMainTextStory];
            }
            catch
            {
            }

            while (storyRange != null)
            {
                yield return storyRange;

                Word.Range nextRange = null;
                try
                {
                    nextRange = storyRange.NextStoryRange;
                }
                catch
                {
                }

                storyRange = nextRange;
            }
        }

        private static int UpdateTotalPagesFieldsInRange(Word.Range range)
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
                    if (field == null)
                    {
                        continue;
                    }

                    if (field.Type != Word.WdFieldType.wdFieldNumPages)
                    {
                        continue;
                    }

                    field.Update();
                    updatedCount++;
                }
            }
            catch
            {
            }

            return updatedCount;
        }

        private static bool TogglePageWhitespace()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            if (app?.ActiveWindow == null)
            {
                throw new InvalidOperationException("当前没有活动文档窗口。");
            }

            Word.View view = app.ActiveWindow.View;
            if (view == null)
            {
                throw new InvalidOperationException("未能获取当前页面视图。");
            }

            try
            {
                view.Type = Word.WdViewType.wdPrintView;
            }
            catch
            {
            }

            bool current = IsPageWhitespaceVisible(view, app);
            bool next = !current;
            SetPageWhitespaceVisible(view, app, next);
            return IsPageWhitespaceVisible(view, app);
        }

        private static bool IsPageWhitespaceVisible()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            Word.View view = app?.ActiveWindow?.View;
            return IsPageWhitespaceVisible(view, app);
        }

        private static bool IsPageWhitespaceVisible(object view, Word.Application app)
        {
            object viewValue = TryGetComProperty(view, "DisplayPageBoundaries");
            if (viewValue != null)
            {
                return Convert.ToBoolean(viewValue);
            }

            object options = TryGetComProperty(app, "Options");
            object optionsValue = TryGetComProperty(options, "DisplayPageBoundaries");
            return optionsValue != null && Convert.ToBoolean(optionsValue);
        }

        private static void SetPageWhitespaceVisible(object view, Word.Application app, bool visible)
        {
            TrySetComProperty(view, "DisplayPageBoundaries", visible);

            object options = TryGetComProperty(app, "Options");
            TrySetComProperty(options, "DisplayPageBoundaries", visible);
        }
    }
}
