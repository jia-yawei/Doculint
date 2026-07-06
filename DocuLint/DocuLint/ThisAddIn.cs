using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Tools;

namespace DocuLint
{
    // Word插件主程序类
    public partial class ThisAddIn
    {
        // 右侧统一导航面板的界面控件
        private NavigationPaneControl navigationPaneControl;
        // Word右侧自定义任务面板
        private CustomTaskPane navigationTaskPane;
        // 需求追踪控制台界面控件
        private RequirementTrackingConsoleControl requirementTrackingConsoleControl;
        // 需求追踪独立任务窗格
        private CustomTaskPane requirementTrackingTaskPane;
        // 文档跳转辅助工具
        private WordDocumentHostAdapter documentHostAdapter;
        // 文档基本信息存储
        private DocumentBasicInfoStore documentBasicInfoStore;
        // 延迟刷新 Ribbon，避免拖选文字时同步读取 Selection 打断 Word 选区。
        private Timer styleRibbonRefreshTimer;

        // 插件启动时执行
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 初始化文档跳转工具
            documentHostAdapter = new WordDocumentHostAdapter(() => Application);
            documentBasicInfoStore = new DocumentBasicInfoStore();
            styleRibbonRefreshTimer = new Timer { Interval = 180 };
            styleRibbonRefreshTimer.Tick += StyleRibbonRefreshTimer_Tick;

            if (Application != null)
            {
                Application.WindowSelectionChange += Application_WindowSelectionChange;
                Application.WindowActivate += Application_WindowActivate;
            }
        }

        // 插件关闭时执行
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (Application != null)
            {
                try
                {
                    Application.WindowSelectionChange -= Application_WindowSelectionChange;
                    Application.WindowActivate -= Application_WindowActivate;
                }
                catch
                {
                }
            }

            if (styleRibbonRefreshTimer != null)
            {
                try
                {
                    styleRibbonRefreshTimer.Stop();
                    styleRibbonRefreshTimer.Tick -= StyleRibbonRefreshTimer_Tick;
                    styleRibbonRefreshTimer.Dispose();
                }
                catch
                {
                }

                styleRibbonRefreshTimer = null;
            }

            RemoveTaskPane(ref navigationTaskPane, ref navigationPaneControl);
            RemoveTaskPane(ref requirementTrackingTaskPane, ref requirementTrackingConsoleControl);
        }

        // 显示右侧题注面板，并加载题注数据
        internal void ShowCaptionListPane(Word.Document doc, IList<CaptionListEntry> entries)
        {
            // 确保面板已经创建
            EnsureNavigationPane();
            if (navigationPaneControl == null || navigationTaskPane == null)
            {
                return;
            }

            // 获取文档名，给面板设置题注数据
            string docName = doc == null ? "当前文档" : doc.Name;
            navigationPaneControl.SetCaptionEntries(ConvertCaptionEntries(entries), docName);
            navigationPaneControl.SelectTab(NavigationPaneTab.Captions);

            // 显示右侧面板
            navigationTaskPane.Visible = true;
        }

        internal void ShowBookmarkListPane(Word.Document doc, IList<NavigationPaneEntry> entries)
        {
            EnsureNavigationPane();
            if (navigationPaneControl == null || navigationTaskPane == null)
            {
                return;
            }

            string docName = doc == null ? "当前文档" : doc.Name;
            navigationPaneControl.SetBookmarkEntries(entries, docName);
            navigationPaneControl.SelectTab(NavigationPaneTab.Bookmarks);
            navigationTaskPane.Visible = true;
        }

        internal void ShowMarkerListPane(Word.Document doc, IList<NavigationPaneEntry> entries, DocumentMarkerDocumentType documentType)
        {
            EnsureNavigationPane();
            if (navigationPaneControl == null || navigationTaskPane == null)
            {
                return;
            }

            string docName = doc == null ? "当前文档" : doc.Name;
            SetMarkerPaneEntries(docName, entries, documentType);
            navigationPaneControl.SelectTab(NavigationPaneTab.Markers);
            navigationTaskPane.Visible = true;
        }

        internal void ConfigureDocumentBasicInfo(Word.Document doc, IWin32Window owner)
        {
            if (doc == null)
            {
                throw new InvalidOperationException("当前没有可配置的文档。");
            }

            DocumentBasicInfo currentInfo = LoadDocumentBasicInfo(doc);
            using (DocumentBasicInfoForm form = new DocumentBasicInfoForm(currentInfo))
            {
                DialogResult result = owner == null
                    ? form.ShowDialog()
                    : form.ShowDialog(owner);

                if (result != DialogResult.OK)
                {
                    return;
                }

                documentBasicInfoStore.Save(doc, form.BuildInfo());
            }
        }

        internal void ShowRequirementTrackingPane()
        {
            if (!Ribbon1.RequirementTrackingEnabled)
            {
                MessageBox.Show("需求追踪功能暂未开放，当前不可用。", "文档不加班");
                return;
            }

            EnsureRequirementTrackingPane();
            if (requirementTrackingConsoleControl == null || requirementTrackingTaskPane == null)
            {
                return;
            }

            requirementTrackingConsoleControl.RefreshDocumentOptions();
            requirementTrackingTaskPane.Visible = true;
        }

        // 创建统一导航面板（如果还没创建）
        private void EnsureNavigationPane()
        {
            // 创建界面控件，并绑定点击事件
            if (navigationPaneControl == null)
            {
                navigationPaneControl = new NavigationPaneControl();
                navigationPaneControl.CaptionActivated += OnCaptionActivated;
                navigationPaneControl.BookmarkActivated += OnBookmarkActivated;
                navigationPaneControl.MarkerActivated += OnMarkerActivated;
                navigationPaneControl.SelectedTabChanged += OnNavigationTabChanged;
            }

            // 创建Word右侧面板，设置标题、位置、宽度
            if (navigationTaskPane == null)
            {
                navigationTaskPane = CustomTaskPanes.Add(navigationPaneControl, "导航窗格");
                navigationTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                navigationTaskPane.Width = 380;
            }
        }

        private void EnsureRequirementTrackingPane()
        {
            if (requirementTrackingConsoleControl == null)
            {
                requirementTrackingConsoleControl = new RequirementTrackingConsoleControl(() => Application);
            }

            if (requirementTrackingTaskPane == null)
            {
                requirementTrackingTaskPane = CustomTaskPanes.Add(requirementTrackingConsoleControl, "需求追踪控制台");
                requirementTrackingTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                requirementTrackingTaskPane.Width = 874;
            }
        }

        // 点击题注时触发：跳转到对应位置
        private void OnCaptionActivated(int start)
        {
            NavigateToStart(start);
        }

        private void OnBookmarkActivated(int start)
        {
            NavigateToStart(start);
        }

        private void OnMarkerActivated(int start)
        {
            NavigateToStart(start);
        }

        private void OnNavigationTabChanged(NavigationPaneTab tab)
        {
            RefreshNavigationPaneContent(tab);
        }

        private static IList<NavigationPaneEntry> ConvertCaptionEntries(IList<CaptionListEntry> entries)
        {
            if (entries == null)
            {
                return Array.Empty<NavigationPaneEntry>();
            }

            List<NavigationPaneEntry> converted = new List<NavigationPaneEntry>(entries.Count);
            foreach (CaptionListEntry entry in entries)
            {
                if (entry == null)
                {
                    continue;
                }

                converted.Add(new NavigationPaneEntry
                {
                    Start = entry.Start,
                    Text = entry.Text
                });
            }

            return converted;
        }

        private void RefreshNavigationPaneContent(NavigationPaneTab tab)
        {
            Word.Document doc = Application?.ActiveDocument;
            if (navigationPaneControl == null || doc == null)
            {
                return;
            }

            string docName = doc.Name;
            switch (tab)
            {
                case NavigationPaneTab.Bookmarks:
                    navigationPaneControl.SetBookmarkEntries(Ribbon1.CollectBookmarkEntries(doc), docName);
                    break;
                case NavigationPaneTab.Captions:
                    navigationPaneControl.SetCaptionEntries(ConvertCaptionEntries(Ribbon1.CollectCaptionListEntries(doc)), docName);
                    break;
                case NavigationPaneTab.Markers:
                    DocumentMarkerCollectionResult markerResult = DocumentMarkerService.CollectMarkers(doc);
                    SetMarkerPaneEntries(docName, markerResult.Entries, markerResult.DocumentType);
                    break;
            }
        }

        private void SetMarkerPaneEntries(string docName, IList<NavigationPaneEntry> entries, DocumentMarkerDocumentType documentType)
        {
            if (navigationPaneControl == null)
            {
                return;
            }

            if (!DocumentMarkerService.IsMarkerPaneSupportedDocumentType(documentType))
            {
                navigationPaneControl.SetMarkerEntries(
                    new List<NavigationPaneEntry>
                    {
                        new NavigationPaneEntry
                        {
                            Start = -1,
                            Text = "不支持此文档"
                        }
                    },
                    docName,
                    "不支持");
                return;
            }

            navigationPaneControl.SetMarkerEntries(
                entries ?? Array.Empty<NavigationPaneEntry>(),
                docName,
                DocumentMarkerService.GetDocumentTypeDisplayName(documentType));
        }

        // 执行文档跳转：跳转到指定位置
        private void NavigateToStart(int start)
        {
            if (documentHostAdapter == null)
            {
                return;
            }

            documentHostAdapter.NavigateTo(start);
        }

        private void Application_WindowSelectionChange(Word.Selection selection)
        {
            ScheduleStyleRibbonRefresh(selection);
        }

        private void Application_WindowActivate(Word.Document doc, Word.Window wn)
        {
            ScheduleStyleRibbonRefresh(null);
        }

        private DocumentBasicInfo LoadDocumentBasicInfo(Word.Document doc)
        {
            if (documentBasicInfoStore == null)
            {
                documentBasicInfoStore = new DocumentBasicInfoStore();
            }

            return documentBasicInfoStore.Load(doc);
        }

        private void RefreshStyleRibbon()
        {
            try
            {
                Ribbon1.RefreshAllStyleIndicators();
            }
            catch
            {
            }
        }

        private void ScheduleStyleRibbonRefresh(Word.Selection selection)
        {
            if (IsNonCollapsedSelection(selection))
            {
                styleRibbonRefreshTimer?.Stop();
                return;
            }

            if (styleRibbonRefreshTimer == null)
            {
                RefreshStyleRibbon();
                return;
            }

            styleRibbonRefreshTimer.Stop();
            styleRibbonRefreshTimer.Start();
        }

        private void StyleRibbonRefreshTimer_Tick(object sender, EventArgs e)
        {
            styleRibbonRefreshTimer?.Stop();

            Word.Selection selection = null;
            try
            {
                selection = Application?.Selection;
            }
            catch
            {
            }

            if (IsNonCollapsedSelection(selection))
            {
                return;
            }

            RefreshStyleRibbon();
        }

        private static bool IsNonCollapsedSelection(Word.Selection selection)
        {
            try
            {
                Word.Range range = selection?.Range;
                return range != null && range.Start != range.End;
            }
            catch
            {
                return false;
            }
        }

        private void RemoveTaskPane<TControl>(ref CustomTaskPane taskPane, ref TControl control)
            where TControl : class
        {
            if (taskPane != null)
            {
                try
                {
                    CustomTaskPanes.Remove(taskPane);
                }
                catch
                {
                }
            }

            taskPane = null;
            control = null;
        }

        #region VSTO 生成的代码
        /// <summary>
        /// 设计器自动生成的方法 - 不要修改
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
