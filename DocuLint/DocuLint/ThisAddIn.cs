using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Threading = System.Threading;
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
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        // 右侧统一导航面板的界面控件
        private NavigationPaneControl navigationPaneControl;
        // Word右侧自定义任务面板
        private CustomTaskPane navigationTaskPane;
        // 与插件绑定的常用语任务窗格，不随文档切换清空。
        private CommonPhrasesPaneControl commonPhrasesPaneControl;
        private CustomTaskPane commonPhrasesTaskPane;
        private CommonPhraseSuggestionForm commonPhraseSuggestionForm;
        private CommonPhraseHotKeyWindow commonPhraseHotKeyWindow;
        private Threading.SynchronizationContext uiSynchronizationContext;
        // 需求追踪控制台界面控件
        private RequirementTrackingConsoleControl requirementTrackingConsoleControl;
        // 需求追踪独立任务窗格
        private CustomTaskPane requirementTrackingTaskPane;
        private readonly Dictionary<int, RequirementTrackingPaneContext> requirementTrackingPanes =
            new Dictionary<int, RequirementTrackingPaneContext>();
        private readonly Dictionary<int, RequirementExtractionPaneContext> requirementExtractionPanes =
            new Dictionary<int, RequirementExtractionPaneContext>();
        private RequirementExtractionPaneControl requirementExtractionPaneControl;
        private CustomTaskPane requirementExtractionTaskPane;
        private bool isSwitchingWordWindow;
        // 文档检查结果界面控件
        private DocumentCheckResultPaneControl documentCheckResultPaneControl;
        // 文档检查结果任务窗格
        private CustomTaskPane documentCheckResultTaskPane;
        private Word.Document documentCheckResultDocument;
        // 文档跳转辅助工具
        private WordDocumentHostAdapter documentHostAdapter;
        // 文档基本信息存储
        private DocumentBasicInfoStore documentBasicInfoStore;
        private Timer ribbonWarmupTimer;
        // 延迟刷新 Ribbon，避免拖选文字时同步读取 Selection 打断 Word 选区。
        private Timer styleRibbonRefreshTimer;
        private Timer navigationPaneStateRefreshTimer;
        private string lastStyleRefreshDocumentKey;
        private int lastStyleRefreshStart = -1;
        private int lastStyleRefreshEnd = -1;
        // 需求提取批量模式：防抖自动添加。
        private Timer requirementExtractionAutoAddTimer;
        // 需求提取普通模式：等待鼠标拖选结束后再显示自定义“添加”按钮。
        private Timer requirementQuickAddPopupTimer;
        private string requirementExtractionAutoAddDocumentKey;
        private int requirementExtractionAutoAddStart;
        private int requirementExtractionAutoAddEnd;
        private string requirementQuickAddPopupDocumentKey;
        private int requirementQuickAddPopupStart;
        private int requirementQuickAddPopupEnd;
        private Point requirementQuickAddPopupLocation;
        private RequirementQuickAddForm requirementQuickAddPopup;
        private bool wordSelectionFloatiesSuppressed;
        private bool previousShowSelectionFloaties;
        private bool previousShowMenuFloaties;

        // 追踪设置仅在本次 Word 会话中生效，不写入插件配置。
        internal bool PreserveForwardMappingsWhenReverseTracing { get; set; } = true;

        // 插件启动时执行
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            uiSynchronizationContext = Threading.SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
            PluginDataStore.EnsureDefaultFiles();
            Ribbon1.EnableCommonPhraseShortcutHook();
            commonPhraseHotKeyWindow = new CommonPhraseHotKeyWindow(this);
            if (!commonPhraseHotKeyWindow.Register())
            {
                SetPluginShortcutStatus("部分插件快捷键注册失败，请在插件配置中检查冲突。", true);
            }
            // 初始化文档跳转工具
            documentHostAdapter = new WordDocumentHostAdapter(() => Application);
            documentBasicInfoStore = new DocumentBasicInfoStore();
            ribbonWarmupTimer = new Timer { Interval = 350 };
            ribbonWarmupTimer.Tick += RibbonWarmupTimer_Tick;
            ribbonWarmupTimer.Start();
            styleRibbonRefreshTimer = new Timer { Interval = 180 };
            styleRibbonRefreshTimer.Tick += StyleRibbonRefreshTimer_Tick;
            navigationPaneStateRefreshTimer = new Timer { Interval = 120 };
            navigationPaneStateRefreshTimer.Tick += NavigationPaneStateRefreshTimer_Tick;
            requirementExtractionAutoAddTimer = new Timer { Interval = 80 };
            requirementExtractionAutoAddTimer.Tick += RequirementExtractionAutoAddTimer_Tick;
            requirementQuickAddPopupTimer = new Timer { Interval = 20 };
            requirementQuickAddPopupTimer.Tick += RequirementQuickAddPopupTimer_Tick;

            WarmupRibbonDuringStartup();

            if (Application != null)
            {
                Application.WindowSelectionChange += Application_WindowSelectionChange;
                Application.WindowActivate += Application_WindowActivate;
                Application.WindowDeactivate += Application_WindowDeactivate;
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
                    Application.WindowDeactivate -= Application_WindowDeactivate;
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

            if (navigationPaneStateRefreshTimer != null)
            {
                try
                {
                    navigationPaneStateRefreshTimer.Stop();
                    navigationPaneStateRefreshTimer.Tick -= NavigationPaneStateRefreshTimer_Tick;
                    navigationPaneStateRefreshTimer.Dispose();
                }
                catch
                {
                }

                navigationPaneStateRefreshTimer = null;
            }

            if (ribbonWarmupTimer != null)
            {
                try
                {
                    ribbonWarmupTimer.Stop();
                    ribbonWarmupTimer.Tick -= RibbonWarmupTimer_Tick;
                    ribbonWarmupTimer.Dispose();
                }
                catch
                {
                }

                ribbonWarmupTimer = null;
            }

            if (requirementExtractionAutoAddTimer != null)
            {
                try
                {
                    requirementExtractionAutoAddTimer.Stop();
                    requirementExtractionAutoAddTimer.Tick -= RequirementExtractionAutoAddTimer_Tick;
                    requirementExtractionAutoAddTimer.Dispose();
                }
                catch
                {
                }

                requirementExtractionAutoAddTimer = null;
            }

            if (requirementQuickAddPopupTimer != null)
            {
                try
                {
                    requirementQuickAddPopupTimer.Stop();
                    requirementQuickAddPopupTimer.Tick -= RequirementQuickAddPopupTimer_Tick;
                    requirementQuickAddPopupTimer.Dispose();
                }
                catch
                {
                }

                requirementQuickAddPopupTimer = null;
            }

            RemoveTaskPane(ref navigationTaskPane, ref navigationPaneControl);
            RemoveTaskPane(ref commonPhrasesTaskPane, ref commonPhrasesPaneControl);
            RemoveAllRequirementTrackingPanes();
            RemoveAllRequirementExtractionPanes();
            RemoveTaskPane(ref documentCheckResultTaskPane, ref documentCheckResultPaneControl);
            DisposeRequirementQuickAddPopup();
            HideCommonPhraseSuggestion();
            commonPhraseHotKeyWindow?.Dispose();
            commonPhraseHotKeyWindow = null;
            Ribbon1.DisableCommonPhraseShortcutHook();
            RestoreWordSelectionFloaties();
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
            SetCaptionPaneEntries(docName, ConvertCaptionEntries(entries));
            navigationPaneControl.SelectTab(NavigationPaneTab.FigureCaptions);

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

        internal void ShowDocumentCheckResultPane(Word.Document doc, IList<NavigationPaneEntry> entries)
        {
            EnsureDocumentCheckResultPane();
            if (documentCheckResultPaneControl == null || documentCheckResultTaskPane == null)
            {
                return;
            }

            string docName = doc == null ? "当前文档" : doc.Name;
            documentCheckResultPaneControl.SetEntries(entries, docName);
            documentCheckResultDocument = doc;
            documentCheckResultTaskPane.Visible = true;
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
            EnsureRequirementTrackingPane();
            if (requirementTrackingConsoleControl == null || requirementTrackingTaskPane == null)
            {
                return;
            }

            requirementTrackingConsoleControl.LoadCurrentDocumentAsSource();
            requirementTrackingTaskPane.Visible = true;
        }

        internal void ShowRequirementExtractionPane()
        {
            EnsureRequirementExtractionPane();
            if (requirementExtractionPaneControl == null || requirementExtractionTaskPane == null)
            {
                return;
            }

            requirementExtractionTaskPane.Visible = true;
            SuppressWordSelectionFloaties();
        }

        internal void ShowCommonPhrasesPane()
        {
            EnsureCommonPhrasesPane();
            if (commonPhrasesPaneControl == null || commonPhrasesTaskPane == null)
            {
                return;
            }

            commonPhrasesPaneControl.ReloadPhrases();
            commonPhrasesTaskPane.Visible = true;
        }

        internal void RefreshCommonPhrasesPane()
        {
            try
            {
                commonPhrasesPaneControl?.ReloadPhrases();
            }
            catch
            {
            }
        }

        internal void OpenPluginSettings()
        {
            commonPhraseHotKeyWindow?.Suspend();
            try
            {
                using (PluginSettingsForm form = new PluginSettingsForm(
                    Properties.Settings.Default.CommonPhraseShortcut,
                    Properties.Settings.Default.InsertImageCaptionShortcut,
                    Properties.Settings.Default.InsertTableCaptionShortcut))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    Properties.Settings.Default.CommonPhraseShortcut = PluginShortcutService.Normalize(form.CommonPhraseShortcut);
                    Properties.Settings.Default.InsertImageCaptionShortcut = PluginShortcutService.Normalize(form.InsertImageCaptionShortcut);
                    Properties.Settings.Default.InsertTableCaptionShortcut = PluginShortcutService.Normalize(form.InsertTableCaptionShortcut);
                    Properties.Settings.Default.Save();
                }
            }
            finally
            {
                RefreshPluginShortcuts();
            }
        }

        internal void RefreshPluginShortcuts()
        {
            if (commonPhraseHotKeyWindow == null)
            {
                return;
            }

            if (!commonPhraseHotKeyWindow.ReRegister())
            {
                SetPluginShortcutStatus("部分插件快捷键注册失败，请检查是否与其他程序或插件冲突。", true);
            }
            else
            {
                SetPluginShortcutStatus("插件快捷键配置已生效。", false);
            }
        }

        private void SetPluginShortcutStatus(string message, bool warning)
        {
            try
            {
                if (Application != null)
                {
                    Application.StatusBar = message;
                }
            }
            catch
            {
            }
        }

        internal void RequestCommonPhraseSuggestionFromShortcut()
        {
            if (uiSynchronizationContext == null)
            {
                return;
            }

            uiSynchronizationContext.Post(_ => ShowCommonPhraseSuggestions(), null);
        }

        private bool IsWordForeground()
        {
            try
            {
                Word.Window activeWindow = Application?.ActiveWindow;
                return activeWindow != null && new IntPtr(activeWindow.Hwnd) == GetForegroundWindow();
            }
            catch
            {
                return false;
            }
        }

        private void ShowCommonPhraseSuggestions()
        {
            try
            {
                Word.Application application = Application;
                Word.Document document = application?.ActiveDocument;
                Word.Selection selection = application?.Selection;
                if (document == null || selection?.Range == null)
                {
                    return;
                }

                Word.Range selectedRange = selection.Range;
                Word.Range replacementRange = selectedRange.Duplicate;
                string input = selectedRange.Text ?? string.Empty;
                if (replacementRange.Start == replacementRange.End
                    && !TryGetTrailingInput(selection, out replacementRange, out input))
                {
                    return;
                }

                input = input.Trim();
                if (input.Length < 2)
                {
                    return;
                }

                IReadOnlyList<CommonPhraseLibrary.Suggestion> suggestions =
                    CommonPhraseLibrary.FindSimilar(input, 6);
                if (suggestions.Count == 0)
                {
                    application.StatusBar = "未找到相似常用语。";
                    return;
                }

                HideCommonPhraseSuggestion();
                commonPhraseSuggestionForm = new CommonPhraseSuggestionForm(
                    () => Application,
                    GetDocumentKey(document),
                    replacementRange,
                    suggestions);
                commonPhraseSuggestionForm.FormClosed += (_, __) => commonPhraseSuggestionForm = null;
                commonPhraseSuggestionForm.ShowAt(application);
            }
            catch
            {
                HideCommonPhraseSuggestion();
            }
        }

        private void HideCommonPhraseSuggestion()
        {
            try
            {
                if (commonPhraseSuggestionForm != null)
                {
                    commonPhraseSuggestionForm.HideSuggestion();
                    commonPhraseSuggestionForm.Dispose();
                    commonPhraseSuggestionForm = null;
                }
            }
            catch
            {
                commonPhraseSuggestionForm = null;
            }
        }

        private static bool TryGetTrailingInput(
            Word.Selection selection,
            out Word.Range inputRange,
            out string input)
        {
            inputRange = null;
            input = string.Empty;
            try
            {
                Word.Range selectionRange = selection?.Range;
                if (selectionRange == null || selectionRange.Start != selectionRange.End)
                {
                    return false;
                }

                int caret = selectionRange.Start;
                Word.Range probe = selectionRange.Paragraphs[1].Range.Duplicate;
                probe.End = caret;
                if (probe.Start < caret - 240)
                {
                    probe.Start = caret - 240;
                }

                string text = probe.Text ?? string.Empty;
                int tailStart = text.Length;
                while (tailStart > 0)
                {
                    char character = text[tailStart - 1];
                    if (character == '\r' || character == '\a' || character == '。' || character == '！'
                        || character == '？' || character == '；' || character == ':' || character == '：'
                        || character == ',' || character == '，' || character == '.' || character == '!'
                        || character == '?')
                    {
                        break;
                    }

                    tailStart--;
                }

                while (tailStart < text.Length && char.IsWhiteSpace(text[tailStart]))
                {
                    tailStart++;
                }

                string candidate = text.Substring(tailStart).TrimEnd();
                if (candidate.Length < 2)
                {
                    return false;
                }

                int leadingWhitespace = text.Substring(tailStart).Length
                    - text.Substring(tailStart).TrimStart().Length;
                inputRange = probe.Duplicate;
                inputRange.Start = probe.Start + tailStart + leadingWhitespace;
                inputRange.End = caret;
                input = candidate;
                return true;
            }
            catch
            {
                return false;
            }
        }

        internal void OpenRequirementExtractionSettings()
        {
            EnsureRequirementExtractionPane();
            if (requirementExtractionPaneControl == null)
            {
                return;
            }

            requirementExtractionPaneControl.OpenExtractionSettings();
        }

        internal void AddSelectedTextToRequirementExtraction()
        {
            EnsureRequirementExtractionPane();
            if (requirementExtractionPaneControl == null || requirementExtractionTaskPane == null)
            {
                return;
            }

            if (!requirementExtractionPaneControl.ExtractionEnabled ||
                requirementExtractionPaneControl.BatchExtractionEnabled)
            {
                return;
            }

            requirementExtractionTaskPane.Visible = true;
            HideRequirementQuickAddPopup();
            requirementExtractionPaneControl.AddSelectionAsRequirement();
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
                navigationTaskPane.Width = 460;
            }
        }

        private void EnsureRequirementTrackingPane()
        {
            Word.Window activeWindow = null;
            Word.Document activeDocument = null;
            try
            {
                activeWindow = Application?.ActiveWindow;
                activeDocument = Application?.ActiveDocument;
            }
            catch
            {
            }

            int windowKey = GetWordWindowKey(activeWindow);
            if (windowKey == 0)
            {
                SetActiveRequirementTrackingPane(null);
                return;
            }

            string documentKey = GetDocumentKey(activeDocument);
            if (requirementTrackingPanes.TryGetValue(
                windowKey,
                out RequirementTrackingPaneContext existingContext))
            {
                if (string.Equals(
                    existingContext.DocumentKey,
                    documentKey,
                    StringComparison.OrdinalIgnoreCase))
                {
                    SetActiveRequirementTrackingPane(existingContext);
                    return;
                }

                RemoveRequirementTrackingPane(windowKey, existingContext);
            }

            RequirementTrackingConsoleControl control =
                new RequirementTrackingConsoleControl(() => Application);
            CustomTaskPane taskPane = CustomTaskPanes.Add(
                control,
                "需求追踪控制台",
                activeWindow);
            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            // Request the widest pane; Word clamps this value to the maximum
            // docked width supported by the current document window.
            taskPane.Width = 2400;

            RequirementTrackingPaneContext context = new RequirementTrackingPaneContext
            {
                WindowKey = windowKey,
                DocumentKey = documentKey,
                Control = control,
                TaskPane = taskPane
            };
            requirementTrackingPanes[windowKey] = context;
            SetActiveRequirementTrackingPane(context);
        }

        private void SetActiveRequirementTrackingPane(RequirementTrackingPaneContext context)
        {
            requirementTrackingConsoleControl = context?.Control;
            requirementTrackingTaskPane = context?.TaskPane;
        }

        private void SelectRequirementTrackingPaneForWindow(
            Word.Window window,
            Word.Document document)
        {
            int windowKey = GetWordWindowKey(window);
            string documentKey = GetDocumentKey(document);
            if (windowKey != 0 &&
                requirementTrackingPanes.TryGetValue(
                    windowKey,
                    out RequirementTrackingPaneContext context))
            {
                if (string.Equals(
                    context.DocumentKey,
                    documentKey,
                    StringComparison.OrdinalIgnoreCase))
                {
                    SetActiveRequirementTrackingPane(context);
                    return;
                }

                RemoveRequirementTrackingPane(windowKey, context);
            }

            SetActiveRequirementTrackingPane(null);
        }

        private void EnsureRequirementExtractionPane()
        {
            Word.Window activeWindow = null;
            Word.Document activeDocument = null;
            try
            {
                activeWindow = Application?.ActiveWindow;
                activeDocument = Application?.ActiveDocument;
            }
            catch
            {
            }

            int windowKey = GetWordWindowKey(activeWindow);
            if (windowKey == 0)
            {
                requirementExtractionPaneControl = null;
                requirementExtractionTaskPane = null;
                return;
            }

            string documentKey = GetDocumentKey(activeDocument);
            if (requirementExtractionPanes.TryGetValue(
                windowKey,
                out RequirementExtractionPaneContext existingContext))
            {
                if (string.Equals(
                    existingContext.DocumentKey,
                    documentKey,
                    StringComparison.OrdinalIgnoreCase))
                {
                    SetActiveRequirementExtractionPane(existingContext);
                    return;
                }

                RemoveRequirementExtractionPane(windowKey, existingContext);
            }

            RequirementExtractionPaneControl control =
                new RequirementExtractionPaneControl(() => Application);
            control.LoadSavedRequirementsFromCurrentDocument();
            control.RequirementActivated += NavigateToStart;
            control.BatchExtractionModeChanged += enabled =>
            {
                if (enabled)
                {
                    HideRequirementQuickAddPopup();
                }
            };
            control.ExtractionModeChanged += enabled =>
            {
                if (!enabled)
                {
                    StopRequirementQuickAddPopup();
                    HideRequirementQuickAddPopup();
                }
            };

            CustomTaskPane taskPane = CustomTaskPanes.Add(control, "需求提取", activeWindow);
            taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            taskPane.Width = GetInitialTaskPaneWidth(0.34, 540, 1100);

            RequirementExtractionPaneContext context = new RequirementExtractionPaneContext
            {
                WindowKey = windowKey,
                DocumentKey = documentKey,
                Control = control,
                TaskPane = taskPane
            };
            taskPane.VisibleChanged += (_, __) =>
                HandleRequirementExtractionPaneVisibleChanged(context);
            requirementExtractionPanes[windowKey] = context;
            SetActiveRequirementExtractionPane(context);
        }

        private static int GetInitialTaskPaneWidth(double screenRatio, int minimum, int maximum)
        {
            try
            {
                int screenWidth = Screen.PrimaryScreen?.WorkingArea.Width ?? maximum;
                int width = (int)Math.Round(screenWidth * screenRatio);
                return Math.Max(minimum, Math.Min(maximum, width));
            }
            catch
            {
                return maximum;
            }
        }

        private void HandleRequirementExtractionPaneVisibleChanged(
            RequirementExtractionPaneContext context)
        {
            if (context == null || context.Removing)
            {
                return;
            }

            bool visible;
            try
            {
                visible = context.TaskPane != null && context.TaskPane.Visible;
            }
            catch
            {
                visible = false;
            }

            if (!visible)
            {
                string activeDocumentKey = GetDocumentKey(Application?.ActiveDocument);
                bool sameDocument = !string.IsNullOrWhiteSpace(context.DocumentKey) &&
                                    string.Equals(
                                        context.DocumentKey,
                                        activeDocumentKey,
                                        StringComparison.OrdinalIgnoreCase);
                if (!isSwitchingWordWindow && sameDocument)
                {
                    context.Control?.ConfirmSaveBeforeClosing();
                }

                context.Control?.ResetOneTimeExtractionSettings();

                StopRequirementQuickAddPopup();
                HideRequirementQuickAddPopup();
                RestoreWordSelectionFloaties();
                return;
            }

            if (ReferenceEquals(context.TaskPane, requirementExtractionTaskPane))
            {
                SuppressWordSelectionFloaties();
            }
        }

        private void SetActiveRequirementExtractionPane(RequirementExtractionPaneContext context)
        {
            requirementExtractionPaneControl = context?.Control;
            requirementExtractionTaskPane = context?.TaskPane;
        }

        private void SelectRequirementExtractionPaneForWindow(
            Word.Window window,
            Word.Document document)
        {
            int windowKey = GetWordWindowKey(window);
            string documentKey = GetDocumentKey(document);
            if (windowKey != 0 &&
                requirementExtractionPanes.TryGetValue(
                    windowKey,
                    out RequirementExtractionPaneContext context))
            {
                if (string.Equals(
                    context.DocumentKey,
                    documentKey,
                    StringComparison.OrdinalIgnoreCase))
                {
                    SetActiveRequirementExtractionPane(context);
                    return;
                }

                RemoveRequirementExtractionPane(windowKey, context);
            }

            SetActiveRequirementExtractionPane(null);
        }

        private static int GetWordWindowKey(Word.Window window)
        {
            try
            {
                return window?.Hwnd ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private void EnsureDocumentCheckResultPane()
        {
            if (documentCheckResultPaneControl == null)
            {
                documentCheckResultPaneControl = new DocumentCheckResultPaneControl();
                documentCheckResultPaneControl.IssueActivated += OnDocumentCheckIssueActivated;
            }

            if (documentCheckResultTaskPane == null)
            {
                documentCheckResultTaskPane = CustomTaskPanes.Add(documentCheckResultPaneControl, "文档检查结果");
                documentCheckResultTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                documentCheckResultTaskPane.Width = 460;
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

        private void OnDocumentCheckIssueActivated(int start)
        {
            try
            {
                documentCheckResultDocument?.Activate();
            }
            catch
            {
            }

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
                case NavigationPaneTab.FigureCaptions:
                case NavigationPaneTab.TableCaptions:
                    SetCaptionPaneEntries(docName, ConvertCaptionEntries(Ribbon1.CollectCaptionListEntries(doc)));
                    break;
                case NavigationPaneTab.Markers:
                    DocumentMarkerCollectionResult markerResult = DocumentMarkerService.CollectMarkers(doc);
                    SetMarkerPaneEntries(docName, markerResult.Entries, markerResult.DocumentType);
                    break;
            }
        }

        private void SetCaptionPaneEntries(string docName, IList<NavigationPaneEntry> entries)
        {
            if (navigationPaneControl == null)
            {
                return;
            }

            List<NavigationPaneEntry> figures = new List<NavigationPaneEntry>();
            List<NavigationPaneEntry> tables = new List<NavigationPaneEntry>();
            if (entries != null)
            {
                foreach (NavigationPaneEntry entry in entries)
                {
                    if (entry == null)
                    {
                        continue;
                    }

                    if (IsTableCaption(entry.Text))
                    {
                        tables.Add(entry);
                    }
                    else if (IsFigureCaption(entry.Text))
                    {
                        figures.Add(entry);
                }
            }
        }

            navigationPaneControl.SetFigureCaptionEntries(figures, docName);
            navigationPaneControl.SetTableCaptionEntries(tables, docName);
        }

        private static bool IsFigureCaption(string text)
        {
            string value = (text ?? string.Empty).TrimStart();
            return value.StartsWith("图", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("Figure", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTableCaption(string text)
        {
            string value = (text ?? string.Empty).TrimStart();
            return value.StartsWith("表", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith("Table", StringComparison.OrdinalIgnoreCase);
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
            Ribbon1.HandleStyleBrushSelectionChange(selection);
            RefreshNavigationPaneState();
            ScheduleStyleRibbonRefresh(selection);
            ScheduleRequirementExtractionAutoAdd(selection);
            ScheduleRequirementQuickAddPopup(selection);
        }

        private void Application_WindowActivate(Word.Document doc, Word.Window wn)
        {
            HideCommonPhraseSuggestion();
            Ribbon1.HandleStyleBrushWindowActivated(doc);
            isSwitchingWordWindow = false;
            SelectRequirementTrackingPaneForWindow(wn, doc);
            SelectRequirementExtractionPaneForWindow(wn, doc);
            StopStyleRibbonRefresh();
            StopRequirementExtractionAutoAdd();
            StopRequirementQuickAddPopup();
            HideRequirementQuickAddPopup();
            // 文档切换时光标可能没有发生移动，主动刷新样式下拉框，
            // 避免已保存的常用样式必须再次打开面板并点击保存才显示。
            try
            {
                Ribbon1.RefreshAllStyleIndicators();
            }
            catch
            {
            }
            RefreshNavigationPaneState();
            ScheduleNavigationPaneStateRefresh();
            if (IsRequirementExtractionPaneVisible())
            {
                SuppressWordSelectionFloaties();
            }
            else
            {
                RestoreWordSelectionFloaties();
            }
        }

        private void Application_WindowDeactivate(Word.Document doc, Word.Window wn)
        {
            isSwitchingWordWindow = true;
            HideCommonPhraseSuggestion();
            StopRequirementQuickAddPopup();
            HideRequirementQuickAddPopup();
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

        private void EnsureCommonPhrasesPane()
        {
            if (commonPhrasesPaneControl == null)
            {
                commonPhrasesPaneControl = new CommonPhrasesPaneControl(() => Application);
            }

            if (commonPhrasesTaskPane == null)
            {
                commonPhrasesTaskPane = CustomTaskPanes.Add(commonPhrasesPaneControl, "常用语");
                commonPhrasesTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                commonPhrasesTaskPane.Width = 360;
            }
        }

        internal void ScheduleNavigationPaneStateRefresh()
        {
            try
            {
                navigationPaneStateRefreshTimer?.Stop();
                navigationPaneStateRefreshTimer?.Start();
            }
            catch
            {
            }
        }

        private void NavigationPaneStateRefreshTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                navigationPaneStateRefreshTimer?.Stop();
                RefreshNavigationPaneState();
            }
            catch
            {
            }
        }

        private static void RefreshNavigationPaneState()
        {
            try
            {
                Ribbon1.RefreshNavigationPaneIndicators();
            }
            catch
            {
            }
        }

        private void RibbonWarmupTimer_Tick(object sender, EventArgs e)
        {
            ribbonWarmupTimer?.Stop();
            WarmupRibbonDuringStartup();
        }

        private static void WarmupRibbonDuringStartup()
        {
            try
            {
                _ = Globals.Ribbons.Ribbon1;
            }
            catch
            {
            }
        }

        private void ScheduleStyleRibbonRefresh(Word.Selection selection)
        {
            if (IsNonCollapsedSelection(selection))
            {
                StopStyleRibbonRefresh();
                return;
            }

            if (styleRibbonRefreshTimer == null)
            {
                return;
            }

            styleRibbonRefreshTimer.Stop();
            styleRibbonRefreshTimer.Start();
        }

        private void StopStyleRibbonRefresh()
        {
            try
            {
                styleRibbonRefreshTimer?.Stop();
            }
            catch
            {
            }
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

            if (!HasStyleRefreshTargetChanged(selection))
            {
                return;
            }

            RefreshStyleRibbon();
        }

        private bool HasStyleRefreshTargetChanged(Word.Selection selection)
        {
            try
            {
                Word.Range range = selection?.Range;
                Word.Document document = selection?.Document;
                if (range == null || document == null)
                {
                    return false;
                }

                string documentKey = GetDocumentKey(document);
                if (string.Equals(lastStyleRefreshDocumentKey, documentKey, StringComparison.OrdinalIgnoreCase) &&
                    lastStyleRefreshStart == range.Start &&
                    lastStyleRefreshEnd == range.End)
                {
                    return false;
                }

                lastStyleRefreshDocumentKey = documentKey;
                lastStyleRefreshStart = range.Start;
                lastStyleRefreshEnd = range.End;
                return true;
            }
            catch
            {
                return true;
            }
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

        private void ScheduleRequirementQuickAddPopup(Word.Selection selection)
        {
            if (!IsRequirementExtractionPaneVisible() ||
                requirementExtractionPaneControl == null ||
                !requirementExtractionPaneControl.ExtractionEnabled ||
                requirementExtractionPaneControl.BatchExtractionEnabled ||
                !IsNonCollapsedSelection(selection))
            {
                StopRequirementQuickAddPopup();
                HideRequirementQuickAddPopup();
                return;
            }

            SuppressWordSelectionFloaties();
            if (!TryCaptureRequirementQuickAddSelection(selection))
            {
                StopRequirementQuickAddPopup();
                HideRequirementQuickAddPopup();
                return;
            }

            if (Control.MouseButtons == MouseButtons.None)
            {
                ShowRequirementQuickAddPopupIfSelectionStable();
                return;
            }

            requirementQuickAddPopupTimer?.Stop();
            requirementQuickAddPopupTimer?.Start();
        }

        private void RequirementQuickAddPopupTimer_Tick(object sender, EventArgs e)
        {
            if (Control.MouseButtons != MouseButtons.None)
            {
                requirementQuickAddPopupTimer?.Stop();
                requirementQuickAddPopupTimer?.Start();
                return;
            }

            ShowRequirementQuickAddPopupIfSelectionStable();
        }

        private void ShowRequirementQuickAddPopupIfSelectionStable()
        {
            StopRequirementQuickAddPopupTimerOnly();
            Word.Selection selection = null;
            try
            {
                selection = Application?.Selection;
            }
            catch
            {
            }

            if (!IsRequirementExtractionPaneVisible() ||
                requirementExtractionPaneControl == null ||
                !requirementExtractionPaneControl.ExtractionEnabled ||
                requirementExtractionPaneControl.BatchExtractionEnabled ||
                !IsNonCollapsedSelection(selection) ||
                !SelectionMatchesQuickAddSelection(selection))
            {
                HideRequirementQuickAddPopup();
                return;
            }

            SuppressWordSelectionFloaties();
            EnsureRequirementQuickAddPopup();
            requirementQuickAddPopup.ShowNear(requirementQuickAddPopupLocation, GetActiveWordWindowOwner());
        }

        private bool IsRequirementExtractionPaneVisible()
        {
            try
            {
                return requirementExtractionTaskPane != null && requirementExtractionTaskPane.Visible;
            }
            catch
            {
                return false;
            }
        }

        private void ScheduleRequirementExtractionAutoAdd(Word.Selection selection)
        {
            if (!IsRequirementExtractionPaneVisible() ||
                requirementExtractionPaneControl == null ||
                !requirementExtractionPaneControl.ExtractionEnabled ||
                !requirementExtractionPaneControl.BatchExtractionEnabled ||
                !IsNonCollapsedSelection(selection))
            {
                StopRequirementExtractionAutoAdd();
                return;
            }

            if (Control.MouseButtons == MouseButtons.None)
            {
                StopRequirementExtractionAutoAdd();
                try
                {
                    if (requirementExtractionPaneControl.TryAutoAddSelection())
                    {
                        HideRequirementQuickAddPopup();
                    }
                }
                catch
                {
                }

                return;
            }

            if (!TryCaptureRequirementExtractionSelection(selection))
            {
                StopRequirementExtractionAutoAdd();
                return;
            }

            requirementExtractionAutoAddTimer?.Stop();
            requirementExtractionAutoAddTimer?.Start();
        }

        private void RequirementExtractionAutoAddTimer_Tick(object sender, EventArgs e)
        {
            StopRequirementExtractionAutoAdd();

            Word.Selection selection = null;
            try
            {
                selection = Application?.Selection;
            }
            catch
            {
            }

            if (!IsRequirementExtractionPaneVisible() ||
                requirementExtractionPaneControl == null ||
                !requirementExtractionPaneControl.BatchExtractionEnabled ||
                !IsNonCollapsedSelection(selection) ||
                !SelectionMatchesCapturedSelection(selection))
            {
                return;
            }

            try
            {
                if (requirementExtractionPaneControl.TryAutoAddSelection())
                {
                    HideRequirementQuickAddPopup();
                }
            }
            catch
            {
            }
        }

        private void StopRequirementExtractionAutoAdd()
        {
            try
            {
                requirementExtractionAutoAddTimer?.Stop();
            }
            catch
            {
            }

            requirementExtractionAutoAddDocumentKey = null;
            requirementExtractionAutoAddStart = 0;
            requirementExtractionAutoAddEnd = 0;
        }

        private void StopRequirementQuickAddPopup()
        {
            StopRequirementQuickAddPopupTimerOnly();
            requirementQuickAddPopupDocumentKey = null;
            requirementQuickAddPopupStart = 0;
            requirementQuickAddPopupEnd = 0;
        }

        private void StopRequirementQuickAddPopupTimerOnly()
        {
            try
            {
                requirementQuickAddPopupTimer?.Stop();
            }
            catch
            {
            }
        }

        private bool TryCaptureRequirementExtractionSelection(Word.Selection selection)
        {
            try
            {
                Word.Range range = selection?.Range;
                Word.Document doc = selection?.Document;
                if (range == null || doc == null)
                {
                    return false;
                }

                requirementExtractionAutoAddDocumentKey = GetDocumentKey(doc);
                requirementExtractionAutoAddStart = range.Start;
                requirementExtractionAutoAddEnd = range.End;
                return requirementExtractionAutoAddEnd > requirementExtractionAutoAddStart;
            }
            catch
            {
                return false;
            }
        }

        private bool SelectionMatchesCapturedSelection(Word.Selection selection)
        {
            try
            {
                Word.Range range = selection?.Range;
                Word.Document doc = selection?.Document;
                if (range == null || doc == null)
                {
                    return false;
                }

                string currentKey = GetDocumentKey(doc);
                return string.Equals(currentKey, requirementExtractionAutoAddDocumentKey, StringComparison.OrdinalIgnoreCase) &&
                       range.Start == requirementExtractionAutoAddStart &&
                       range.End == requirementExtractionAutoAddEnd;
            }
            catch
            {
                return false;
            }
        }

        private bool TryCaptureRequirementQuickAddSelection(Word.Selection selection)
        {
            try
            {
                Word.Range range = selection?.Range;
                Word.Document doc = selection?.Document;
                if (range == null || doc == null || range.End <= range.Start)
                {
                    return false;
                }

                requirementQuickAddPopupDocumentKey = GetDocumentKey(doc);
                requirementQuickAddPopupStart = range.Start;
                requirementQuickAddPopupEnd = range.End;
                requirementQuickAddPopupLocation = Cursor.Position;
                requirementQuickAddPopupLocation.Offset(12, 16);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool SelectionMatchesQuickAddSelection(Word.Selection selection)
        {
            try
            {
                Word.Range range = selection?.Range;
                Word.Document doc = selection?.Document;
                if (range == null || doc == null)
                {
                    return false;
                }

                string currentKey = GetDocumentKey(doc);
                return string.Equals(currentKey, requirementQuickAddPopupDocumentKey, StringComparison.OrdinalIgnoreCase) &&
                       range.Start == requirementQuickAddPopupStart &&
                       range.End == requirementQuickAddPopupEnd;
            }
            catch
            {
                return false;
            }
        }

        private static string GetDocumentKey(Word.Document doc)
        {
            try
            {
                return string.IsNullOrWhiteSpace(doc?.FullName) ? doc?.Name ?? string.Empty : doc.FullName;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void EnsureRequirementQuickAddPopup()
        {
            if (requirementQuickAddPopup != null && !requirementQuickAddPopup.IsDisposed)
            {
                return;
            }

            requirementQuickAddPopup = new RequirementQuickAddForm(AddSelectedTextToRequirementExtraction);
        }

        private IWin32Window GetActiveWordWindowOwner()
        {
            try
            {
                int hwnd = Application?.ActiveWindow?.Hwnd ?? 0;
                return hwnd == 0 ? null : new WindowHandle(new IntPtr(hwnd));
            }
            catch
            {
                return null;
            }
        }

        private void HideRequirementQuickAddPopup()
        {
            try
            {
                requirementQuickAddPopup?.Hide();
            }
            catch
            {
            }
        }

        private void DisposeRequirementQuickAddPopup()
        {
            try
            {
                requirementQuickAddPopup?.Dispose();
            }
            catch
            {
            }

            requirementQuickAddPopup = null;
        }

        private void SuppressWordSelectionFloaties()
        {
            try
            {
                if (Application?.Options == null)
                {
                    return;
                }

                if (!wordSelectionFloatiesSuppressed)
                {
                    previousShowSelectionFloaties = Application.Options.ShowSelectionFloaties;
                    previousShowMenuFloaties = Application.Options.ShowMenuFloaties;
                    wordSelectionFloatiesSuppressed = true;
                }

                Application.Options.ShowSelectionFloaties = false;
                Application.Options.ShowMenuFloaties = false;
            }
            catch
            {
            }
        }

        private void RestoreWordSelectionFloaties()
        {
            try
            {
                if (!wordSelectionFloatiesSuppressed || Application?.Options == null)
                {
                    return;
                }

                Application.Options.ShowSelectionFloaties = previousShowSelectionFloaties;
                Application.Options.ShowMenuFloaties = previousShowMenuFloaties;
                wordSelectionFloatiesSuppressed = false;
            }
            catch
            {
            }
        }

        private void RemoveRequirementExtractionPane(
            int windowKey,
            RequirementExtractionPaneContext context)
        {
            if (context == null)
            {
                requirementExtractionPanes.Remove(windowKey);
                return;
            }

            context.Removing = true;
            if (ReferenceEquals(requirementExtractionTaskPane, context.TaskPane))
            {
                SetActiveRequirementExtractionPane(null);
            }

            try
            {
                if (context.TaskPane != null)
                {
                    CustomTaskPanes.Remove(context.TaskPane);
                }
            }
            catch
            {
            }

            try
            {
                context.Control?.Dispose();
            }
            catch
            {
            }

            requirementExtractionPanes.Remove(windowKey);
        }

        private void RemoveRequirementTrackingPane(
            int windowKey,
            RequirementTrackingPaneContext context)
        {
            if (context == null)
            {
                requirementTrackingPanes.Remove(windowKey);
                return;
            }

            if (ReferenceEquals(requirementTrackingTaskPane, context.TaskPane))
            {
                SetActiveRequirementTrackingPane(null);
            }

            try
            {
                if (context.TaskPane != null)
                {
                    CustomTaskPanes.Remove(context.TaskPane);
                }
            }
            catch
            {
            }

            try
            {
                context.Control?.Dispose();
            }
            catch
            {
            }

            requirementTrackingPanes.Remove(windowKey);
        }

        private void RemoveAllRequirementTrackingPanes()
        {
            foreach (KeyValuePair<int, RequirementTrackingPaneContext> item in
                new List<KeyValuePair<int, RequirementTrackingPaneContext>>(
                    requirementTrackingPanes))
            {
                RemoveRequirementTrackingPane(item.Key, item.Value);
            }

            requirementTrackingPanes.Clear();
            SetActiveRequirementTrackingPane(null);
        }

        private void RemoveAllRequirementExtractionPanes()
        {
            foreach (KeyValuePair<int, RequirementExtractionPaneContext> item in
                new List<KeyValuePair<int, RequirementExtractionPaneContext>>(
                    requirementExtractionPanes))
            {
                RemoveRequirementExtractionPane(item.Key, item.Value);
            }

            requirementExtractionPanes.Clear();
            SetActiveRequirementExtractionPane(null);
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

        private sealed class RequirementExtractionPaneContext
        {
            internal int WindowKey { get; set; }
            internal string DocumentKey { get; set; }
            internal RequirementExtractionPaneControl Control { get; set; }
            internal CustomTaskPane TaskPane { get; set; }
            internal bool Removing { get; set; }
        }

        private sealed class RequirementTrackingPaneContext
        {
            internal int WindowKey { get; set; }
            internal string DocumentKey { get; set; }
            internal RequirementTrackingConsoleControl Control { get; set; }
            internal CustomTaskPane TaskPane { get; set; }
        }

        private sealed class RequirementQuickAddForm : Form
        {
            private const int WsExNoActivate = 0x08000000;
            private readonly System.Action addAction;

            internal RequirementQuickAddForm(System.Action addAction)
            {
                this.addAction = addAction;
                AutoScaleMode = AutoScaleMode.None;
                BackColor = Color.White;
                ClientSize = new Size(74, 34);
                FormBorderStyle = FormBorderStyle.None;
                ShowInTaskbar = false;
                StartPosition = FormStartPosition.Manual;
                TopMost = false;

                Button button = new Button
                {
                    Dock = DockStyle.Fill,
                    Text = "提取",
                    Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                    BackColor = Color.FromArgb(47, 111, 237),
                    ForeColor = Color.White,
                    FlatStyle = FlatStyle.Flat
                };
                button.FlatAppearance.BorderSize = 0;
                button.Click += (_, __) =>
                {
                    Hide();
                    this.addAction?.Invoke();
                };
                Controls.Add(button);
            }

            protected override bool ShowWithoutActivation => true;

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams cp = base.CreateParams;
                    cp.ExStyle |= WsExNoActivate;
                    return cp;
                }
            }

            internal void ShowNear(Point location, IWin32Window owner)
            {
                Location = location;
                if (!Visible)
                {
                    if (owner == null)
                    {
                        Show();
                    }
                    else
                    {
                        Show(owner);
                    }
                }
            }
        }

        private sealed class CommonPhraseHotKeyWindow : NativeWindow, IDisposable
        {
            private const int WmHotKey = 0x0312;
            private const int CommonPhraseHotKeyId = 0x4443;
            private const int ImageCaptionHotKeyId = 0x4444;
            private const int TableCaptionHotKeyId = 0x4445;
            private readonly ThisAddIn owner;
            private readonly HashSet<int> registeredIds = new HashSet<int>();
            private bool registered;

            internal CommonPhraseHotKeyWindow(ThisAddIn owner)
            {
                this.owner = owner;
            }

            internal bool Register()
            {
                if (!registered)
                {
                    CreateHandle(new CreateParams());
                    registered = true;
                }

                return ReRegister();
            }

            internal bool ReRegister()
            {
                if (Handle == IntPtr.Zero)
                {
                    return false;
                }

                foreach (int id in registeredIds.ToArray())
                {
                    UnregisterHotKey(Handle, id);
                }

                registeredIds.Clear();
                bool success = true;
                HashSet<string> configured = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                success &= TryRegisterShortcut(
                    CommonPhraseHotKeyId,
                    Properties.Settings.Default.CommonPhraseShortcut,
                    configured);
                success &= TryRegisterShortcut(
                    ImageCaptionHotKeyId,
                    Properties.Settings.Default.InsertImageCaptionShortcut,
                    configured);
                success &= TryRegisterShortcut(
                    TableCaptionHotKeyId,
                    Properties.Settings.Default.InsertTableCaptionShortcut,
                    configured);
                return success;
            }

            internal void Suspend()
            {
                if (Handle == IntPtr.Zero)
                {
                    return;
                }

                foreach (int id in registeredIds.ToArray())
                {
                    UnregisterHotKey(Handle, id);
                }

                registeredIds.Clear();
            }

            protected override void WndProc(ref Message message)
            {
                if (message.Msg == WmHotKey && registeredIds.Contains(message.WParam.ToInt32()))
                {
                    if (owner?.IsWordForeground() == true)
                    {
                        switch (message.WParam.ToInt32())
                        {
                            case CommonPhraseHotKeyId:
                                owner.RequestCommonPhraseSuggestionFromShortcut();
                                break;
                            case ImageCaptionHotKeyId:
                                Ribbon1.ExecuteInsertImageCaption();
                                break;
                            case TableCaptionHotKeyId:
                                Ribbon1.ExecuteInsertTableCaption();
                                break;
                        }
                    }
                }

                base.WndProc(ref message);
            }

            public void Dispose()
            {
                if (registered && Handle != IntPtr.Zero)
                {
                    foreach (int id in registeredIds.ToArray())
                    {
                        UnregisterHotKey(Handle, id);
                    }

                    registeredIds.Clear();
                    registered = false;
                }

                if (Handle != IntPtr.Zero)
                {
                    ReleaseHandle();
                }
            }

            private bool TryRegisterShortcut(int id, string value, HashSet<string> configured)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    return true;
                }

                string normalized = PluginShortcutService.Normalize(value);
                if (string.IsNullOrWhiteSpace(normalized)
                    || !configured.Add(normalized)
                    || !PluginShortcutService.TryParse(normalized, out PluginShortcutService.ShortcutDefinition definition))
                {
                    return false;
                }

                bool result = RegisterHotKey(
                    Handle,
                    id,
                    definition.Modifiers | PluginShortcutService.ModNoRepeat,
                    definition.VirtualKey);
                if (result)
                {
                    registeredIds.Add(id);
                }

                return result;
            }

            [System.Runtime.InteropServices.DllImport("user32.dll")]
            private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint modifiers, uint virtualKey);

            [System.Runtime.InteropServices.DllImport("user32.dll")]
            private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        }

        private sealed class WindowHandle : IWin32Window
        {
            internal WindowHandle(IntPtr handle)
            {
                Handle = handle;
            }

            public IntPtr Handle { get; }
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
