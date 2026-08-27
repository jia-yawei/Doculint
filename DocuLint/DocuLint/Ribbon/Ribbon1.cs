using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    // Word 工具栏（Ribbon）功能类：快速应用文档样式
    public partial class Ribbon1
    {
        private static readonly List<Ribbon1> LoadedInstances = new List<Ribbon1>();
        private bool updatingOutlineLevel;
        private bool updatingCommonStyles;
        private string commonStylesDocumentKey = string.Empty;
        private static readonly Dictionary<string, List<string>> headingStyleLibraryCache =
            new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        private static Dictionary<int, StyleDefinitionRequest> styleDefinitions;
        private static OutlineNumberPattern outlineNumberPattern = OutlineNumberPattern.Decimal;
        private static int outlineNumberTextSpacing = 1;
        private static volatile bool operationCancelRequested;
        private static volatile bool styleBrushActive;
        private static volatile bool styleBrushPersistent;
        private static volatile bool styleBrushApplying;
        private static long styleBrushLastClickTicks;
        private static string styleBrushSourceDocumentKey = string.Empty;
        private static int styleBrushSourceStart = -1;
        private static int styleBrushSourceEnd = -1;
        private static IntPtr keyboardHookHandle = IntPtr.Zero;
        private static LowLevelKeyboardProc keyboardHookProc;
        internal static bool RequirementTrackingEnabled => true;

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        private const int SwRestore = 9;
        private const int WhKeyboardLl = 13;
        private const int WmKeyDown = 0x0100;
        private const int WmSysKeyDown = 0x0104;
        private const int VkEscape = 0x1B;
        private const int VkA = 0x41;
        private const int VkControl = 0x11;
        private const int VkMenu = 0x12;

        private sealed class WordWindowItem
        {
            internal IntPtr Handle { get; set; }

            internal string Title { get; set; }
        }

        private sealed class WordPerformanceScope : IDisposable
        {
            private readonly Word.Application app;
            private readonly object previousScreenUpdating;
            private readonly object previousDisplayAlerts;
            private readonly object previousPagination;
            private readonly object previousCheckSpellingAsYouType;
            private readonly object previousCheckGrammarAsYouType;

            public WordPerformanceScope(Word.Application appInstance)
            {
                app = appInstance;
                if (app == null)
                {
                    return;
                }

                previousScreenUpdating = TryGetComProperty(app, "ScreenUpdating");
                previousDisplayAlerts = TryGetComProperty(app, "DisplayAlerts");
                object options = TryGetComProperty(app, "Options");
                previousPagination = TryGetComProperty(options, "Pagination");
                previousCheckSpellingAsYouType = TryGetComProperty(options, "CheckSpellingAsYouType");
                previousCheckGrammarAsYouType = TryGetComProperty(options, "CheckGrammarAsYouType");

                TrySetComProperty(app, "ScreenUpdating", false);
                TrySetComProperty(app, "DisplayAlerts", Word.WdAlertLevel.wdAlertsNone);
                TrySetComProperty(options, "Pagination", false);
                TrySetComProperty(options, "CheckSpellingAsYouType", false);
                TrySetComProperty(options, "CheckGrammarAsYouType", false);
            }

            public void Dispose()
            {
                if (app == null)
                {
                    return;
                }

                object options = TryGetComProperty(app, "Options");
                TryRestoreComProperty(options, "CheckGrammarAsYouType", previousCheckGrammarAsYouType);
                TryRestoreComProperty(options, "CheckSpellingAsYouType", previousCheckSpellingAsYouType);
                TryRestoreComProperty(options, "Pagination", previousPagination);
                TryRestoreComProperty(app, "DisplayAlerts", previousDisplayAlerts);
                TryRestoreComProperty(app, "ScreenUpdating", previousScreenUpdating);
            }
        }

        // 工具栏加载时：初始化按钮文字 + 刷新样式选中状态
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            if (IsDesignTime())
            {
                return;
            }

            // Ribbon 选项卡必须优先显示，后续初始化失败也不能影响其可见性。
            tab1.Label = "搞快点";
            tab1.Visible = true;

            try
            {
                RegisterInstance();
                UpdateHelpVersionLabel();
                InitializeOutlineLevelDropDown();
                button9.Click += button9_Click;
                button10.Click += button10_Click;
                button11.Click += button11_Click;
                button12.Click += button12_Click;
                InitializeStyleGalleriesLightweight();
            }
            catch
            {
                // 保留已经创建的选项卡，避免单个控件初始化异常导致整个 Ribbon 不显示。
            }
        }

        private void ApplyFeatureAvailability()
        {
            bool softwareDocument = IsCurrentDocumentSoftwareDocument();
            if (btnRequirementExtraction != null)
            {
                btnRequirementExtraction.Enabled = softwareDocument;
            }

            if (btnSoftwareDocumentCheck != null)
            {
                btnSoftwareDocumentCheck.Enabled = softwareDocument;
            }

            if (button12 != null)
            {
                button12.Enabled = softwareDocument && RequirementTrackingEnabled;
            }
        }

        private static bool IsCurrentDocumentSoftwareDocument()
        {
            try
            {
                Word.Document doc = Globals.ThisAddIn?.Application?.ActiveDocument;
                string name = doc?.Name ?? string.Empty;
                return name.IndexOf("软件", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        internal static void RefreshAllStyleIndicators()
        {
            foreach (Ribbon1 ribbon in LoadedInstances.ToArray())
            {
                ribbon?.RefreshCurrentStyleIndicator();
                ribbon?.RefreshNavigationPaneState();
                ribbon?.ApplyFeatureAvailability();
            }
        }

        private void btnSwitchWindows_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            btnSwitchWindows.Items.Clear();
            List<WordWindowItem> windows = GetOpenWordWindows();
            if (windows.Count == 0)
            {
                RibbonButton empty = Factory.CreateRibbonButton();
                empty.Label = "没有打开的 Word 文档";
                empty.Enabled = false;
                btnSwitchWindows.Items.Add(empty);
                return;
            }

            for (int i = 0; i < windows.Count; i++)
            {
                WordWindowItem window = windows[i];
                RibbonMenu item = Factory.CreateRibbonMenu();
                item.Name = "btnSwitchWindowDoc" + i.ToString();
                item.Label = (i + 1) + "  " + window.Title;
                item.OfficeImageId = "FileDocument";
                item.ShowImage = true;
                item.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;

                RibbonButton activateButton = Factory.CreateRibbonButton();
                activateButton.Name = "btnActivateWindow" + i.ToString();
                activateButton.Label = "切换到此文档";
                activateButton.OfficeImageId = "WindowSwitchWindowsMenuWord";
                activateButton.ShowImage = true;
                activateButton.Click += (_, __) => ActivateWordDocumentWindow(window.Handle);

                RibbonButton closeButton = Factory.CreateRibbonButton();
                closeButton.Name = "btnCloseWindow" + i.ToString();
                closeButton.Label = "关闭此文档";
                closeButton.OfficeImageId = "WindowClose";
                closeButton.ShowImage = true;
                closeButton.Click += (_, __) => CloseWordDocumentWindow(window.Handle);

                item.Items.Add(activateButton);
                item.Items.Add(closeButton);
                btnSwitchWindows.Items.Add(item);
            }
        }

        private static void ActivateWordDocumentWindow(IntPtr handle)
        {
            try
            {
                if (IsIconic(handle))
                {
                    ShowWindow(handle, SwRestore);
                }

                SetForegroundWindow(handle);
                TryUpdateStatusBar(Globals.ThisAddIn?.Application, "已切换窗口");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"切换窗口失败: {ex.Message}", "文档管理");
            }
        }

        private static List<WordWindowItem> GetOpenWordWindows()
        {
            List<WordWindowItem> windows = new List<WordWindowItem>();
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd))
                {
                    return true;
                }

                StringBuilder title = new StringBuilder(512);
                GetWindowText(hWnd, title, title.Capacity);
                string text = title.ToString();
                if (IsWordDocumentWindowTitle(text))
                {
                    windows.Add(new WordWindowItem
                    {
                        Handle = hWnd,
                        Title = text
                    });
                }

                return true;
            }, IntPtr.Zero);

            return windows
                .GroupBy(item => item.Handle)
                .Select(group => group.First())
                .OrderBy(item => item.Title, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private static bool IsWordDocumentWindowTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return false;
            }

            return title.IndexOf(" - Word", StringComparison.OrdinalIgnoreCase) >= 0
                || title.IndexOf(" - Microsoft Word", StringComparison.OrdinalIgnoreCase) >= 0
                || title.IndexOf(".doc", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string GetDocumentDisplayName(Word.Document doc)
        {
            try
            {
                return string.IsNullOrWhiteSpace(doc?.Name) ? "未命名文档" : doc.Name;
            }
            catch
            {
                return "未命名文档";
            }
        }

        private void btnOpenCurrentFolder_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            try
            {
                Word.Document doc = app?.ActiveDocument;
                string fullName = doc?.FullName;
                if (doc == null || string.IsNullOrWhiteSpace(fullName) || !File.Exists(fullName))
                {
                    MessageBox.Show("当前文档尚未保存，无法打开所在文件夹。", "文档管理");
                    return;
                }

                Process.Start("explorer.exe", $"/select,\"{fullName}\"");
                TryUpdateStatusBar(app, "已打开所在文件夹");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开所在文件夹失败: {ex.Message}", "文档管理");
            }
        }

        private void btnSaveAllDocuments_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            try
            {
                Word.Documents docs = app?.Documents;
                int count = docs?.Count ?? 0;
                int saved = 0;
                for (int i = 1; i <= count; i++)
                {
                    Word.Document doc = docs[i];
                    doc?.Save();
                    saved++;
                }

                TryUpdateStatusBar(app, $"已保存 {saved} 个文档");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存所有文档失败: {ex.Message}", "文档管理");
            }
        }

        private void btnCloseOtherDocuments_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            try
            {
                Word.Document activeDoc = app?.ActiveDocument;
                Word.Documents docs = app?.Documents;
                int count = docs?.Count ?? 0;
                if (activeDoc == null || count <= 1)
                {
                    TryUpdateStatusBar(app, "没有需要关闭的其他文档");
                    return;
                }

                List<Word.Document> targets = new List<Word.Document>();
                for (int i = 1; i <= count; i++)
                {
                    Word.Document doc = docs[i];
                    if (doc != null && !ReferenceEquals(doc, activeDoc))
                    {
                        targets.Add(doc);
                    }
                }

                foreach (Word.Document doc in targets)
                {
                    doc.Close(Word.WdSaveOptions.wdPromptToSaveChanges);
                }

                activeDoc.Activate();
                TryUpdateStatusBar(app, $"已关闭 {targets.Count} 个其他文档");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"关闭其他文档失败: {ex.Message}", "文档管理");
            }
        }

        private void btnDocumentVersions_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            Word.Document document = app?.ActiveDocument;
            if (document == null)
            {
                MessageBox.Show("当前没有活动文档。", "文档版本管理");
                return;
            }

            try
            {
                if (string.IsNullOrWhiteSpace(document.FullName) || !File.Exists(document.FullName))
                {
                    MessageBox.Show("请先保存当前文档，再使用版本管理。", "文档版本管理");
                    return;
                }

                using (DocumentVersionManagementForm form = new DocumentVersionManagementForm(app, document))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开版本管理失败：\r\n" + ex.Message, "文档版本管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // 刷新当前样式显示，保持和 Word 当前光标所在段落样式一致。
        internal void RefreshCurrentStyleIndicator()
        {
            string currentStyleName = GetCurrentParagraphStyleName();
            RefreshCommonStyleDropDown(currentStyleName);
            SelectOutlineLevelItem(GetCurrentSelectionOutlineLevel());
        }

        private void btnStandardReading_Click(object sender, RibbonControlEventArgs e)
        {
            OpenStandardsFolder();
        }

        private static void OpenStandardsFolder()
        {
            string standardPath = PluginDataStore.StandardPath;
            if (!File.Exists(standardPath))
            {
                MessageBox.Show(
                    "未找到标准文件，请将 GJB438C.pdf 放入以下目录：\r\n" + PluginDataStore.StandardsFolder,
                    "查看标准",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Process.Start("explorer.exe", "\"" + PluginDataStore.StandardsFolder + "\"");
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开标准文件夹失败：\r\n" + ex.Message, "查看标准", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        internal static void RefreshNavigationPaneIndicators()
        {
            foreach (Ribbon1 ribbon in LoadedInstances.ToArray())
            {
                ribbon?.RefreshNavigationPaneState();
            }
        }

        private void btnToggleNavigationPane_Click(object sender, RibbonControlEventArgs e)
        {
            bool documentMapToggledDirectly = false;
            try
            {
                Word.Window activeWindow = Globals.ThisAddIn?.Application?.ActiveWindow;
                if (activeWindow == null)
                {
                    return;
                }

                bool currentlyVisible = activeWindow.DocumentMap;
                activeWindow.DocumentMap = !currentlyVisible;
                documentMapToggledDirectly = true;
            }
            catch
            {
                try
                {
                    Globals.ThisAddIn?.Application?.CommandBars?.ExecuteMso("NavigationPane");
                }
                catch
                {
                }
            }
            finally
            {
                RefreshNavigationPaneState();
                if (!documentMapToggledDirectly)
                {
                    Globals.ThisAddIn?.ScheduleNavigationPaneStateRefresh();
                }
            }
        }

        private void btnOfficeClipboard_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn?.Application?.CommandBars?.ExecuteMso("ShowClipboard");
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开原生剪贴板失败：\r\n" + ex.Message, "剪贴板");
            }
        }

        private void RefreshNavigationPaneState()
        {
            if (btnToggleNavigationPane == null)
            {
                return;
            }

            try
            {
                Word.Window activeWindow = Globals.ThisAddIn?.Application?.ActiveWindow;
                btnToggleNavigationPane.Checked = activeWindow != null && activeWindow.DocumentMap;
            }
            catch
            {
            }
        }

        private void InitializeStyleGalleriesLightweight()
        {
            RefreshCurrentStyleIndicator();
        }

        private void InitializeOutlineLevelDropDown()
        {
            if (outlineLevelDropDown == null)
            {
                return;
            }

            updatingOutlineLevel = true;
            try
            {
                outlineLevelDropDown.Items.Clear();
                for (int level = 1; level <= 9; level++)
                {
                    AddOutlineLevelItem($"{level}级");
                }

                AddOutlineLevelItem("正文");
                SelectOutlineLevelItem(10);
            }
            finally
            {
                updatingOutlineLevel = false;
            }
        }

        private RibbonDropDownItem AddOutlineLevelItem(string label)
        {
            RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
            item.Label = label;
            outlineLevelDropDown.Items.Add(item);
            return item;
        }

        private void SelectOutlineLevelItem(int outlineLevel)
        {
            if (outlineLevelDropDown == null || outlineLevelDropDown.Items.Count == 0)
            {
                return;
            }

            string label = outlineLevel >= 1 && outlineLevel <= 9 ? $"{outlineLevel}级" : "正文";
            updatingOutlineLevel = true;
            try
            {
                foreach (RibbonDropDownItem item in outlineLevelDropDown.Items)
                {
                    if (string.Equals(item.Label, label, StringComparison.OrdinalIgnoreCase))
                    {
                        outlineLevelDropDown.SelectedItem = item;
                        return;
                    }
                }
            }
            finally
            {
                updatingOutlineLevel = false;
            }
        }

        private static int GetCurrentSelectionOutlineLevel()
        {
            try
            {
                Word.Selection selection = Globals.ThisAddIn.Application?.Selection;
                Word.Paragraphs paragraphs = selection?.Range?.Paragraphs;
                if (paragraphs == null || paragraphs.Count == 0)
                {
                    return 10;
                }

                Word.WdOutlineLevel level = paragraphs[1].OutlineLevel;
                return level >= Word.WdOutlineLevel.wdOutlineLevel1 && level <= Word.WdOutlineLevel.wdOutlineLevel9
                    ? (int)level
                    : 10;
            }
            catch
            {
                return 10;
            }
        }

        private static IEnumerable<string> GetDocumentStyleNames(Word.Document doc)
        {
            SortedSet<string> names = new SortedSet<string>(StringComparer.CurrentCultureIgnoreCase);
            try
            {
                Word.Styles styles = doc?.Styles;
                int count = styles?.Count ?? 0;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        Word.Style style = styles[i];
                        if (style == null)
                        {
                            continue;
                        }

                        string name = style.NameLocal;
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name);
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

            return names;
        }


        // 获取Word当前光标所在段落/选区的样式名称
        private string GetCurrentParagraphStyleName()
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Selection selection = app?.Selection;
                Word.Paragraph paragraph = selection?.Range?.Paragraphs?[1];
                if (paragraph?.Range == null)
                    return string.Empty;

                object styleObj = TryGetStyle(paragraph.Range);

                return ResolveStyleName(styleObj, selection?.Document);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ResolveStyleName(object styleObj, Word.Document doc)
        {
            if (styleObj == null)
                return string.Empty;

            if (styleObj is string styleText)
                return styleText;

            if (styleObj is Word.Style wordStyle)
            {
                if (!string.IsNullOrWhiteSpace(wordStyle.NameLocal))
                    return wordStyle.NameLocal;
            }

            string nameLocal = TryGetComPropertyAsString(styleObj, "NameLocal");
            if (!string.IsNullOrWhiteSpace(nameLocal))
                return nameLocal;

            string name = TryGetComPropertyAsString(styleObj, "Name");
            if (!string.IsNullOrWhiteSpace(name))
                return name;

            try
            {
                if (doc != null)
                {
                    object key = styleObj;
                    Word.Style resolvedStyle = doc.Styles[key];
                    if (resolvedStyle != null)
                    {
                        if (!string.IsNullOrWhiteSpace(resolvedStyle.NameLocal))
                            return resolvedStyle.NameLocal;
                    }
                }
            }
            catch
            {
            }

            return Convert.ToString(styleObj) ?? string.Empty;
        }

        private void group2_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            ShowTablesAndFiguresFormattingSettingsDialog();
        }

        private void groupSoftwareTools_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (Globals.ThisAddIn?.Application?.ActiveDocument == null)
                {
                    MessageBox.Show("当前没有活动文档。", "提取设置");
                    return;
                }

                Globals.ThisAddIn.OpenRequirementExtractionSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "打开提取设置失败：\r\n" + ex.Message,
                    "提取设置",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void button25_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteNormalizeSelectedTableAction();
        }

        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertImageCaption();
        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertTableCaption();
        }

        private void btnUpdateCaptions_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                ResetOperationCancellation("更新全部题注");
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Document doc = app?.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("当前没有活动文档。", "文档不加班");
                    return;
                }

                int imageCount = RefreshImageCaptions(doc);
                int tableCount = RefreshTableCaptions(doc);
                UpdateCaptionReferenceFields(doc);

                MessageBox.Show(
                    imageCount + tableCount > 0
                        ? $"已更新 {imageCount} 个图片题注、{tableCount} 个表格题注。"
                        : "未找到可更新的题注。",
                    "文档不加班");
            }
            catch (OperationCanceledException)
            {
                TryUpdateStatusBar(Globals.ThisAddIn?.Application, "题注更新已停止");
                MessageBox.Show("题注更新已停止。", "文档不加班");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"更新题注失败: {ex.Message}", "文档不加班");
            }
        }

        private void outlineLevelDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (updatingOutlineLevel || outlineLevelDropDown?.SelectedItem == null)
            {
                return;
            }

            ApplyOutlineLevelToSelection(ParseOutlineLevelLabel(outlineLevelDropDown.SelectedItem.Label));
        }

        private static int ParseOutlineLevelLabel(string label)
        {
            return label == "正文"
                ? 10
                : int.TryParse((label ?? string.Empty).Replace("级", string.Empty), out int level) ? level : 10;
        }

        private void ApplyOutlineLevelToSelection(int level)
        {
            if ((level < 1 || level > 9) && level != 10)
            {
                return;
            }

            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Selection selection = app?.Selection;
                Word.Paragraphs paragraphs = selection?.Range?.Paragraphs;
                if (paragraphs == null || paragraphs.Count == 0)
                {
                    return;
                }

                Word.WdOutlineLevel outlineLevel = level == 10
                    ? Word.WdOutlineLevel.wdOutlineLevelBodyText
                    : (Word.WdOutlineLevel)level;
                foreach (Word.Paragraph paragraph in paragraphs)
                {
                    try
                    {
                        paragraph.OutlineLevel = outlineLevel;
                        paragraph.Range.ParagraphFormat.OutlineLevel = outlineLevel;
                    }
                    catch
                    {
                    }
                }

                app?.ScreenRefresh();
                SelectOutlineLevelItem(level);
                TryUpdateStatusBar(app, level == 10 ? "正文" : $"大纲{level}级");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"设置大纲级别失败: {ex.Message}", "文档不加班");
            }
        }

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            Word.Document doc = Globals.ThisAddIn?.Application?.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("当前没有活动文档。", "常用样式");
                return;
            }

            try
            {
                string styleLibraryDocumentKey = GetCommonStylesDocumentKey(doc);
                List<string> cachedStyleNames = GetCachedHeadingStyleNames(styleLibraryDocumentKey);
                using (CommonStyleSettingsForm form = new CommonStyleSettingsForm(
                    progress => LoadHeadingStyleNames(doc, styleLibraryDocumentKey, progress),
                    GetConfiguredCommonStyleNames(),
                    cachedStyleNames,
                    GetChapterNumberRepairFormattingSettings()))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    SaveConfiguredCommonStyleNames(form.SelectedStyleNames);
                    SaveChapterNumberRepairFormattingSettings(form.FormattingSettings);
                    RefreshCommonStyleDropDown(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存常用样式失败: {ex.Message}", "常用样式");
            }
        }

        private void commonStylesDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (updatingCommonStyles)
            {
                return;
            }

            string styleName = commonStylesDropDown?.SelectedItem?.Label;
            if (string.IsNullOrWhiteSpace(styleName))
            {
                return;
            }

            Word.Application app = Globals.ThisAddIn?.Application;
            Word.Document doc = app?.ActiveDocument;
            Word.Selection selection = app?.Selection;
            if (doc == null || selection?.Range == null)
            {
                MessageBox.Show("当前没有可设置样式的段落。", "常用样式");
                return;
            }

            if (!DocumentContainsStyle(doc, styleName))
            {
                MessageBox.Show($"当前文档中不存在样式“{styleName}”。", "常用样式");
                RefreshCommonStyleDropDown(true);
                return;
            }

            try
            {
                Word.Paragraph paragraph = selection.Range.Paragraphs[1];
                if (!TrySetDocumentParagraphStyle(doc, paragraph?.Range, styleName))
                {
                    throw new InvalidOperationException("Word 未能应用该样式。");
                }

                TryUpdateStatusBar(app, "已应用常用样式：" + styleName);
                RefreshCurrentStyleIndicator();
                // 清空当前下拉选择，使同一种样式可连续应用到不同段落。
                RefreshCommonStyleDropDown(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("应用常用样式失败：\r\n" + ex.Message, "常用样式");
            }
        }

        private void RefreshCommonStyleDropDown(bool force = false)
        {
            RefreshCommonStyleDropDown(GetCurrentParagraphStyleName(), force);
        }

        private void RefreshCommonStyleDropDown(string currentStyleName, bool force = false)
        {
            if (commonStylesDropDown == null)
            {
                return;
            }

            Word.Document doc = Globals.ThisAddIn?.Application?.ActiveDocument;
            string documentKey = GetCommonStylesDocumentKey(doc);
            updatingCommonStyles = true;
            try
            {
                bool documentChanged = !string.Equals(commonStylesDocumentKey, documentKey, StringComparison.OrdinalIgnoreCase);
                if (force || documentChanged)
                {
                    commonStylesDropDown.Items.Clear();
                    if (doc != null)
                    {
                        // 常用样式是插件级配置，不应因为切换文档或 Word 的样式索引
                        // 偶发解析失败而从下拉框消失。真正应用前仍会校验当前文档样式。
                        foreach (string styleName in GetConfiguredCommonStyleNames())
                        {
                            RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                            item.Label = styleName;
                            commonStylesDropDown.Items.Add(item);
                        }
                    }
                }

                commonStylesDropDown.Enabled = commonStylesDropDown.Items.Count > 0;
                RibbonDropDownItem currentItem = commonStylesDropDown.Items
                    .Cast<RibbonDropDownItem>()
                    .FirstOrDefault(item =>
                        !string.IsNullOrWhiteSpace(currentStyleName) &&
                        string.Equals(item.Label, currentStyleName.Trim(), StringComparison.OrdinalIgnoreCase));
                commonStylesDropDown.SelectedItem = currentItem;
                commonStylesDocumentKey = documentKey;
            }
            finally
            {
                updatingCommonStyles = false;
            }
        }

        private static List<string> GetConfiguredCommonStyleNames()
        {
            try
            {
                return (Properties.Settings.Default.CommonStyleNames ?? string.Empty)
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(name => name.Trim())
                    .Where(name => name.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(9)
                    .ToList();
            }
            catch
            {
                return new List<string>();
            }
        }

        private static void SaveConfiguredCommonStyleNames(IEnumerable<string> styleNames)
        {
            List<string> names = (styleNames ?? Enumerable.Empty<string>())
                .Select(name => (name ?? string.Empty).Trim())
                .Where(name => name.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(9)
                .ToList();
            Properties.Settings.Default.CommonStyleNames = string.Join("\n", names);
            Properties.Settings.Default.Save();
        }

        private static ChapterNumberRepairFormattingSettings GetChapterNumberRepairFormattingSettings()
        {
            ChapterNumberRepairFormattingSettings settings = new ChapterNumberRepairFormattingSettings();
            try
            {
                string[] values = (Properties.Settings.Default.ChapterNumberRepairFormatting ?? string.Empty)
                    .Split(new[] { '\u001f' });
                if (values.Length != 10)
                {
                    return settings;
                }

                settings.ApplyFormatting = bool.TryParse(values[0], out bool apply) && apply;
                settings.LevelOneFontName = string.IsNullOrWhiteSpace(values[1]) ? "黑体" : values[1];
                settings.LevelOneFontSize = ParseSettingDecimal(values[2], 12M);
                settings.OtherLevelsFontName = string.IsNullOrWhiteSpace(values[3]) ? "宋体" : values[3];
                settings.OtherLevelsFontSize = ParseSettingDecimal(values[4], 12M);
                settings.Bold = bool.TryParse(values[5], out bool bold) && bold;
                settings.Alignment = ParseSettingInt(values[6], 0, 0, 3);
                settings.LineSpacingRule = ParseSettingInt(values[7], 0, 0, 2);
                settings.SpaceBefore = ParseSettingDecimal(values[8], 0M);
                settings.SpaceAfter = ParseSettingDecimal(values[9], 0M);
            }
            catch
            {
            }

            return settings;
        }

        private static void SaveChapterNumberRepairFormattingSettings(ChapterNumberRepairFormattingSettings settings)
        {
            settings = settings ?? new ChapterNumberRepairFormattingSettings();
            Properties.Settings.Default.ChapterNumberRepairFormatting = string.Join("\u001f", new[]
            {
                settings.ApplyFormatting.ToString(),
                settings.LevelOneFontName ?? "黑体",
                settings.LevelOneFontSize.ToString(System.Globalization.CultureInfo.InvariantCulture),
                settings.OtherLevelsFontName ?? "宋体",
                settings.OtherLevelsFontSize.ToString(System.Globalization.CultureInfo.InvariantCulture),
                settings.Bold.ToString(),
                settings.Alignment.ToString(),
                settings.LineSpacingRule.ToString(),
                settings.SpaceBefore.ToString(System.Globalization.CultureInfo.InvariantCulture),
                settings.SpaceAfter.ToString(System.Globalization.CultureInfo.InvariantCulture)
            });
            Properties.Settings.Default.Save();
        }

        private static decimal ParseSettingDecimal(string value, decimal defaultValue)
        {
            return decimal.TryParse(value, System.Globalization.NumberStyles.Number,
                System.Globalization.CultureInfo.InvariantCulture, out decimal parsed)
                ? Math.Max(0M, parsed)
                : defaultValue;
        }

        private static int ParseSettingInt(string value, int defaultValue, int minimum, int maximum)
        {
            return int.TryParse(value, out int parsed)
                ? Math.Max(minimum, Math.Min(maximum, parsed))
                : defaultValue;
        }

        private static List<string> GetDocumentParagraphStyleNames(
            Word.Document doc,
            Action<int, int> progress = null)
        {
            List<string> names = new List<string>();
            if (doc == null)
            {
                return names;
            }

            try
            {
                Word.Styles styles = doc.Styles;
                int count = styles?.Count ?? 0;
                progress?.Invoke(0, count);
                for (int index = 1; index <= count; index++)
                {
                    if (index == 1 || index == count || index % 10 == 0)
                    {
                        progress?.Invoke(index, count);
                    }

                    Word.Style style = styles[index];
                    if (style == null || style.Type != Word.WdStyleType.wdStyleTypeParagraph)
                    {
                        continue;
                    }

                    string styleName = style.NameLocal;
                    if (!string.IsNullOrWhiteSpace(styleName))
                    {
                        names.Add(styleName.Trim());
                    }
                }
            }
            catch
            {
            }

            return names
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private static string GetCommonStylesDocumentKey(Word.Document doc)
        {
            if (doc == null)
            {
                return string.Empty;
            }

            try
            {
                return string.IsNullOrWhiteSpace(doc.FullName) ? doc.Name ?? string.Empty : doc.FullName;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool DocumentContainsStyle(Word.Document doc, string styleName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(styleName))
            {
                return false;
            }

            try
            {
                if (doc.Styles[styleName] != null)
                {
                    return true;
                }
            }
            catch
            {
            }

            // 已由“加载样式库”读出的本地样式名可避免 Word 样式索引偶发查找失败。
            return GetCachedHeadingStyleNames(GetCommonStylesDocumentKey(doc))
                .Any(name => string.Equals(name, styleName, StringComparison.OrdinalIgnoreCase));
        }

        private static bool TrySetDocumentParagraphStyle(
            Word.Document doc,
            Word.Range range,
            string styleName)
        {
            if (doc == null || range == null || string.IsNullOrWhiteSpace(styleName))
            {
                return false;
            }

            try
            {
                Word.Style style = doc.Styles[styleName];
                if (style != null)
                {
                    object styleValue = style;
                    range.set_Style(ref styleValue);
                    return true;
                }
            }
            catch
            {
            }

            return TrySetStyle(range, styleName);
        }

        private static StyleDefinitionRequest GetStyleDefinitionForOutlineLevel(int level)
        {
            EnsureStyleDefinitionsInitialized();
            if (styleDefinitions.TryGetValue(level, out StyleDefinitionRequest definition))
            {
                return definition;
            }

            return StyleDefinitionRequest.CreateDefaultSet().FirstOrDefault(item => item.Level == level);
        }

        private static bool TrySetStyle(object target, string styleName)
        {
            if (target == null || string.IsNullOrWhiteSpace(styleName))
                return false;

            object styleValue = styleName;

            try
            {
                if (target is Word.Range wordRange)
                {
                    wordRange.set_Style(ref styleValue);
                    return true;
                }

                if (target is Word.Selection wordSelection)
                {
                    wordSelection.set_Style(ref styleValue);
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                target.GetType().InvokeMember("Style", BindingFlags.SetProperty, null, target, new object[] { styleName });
                return true;
            }
            catch
            {
            }

            try
            {
                target.GetType().InvokeMember("set_Style", BindingFlags.InvokeMethod, null, target, new object[] { styleName });
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static object TryGetStyle(object target)
        {
            if (target == null)
                return null;

            try
            {
                if (target is Word.Range wordRange)
                {
                    return wordRange.get_Style();
                }

                if (target is Word.Selection wordSelection)
                {
                    return wordSelection.get_Style();
                }
            }
            catch
            {
            }

            return null;
        }

        private static string TryGetComPropertyAsString(object target, string propertyName)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
                return string.Empty;

            try
            {
                object value = target.GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    target,
                    null);

                return Convert.ToString(value) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static object TryGetComProperty(object target, string propertyName)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
                return null;

            try
            {
                return target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty,
                    null,
                    target,
                    null);
            }
            catch
            {
                return null;
            }
        }

        private static void TrySetComProperty(object target, string propertyName, object value)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
                return;

            try
            {
                target.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.SetProperty,
                    null,
                    target,
                    new object[] { value });
            }
            catch
            {
            }
        }

        private static void TryRestoreComProperty(object target, string propertyName, object previousValue)
        {
            if (previousValue == null)
                return;

            TrySetComProperty(target, propertyName, previousValue);
        }

        private static object TryGetParagraphStyle(Word.Range range)
        {
            if (range == null)
                return null;

            try
            {
                if (range.Paragraphs == null || range.Paragraphs.Count < 1)
                    return null;

                Word.Paragraph firstParagraph = range.Paragraphs.First;
                return firstParagraph?.get_Style();
            }
            catch
            {
                return null;
            }
        }

        private void RegisterInstance()
        {
            if (!LoadedInstances.Contains(this))
            {
                LoadedInstances.Add(this);
            }
        }

        private void DisposeRuntimeResources()
        {
            LoadedInstances.Remove(this);
            if (LoadedInstances.Count == 0)
            {
                CancelStyleBrush(false);
            }
        }

        private static void CloseWordDocumentWindow(IntPtr handle)
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            Word.Windows wordWindows = null;
            Word.Window targetWindow = null;
            Word.Document targetDocument = null;
            try
            {
                wordWindows = app?.Windows;
                int count = wordWindows?.Count ?? 0;
                for (int i = 1; i <= count; i++)
                {
                    Word.Window candidate = wordWindows[i];
                    if (candidate != null && new IntPtr(candidate.Hwnd) == handle)
                    {
                        targetWindow = candidate;
                        break;
                    }

                    if (candidate != null)
                    {
                        Marshal.ReleaseComObject(candidate);
                    }
                }

                if (targetWindow == null)
                {
                    MessageBox.Show("所选文档窗口已经关闭。", "文档管理");
                    return;
                }

                targetDocument = targetWindow.Document;
                string documentName = GetDocumentDisplayName(targetDocument);
                targetDocument.Close(Word.WdSaveOptions.wdPromptToSaveChanges);
                TryUpdateStatusBar(app, "已关闭 " + documentName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"关闭文档失败: {ex.Message}", "文档管理");
            }
            finally
            {
                if (targetDocument != null)
                {
                    Marshal.ReleaseComObject(targetDocument);
                }

                if (targetWindow != null)
                {
                    Marshal.ReleaseComObject(targetWindow);
                }

                if (wordWindows != null)
                {
                    Marshal.ReleaseComObject(wordWindows);
                }
            }
        }

        private void btnPersistentStyleBrush_Click(object sender, RibbonControlEventArgs e)
        {
            ToggleStyleBrush();
        }

        internal static void HandleStyleBrushSelectionChange(Word.Selection selection)
        {
            if (!styleBrushActive || styleBrushApplying || selection?.Range == null)
            {
                return;
            }

            string documentKey = GetStyleBrushDocumentKey(selection.Document);
            if (!string.Equals(documentKey, styleBrushSourceDocumentKey, StringComparison.OrdinalIgnoreCase))
            {
                CancelStyleBrush(false);
                return;
            }

            int start;
            int end;
            try
            {
                start = selection.Range.Start;
                end = selection.Range.End;
            }
            catch
            {
                return;
            }

            if (start == styleBrushSourceStart && end == styleBrushSourceEnd)
            {
                return;
            }

            // 导航窗格也会移动 Word 插入点。只有鼠标位于正文插入点附近时，
            // 才把折叠选区视为一次正文格式刷操作。
            if (start == end && !IsMouseNearStyleBrushSelection(selection))
            {
                return;
            }

            styleBrushApplying = true;
            try
            {
                selection.PasteFormat();
                if (styleBrushPersistent)
                {
                    SetStyleBrushStatus("格式刷已锁定，可继续选择目标内容；再次点击格式刷或按 Esc 退出。");
                }
                else
                {
                    CancelStyleBrush(false);
                }
            }
            catch
            {
                CancelStyleBrush(false);
            }
            finally
            {
                styleBrushApplying = false;
            }
        }

        internal static void HandleStyleBrushWindowActivated(Word.Document document)
        {
            if (!styleBrushActive)
            {
                return;
            }

            string documentKey = GetStyleBrushDocumentKey(document);
            if (!string.Equals(documentKey, styleBrushSourceDocumentKey, StringComparison.OrdinalIgnoreCase))
            {
                CancelStyleBrush(false);
            }
        }

        private static void ToggleStyleBrush()
        {
            Word.Application app = Globals.ThisAddIn?.Application;
            Word.Selection selection = app?.Selection;
            if (selection?.Range == null)
            {
                return;
            }

            long nowTicks = DateTime.UtcNow.Ticks;
            long elapsedTicks = nowTicks - styleBrushLastClickTicks;
            long doubleClickTicks = TimeSpan.FromMilliseconds(SystemInformation.DoubleClickTime).Ticks;
            bool secondClick = styleBrushActive
                && elapsedTicks >= 0
                && elapsedTicks <= doubleClickTicks;

            if (secondClick)
            {
                styleBrushPersistent = true;
                styleBrushLastClickTicks = 0;
                SetStyleBrushVisualState(true);
                SetStyleBrushStatus("格式刷已锁定，可连续应用格式；再次点击格式刷或按 Esc 退出。");
                return;
            }

            if (styleBrushActive)
            {
                CancelStyleBrush(true);
                return;
            }

            try
            {
                selection.CopyFormat();
                styleBrushSourceDocumentKey = GetStyleBrushDocumentKey(selection.Document);
                styleBrushSourceStart = selection.Range.Start;
                styleBrushSourceEnd = selection.Range.End;
                styleBrushPersistent = false;
                styleBrushActive = true;
                styleBrushLastClickTicks = nowTicks;
                SetStyleBrushVisualState(false);
                EnsureStopShortcutHook();
                SetStyleBrushStatus("格式刷已启用，请选择要应用格式的内容。");
            }
            catch (Exception ex)
            {
                CancelStyleBrush(false);
                MessageBox.Show($"启动格式刷失败: {ex.Message}", "文档不加班");
            }
        }

        private static void CancelStyleBrush(bool showStatus)
        {
            bool wasActive = styleBrushActive;
            styleBrushActive = false;
            styleBrushPersistent = false;
            styleBrushLastClickTicks = 0;
            styleBrushSourceDocumentKey = string.Empty;
            styleBrushSourceStart = -1;
            styleBrushSourceEnd = -1;
            SetStyleBrushVisualState(false);
            if (showStatus && wasActive)
            {
                SetStyleBrushStatus("格式刷已取消。");
            }
        }

        private static string GetStyleBrushDocumentKey(Word.Document document)
        {
            try
            {
                return !string.IsNullOrWhiteSpace(document?.FullName)
                    ? document.FullName
                    : document?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static bool IsMouseNearStyleBrushSelection(Word.Selection selection)
        {
            try
            {
                int left;
                int top;
                int width;
                int height;
                selection.Application.ActiveWindow.GetPoint(
                    out left,
                    out top,
                    out width,
                    out height,
                    selection.Range);

                var cursor = Cursor.Position;
                const int horizontalTolerance = 18;
                const int verticalTolerance = 10;
                return cursor.X >= left - horizontalTolerance
                    && cursor.X <= left + Math.Max(width, 2) + horizontalTolerance
                    && cursor.Y >= top - verticalTolerance
                    && cursor.Y <= top + Math.Max(height, 16) + verticalTolerance;
            }
            catch
            {
                return false;
            }
        }

        private static void SetStyleBrushStatus(string message)
        {
            try
            {
                Word.Application app = Globals.ThisAddIn?.Application;
                if (app != null)
                {
                    app.StatusBar = message;
                }
            }
            catch
            {
            }
        }

        private static void SetStyleBrushVisualState(bool locked)
        {
            foreach (Ribbon1 ribbon in LoadedInstances.ToArray())
            {
                try
                {
                    if (ribbon?.btnStyleBrush != null)
                    {
                        ribbon.btnStyleBrush.Checked = locked;
                    }
                }
                catch
                {
                }
            }
        }

        private static void TryUpdateStatusBar(Word.Application app, string styleName)
        {
            if (app == null)
                return;

            try
            {
                app.StatusBar = string.IsNullOrWhiteSpace(styleName)
                    ? "DocuLint 当前样式: <空>"
                    : "DocuLint 当前样式: " + styleName;
            }
            catch
            {
            }
        }

        internal static void ResetOperationCancellation(string operationName = null)
        {
            operationCancelRequested = false;
            EnsureStopShortcutHook();

            if (!string.IsNullOrWhiteSpace(operationName))
            {
                TryShowCancellableOperationHint(Globals.ThisAddIn?.Application, operationName);
            }
        }

        private static void TryShowCancellableOperationHint(Word.Application app, string operationName)
        {
            if (app == null || string.IsNullOrWhiteSpace(operationName))
            {
                return;
            }

            try
            {
                app.StatusBar = "按 ESC 取消：" + operationName.Trim();
            }
            catch
            {
            }
        }

        internal static void RequestOperationCancellation()
        {
            operationCancelRequested = true;
        }

        internal static void EnableCommonPhraseShortcutHook()
        {
            EnsureStopShortcutHook();
        }

        internal static void DisableCommonPhraseShortcutHook()
        {
            ReleaseStopShortcutHook();
        }

        internal static void ThrowIfOperationCancelled()
        {
            try
            {
                Application.DoEvents();
            }
            catch
            {
            }

            if (operationCancelRequested)
            {
                throw new OperationCanceledException("操作已停止。");
            }
        }

        private static void EnsureStopShortcutHook()
        {
            if (keyboardHookHandle != IntPtr.Zero)
            {
                return;
            }

            keyboardHookProc = StopShortcutHookCallback;
            IntPtr moduleHandle = IntPtr.Zero;
            try
            {
                moduleHandle = GetModuleHandle(Process.GetCurrentProcess().MainModule?.ModuleName);
            }
            catch
            {
            }

            keyboardHookHandle = SetWindowsHookEx(WhKeyboardLl, keyboardHookProc, moduleHandle, 0);
        }

        private static void ReleaseStopShortcutHook()
        {
            if (keyboardHookHandle == IntPtr.Zero)
            {
                return;
            }

            try
            {
                UnhookWindowsHookEx(keyboardHookHandle);
            }
            catch
            {
            }

            keyboardHookHandle = IntPtr.Zero;
            keyboardHookProc = null;
        }

        private static IntPtr StopShortcutHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WmKeyDown || wParam == (IntPtr)WmSysKeyDown))
            {
                int keyCode = Marshal.ReadInt32(lParam);
                if (keyCode == VkEscape)
                {
                    CancelStyleBrush(true);
                    RequestOperationCancellation();
                }
                else if (keyCode == VkA && IsKeyDown(VkControl) && IsKeyDown(VkMenu))
                {
                    Globals.ThisAddIn?.AddSelectedTextToRequirementExtraction();
                }
            }

            return CallNextHookEx(keyboardHookHandle, nCode, wParam, lParam);
        }

        private static bool IsKeyDown(int virtualKey)
        {
            return (GetKeyState(virtualKey) & unchecked((short)0x8000)) != 0;
        }

        private static void EnsureStyleDefinitionsInitialized()
        {
            if (styleDefinitions != null)
            {
                return;
            }

            styleDefinitions = StyleDefinitionRequest.CreateDefaultSet()
                .ToDictionary(item => item.Level, CloneStyleDefinition);
        }

        // 大纲编号仍使用独立的默认字体定义，避免与常用样式选择产生耦合。
        private static StyleDefinitionRequest CloneStyleDefinition(StyleDefinitionRequest definition)
        {
            if (definition == null)
            {
                return null;
            }

            return new StyleDefinitionRequest
            {
                Level = definition.Level,
                OutlineLevel = definition.OutlineLevel,
                ShouldCreate = definition.ShouldCreate,
                StyleName = definition.StyleName,
                FontName = definition.FontName,
                FontSize = definition.FontSize,
                ListFontName = definition.ListFontName,
                ListFontSize = definition.ListFontSize,
                Alignment = definition.Alignment,
                Bold = definition.Bold,
                LineSpacing = definition.LineSpacing
            };
        }

        private static bool IsDesignTime()
        {
            return LicenseManager.UsageMode == LicenseUsageMode.Designtime;
        }

        private void button12_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.ShowRequirementTrackingPane();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开需求追踪控制台失败: {ex.Message}", "文档不加班");
            }
        }

        private void btnRequirementExtraction_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.ShowRequirementExtractionPane();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开需求提取窗格失败: {ex.Message}", "文档不加班");
            }
        }

    }
}
