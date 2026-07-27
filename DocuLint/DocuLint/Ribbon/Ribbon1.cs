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
        private readonly Dictionary<string, List<string>> documentStyleNamesCache =
            new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<int, string> outlineLevelStyleLinks = new Dictionary<int, string>();
        private static Dictionary<int, StyleDefinitionRequest> styleDefinitions;
        private static OutlineNumberPattern outlineNumberPattern = OutlineNumberPattern.Decimal;
        private static int outlineNumberTextSpacing = 1;
        private static volatile bool operationCancelRequested;
        private static IntPtr keyboardHookHandle = IntPtr.Zero;
        private static LowLevelKeyboardProc keyboardHookProc;
        private const string StyleGalleryPlaceholderLabel = "当前样式";
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

            RegisterInstance();
            UpdateHelpVersionLabel();
            InitializeOutlineLevelDropDown();
            button9.Click += button9_Click;
            button10.Click += button10_Click;
            button11.Click += button11_Click;
            button12.Click += button12_Click;

            InitializeStyleGalleriesLightweight();
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
            if (btnSwitchWindows == null)
            {
                return;
            }

            btnSwitchWindows.Items.Clear();
            List<WordWindowItem> windows = GetOpenWordWindows();
            if (windows.Count == 0)
            {
                RibbonButton empty = Factory.CreateRibbonButton();
                empty.Label = "无打开文档";
                empty.Enabled = false;
                btnSwitchWindows.Items.Add(empty);
                return;
            }

            for (int i = 0; i < windows.Count; i++)
            {
                WordWindowItem window = windows[i];
                RibbonButton item = Factory.CreateRibbonButton();
                item.Name = "btnSwitchWindowDoc" + i.ToString();
                item.Label = (i + 1) + " " + window.Title;
                item.OfficeImageId = "FileDocument";
                item.ShowImage = true;
                item.Click += btnSwitchWindowDocument_Click;
                btnSwitchWindows.Items.Add(item);
            }
        }

        private void btnSwitchWindowDocument_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (!(sender is RibbonButton button))
                {
                    return;
                }

                string digits = new string((button.Name ?? string.Empty).Where(char.IsDigit).ToArray());
                if (!int.TryParse(digits, out int index))
                {
                    return;
                }

                List<WordWindowItem> windows = GetOpenWordWindows();
                if (index < 0 || index >= windows.Count)
                {
                    return;
                }

                if (IsIconic(windows[index].Handle))
                {
                    ShowWindow(windows[index].Handle, SwRestore);
                }

                SetForegroundWindow(windows[index].Handle);
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

        // 刷新当前样式显示，保持和 Word 当前光标所在段落样式一致。
        internal void RefreshCurrentStyleIndicator()
        {
            if (styleGalleryDropDown == null)
            {
                return;
            }

            styleGalleryDropDown.Label = "当前样式：" + BuildCurrentStyleGalleryLabel(GetCurrentParagraphStyleName());
            SelectOutlineLevelItem(GetCurrentSelectionOutlineLevel());
        }

        private void btnToggleNavigationPane_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Word.Window activeWindow = Globals.ThisAddIn?.Application?.ActiveWindow;
                if (activeWindow == null)
                {
                    return;
                }

                activeWindow.DocumentMap = btnToggleNavigationPane.Checked;
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

        private static string FormatStyleGalleryLabel(string styleName)
        {
            const int maxLength = 24;
            if (string.IsNullOrWhiteSpace(styleName) || styleName.Length <= maxLength)
            {
                return styleName;
            }

            return styleName.Substring(0, maxLength - 3) + "...";
        }

        private static string BuildCurrentStyleGalleryLabel(string styleName)
        {
            return string.IsNullOrWhiteSpace(styleName) ? "<空>" : FormatStyleGalleryLabel(styleName);
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
                ResetOperationCancellation();
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
                        string linkedStyle = GetLinkedStyleForOutlineLevel(level);
                        paragraph.OutlineLevel = outlineLevel;
                        paragraph.Range.ParagraphFormat.OutlineLevel = outlineLevel;
                        if (!string.IsNullOrWhiteSpace(linkedStyle) && TrySetStyle(paragraph.Range, linkedStyle))
                        {
                            paragraph.OutlineLevel = outlineLevel;
                            paragraph.Range.ParagraphFormat.OutlineLevel = outlineLevel;
                        }
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

        private void btnStyleBinding_Click(object sender, RibbonControlEventArgs e)
        {
            ShowStyleBindingDialog();
        }

        private void ShowStyleBindingDialog()
        {
            try
            {
                Word.Document activeDoc = null;
                try
                {
                    activeDoc = Globals.ThisAddIn?.Application?.ActiveDocument;
                }
                catch
                {
                }

                List<string> styleNames = GetInitialStyleNamesForBinding();

                using (StyleLinkSettingsForm form = new StyleLinkSettingsForm(
                    outlineLevelStyleLinks,
                    styleNames,
                    () => GetCachedDocumentStyleNames(activeDoc),
                    GetCustomStyleNamesForBinding(),
                    outlineNumberPattern,
                    outlineNumberTextSpacing,
                    styleDefinitions.Values))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    outlineLevelStyleLinks.Clear();
                    foreach (KeyValuePair<int, string> item in form.StyleLinks)
                    {
                        if (!string.IsNullOrWhiteSpace(item.Value))
                        {
                            outlineLevelStyleLinks[item.Key] = item.Value;
                        }
                    }
                    outlineNumberPattern = form.NumberPattern;
                    outlineNumberTextSpacing = form.NumberTextSpacing;
                    SaveStyleDefinitions(form.StyleDefinitions);

                    ApplyStyleLinksToCurrentDocument(activeDoc);
                    InitializeStyleGalleriesLightweight();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开样式绑定面板失败: {ex.Message}", "文档不加班");
            }
        }

        private List<string> GetInitialStyleNamesForBinding()
        {
            return outlineLevelStyleLinks.Values
                .Concat(new[] { GetCurrentParagraphStyleName() })
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(name => name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private static List<string> GetCustomStyleNamesForBinding()
        {
            EnsureStyleDefinitionsInitialized();
            return styleDefinitions.Values
                .Select(definition => definition.StyleName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private List<string> GetCachedDocumentStyleNames(Word.Document doc)
        {
            string key = GetDocumentCacheKey(doc);
            if (!string.IsNullOrWhiteSpace(key) && documentStyleNamesCache.TryGetValue(key, out List<string> cached))
            {
                return cached;
            }

            List<string> styles = GetDocumentStyleNames(doc).ToList();
            if (!string.IsNullOrWhiteSpace(key))
            {
                documentStyleNamesCache[key] = styles;
            }

            return styles;
        }

        private static string GetDocumentCacheKey(Word.Document doc)
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

        private void InvalidateDocumentStyleCache(Word.Document doc)
        {
            string key = GetDocumentCacheKey(doc);
            if (!string.IsNullOrWhiteSpace(key))
            {
                documentStyleNamesCache.Remove(key);
            }
        }

        private void btnCreateCustomStyles_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                EnsureStyleDefinitionsInitialized();
                Word.Application app = Globals.ThisAddIn?.Application;
                using (CustomStyleLibraryForm form = new CustomStyleLibraryForm(styleDefinitions.Values))
                {
                    if (form.ShowDialog() != DialogResult.OK)
                    {
                        return;
                    }

                    Word.Document doc = app?.ActiveDocument;
                    if (doc == null)
                    {
                        MessageBox.Show("当前没有活动文档。", "文档不加班");
                        return;
                    }

                    SaveStyleDefinitions(form.StyleDefinitions);
                    List<StyleDefinitionRequest> definitionsToCreate = form.StyleDefinitions
                        .Where(definition => definition.ShouldCreate)
                        .ToList();

                    List<string> duplicateStyleNames = definitionsToCreate
                        .Where(definition => DocumentContainsStyle(doc, definition.StyleName))
                        .Select(definition => definition.StyleName)
                        .ToList();
                    if (duplicateStyleNames.Count > 0)
                    {
                        string names = string.Join("、", duplicateStyleNames);
                        DialogResult overwriteResult = MessageBox.Show(
                            "当前文档已存在同名样式：" + names + "。\r\n是否覆盖这些样式并继续创建？",
                            "创建自定义样式",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning);
                        if (overwriteResult != DialogResult.Yes)
                        {
                            return;
                        }
                    }

                    using (new WordPerformanceScope(app))
                    {
                        foreach (StyleDefinitionRequest definition in definitionsToCreate)
                        {
                            CreateOrUpdateDocumentStyle(doc, definition);
                        }
                    }

                    InvalidateDocumentStyleCache(doc);
                    InitializeStyleGalleriesLightweight();
                    TryUpdateStatusBar(app, "自定义样式已创建");
                    MessageBox.Show("所选自定义样式已应用到当前文档。", "创建自定义样式");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建自定义样式失败: {ex.Message}", "文档不加班");
            }
        }

        private static void SaveStyleDefinitions(IEnumerable<StyleDefinitionRequest> definitions)
        {
            EnsureStyleDefinitionsInitialized();
            styleDefinitions.Clear();
            foreach (StyleDefinitionRequest definition in definitions ?? StyleDefinitionRequest.CreateDefaultSet())
            {
                if (definition != null)
                {
                    styleDefinitions[definition.Level] = CloneStyleDefinition(definition);
                }
            }
        }

        private static StyleDefinitionRequest CloneStyleDefinition(StyleDefinitionRequest definition)
        {
            return new StyleDefinitionRequest
            {
                Level = definition.Level,
                OutlineLevel = definition.OutlineLevel,
                ShouldCreate = definition.ShouldCreate,
                StyleName = definition.StyleName,
                FontName = definition.FontName,
                FontSize = definition.FontSize,
                ListFontName = string.IsNullOrWhiteSpace(definition.ListFontName) ? definition.FontName : definition.ListFontName,
                ListFontSize = definition.ListFontSize > 0f ? definition.ListFontSize : definition.FontSize,
                Alignment = definition.Alignment,
                Bold = definition.Bold,
                LineSpacing = definition.LineSpacing
            };
        }

        private static void CreateOrUpdateDocumentStyle(Word.Document doc, StyleDefinitionRequest definition)
        {
            if (doc == null || definition == null || string.IsNullOrWhiteSpace(definition.StyleName))
            {
                return;
            }

            Word.Style style = null;
            try
            {
                object key = definition.StyleName;
                style = doc.Styles[key];
            }
            catch
            {
            }

            if (style == null)
            {
                style = doc.Styles.Add(definition.StyleName, Word.WdStyleType.wdStyleTypeParagraph);
            }

            style.Font.NameFarEast = definition.FontName;
            style.Font.Name = definition.FontName;
            style.Font.Size = definition.FontSize;
            style.Font.Bold = definition.Bold ? -1 : 0;
            style.ParagraphFormat.Alignment = ToWordParagraphAlignment(definition.Alignment);
            style.ParagraphFormat.SpaceBefore = 0f;
            style.ParagraphFormat.SpaceAfter = 0f;
            if (Math.Abs(definition.LineSpacing - 1f) < 0.01f)
            {
                style.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            }
            else
            {
                style.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                style.ParagraphFormat.LineSpacing = definition.LineSpacing;
            }

            int outlineLevel = definition.OutlineLevel > 0 ? definition.OutlineLevel : definition.Level;
            style.ParagraphFormat.OutlineLevel = outlineLevel == 10
                ? Word.WdOutlineLevel.wdOutlineLevelBodyText
                : (Word.WdOutlineLevel)Math.Max(1, Math.Min(9, outlineLevel));
            TrySetComProperty(style, "QuickStyle", true);
            TrySetComProperty(style, "Visibility", false);
            TrySetComProperty(style, "UnhideWhenUsed", true);
        }

        private static bool DocumentContainsStyle(Word.Document doc, string styleName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(styleName))
            {
                return false;
            }

            try
            {
                return doc.Styles[styleName] != null;
            }
            catch
            {
                return false;
            }
        }

        private static Word.WdParagraphAlignment ToWordParagraphAlignment(int alignment)
        {
            switch (alignment)
            {
                case 1:
                    return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case 2:
                    return Word.WdParagraphAlignment.wdAlignParagraphRight;
                case 3:
                    return Word.WdParagraphAlignment.wdAlignParagraphJustify;
                default:
                    return Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
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

        private static string GetLinkedStyleForOutlineLevel(int level)
        {
            return outlineLevelStyleLinks.TryGetValue(level, out string styleName)
                ? styleName
                : string.Empty;
        }

        private static int GetOutlineLevelForLinkedStyle(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
            {
                return 0;
            }

            foreach (KeyValuePair<int, string> item in outlineLevelStyleLinks)
            {
                if (string.Equals(item.Value, styleName, StringComparison.CurrentCultureIgnoreCase))
                {
                    return item.Key;
                }
            }

            return 0;
        }

        private void ApplyStyleLinksToCurrentDocument(Word.Document doc)
        {
            if (doc == null || outlineLevelStyleLinks.Count == 0)
            {
                return;
            }

            Word.Application app = Globals.ThisAddIn?.Application;
            HashSet<int> updatedStarts = new HashSet<int>();
            List<Word.Paragraph> updatedHeadingParagraphs = new List<Word.Paragraph>();
            try
            {
                using (new WordPerformanceScope(app))
                {
                    ConfigureLinkedStyleOutlineLevels(doc);

                    HashSet<int> linkedOutlineLevels = new HashSet<int>(
                        outlineLevelStyleLinks.Keys.Where(level => level >= 1 && level <= 9));
                    foreach (Word.Range headingRange in CollectOutlineRangesByFind(
                        doc,
                        linkedOutlineLevels,
                        doc.Content.Start,
                        doc.Content.End))
                    {
                        Word.Paragraph paragraph = GetHostParagraph(headingRange);
                        if (paragraph?.Range == null || !updatedStarts.Add(paragraph.Range.Start))
                        {
                            continue;
                        }

                        ApplyLinkedStyleToParagraph(paragraph);
                        int level = GetParagraphOutlineLevel(paragraph);
                        if (level >= 1 && level <= 6)
                        {
                            updatedHeadingParagraphs.Add(paragraph);
                        }
                    }

                    if (updatedHeadingParagraphs.Count > 0)
                    {
                        AutoUpdateOutlineListForParagraphs(updatedHeadingParagraphs);
                    }
                }

                app?.ScreenRefresh();
                RefreshCurrentStyleIndicator();
                TryUpdateStatusBar(app, "样式绑定已更新");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式绑定失败: {ex.Message}", "文档不加班");
            }
        }

        private static void ConfigureLinkedStyleOutlineLevels(Word.Document doc)
        {
            if (doc == null)
            {
                return;
            }

            foreach (KeyValuePair<int, string> item in outlineLevelStyleLinks)
            {
                SetStyleOutlineLevel(doc, item.Value, item.Key);
            }
        }

        private static void SetStyleOutlineLevel(Word.Document doc, string styleName, int level)
        {
            if (doc == null || string.IsNullOrWhiteSpace(styleName))
            {
                return;
            }

            try
            {
                object key = styleName;
                Word.Style style = doc.Styles[key];
                if (style == null)
                {
                    return;
                }

                style.ParagraphFormat.OutlineLevel = level == 10
                    ? Word.WdOutlineLevel.wdOutlineLevelBodyText
                    : (Word.WdOutlineLevel)level;
            }
            catch
            {
            }
        }

        private static void ApplyLinkedOutlineLevelToSelection(Word.Selection selection, string styleName)
        {
            int level = GetOutlineLevelForLinkedStyle(styleName);
            if ((level < 1 || level > 9) && level != 10)
            {
                return;
            }

            Word.WdOutlineLevel outlineLevel = level == 10
                ? Word.WdOutlineLevel.wdOutlineLevelBodyText
                : (Word.WdOutlineLevel)level;

            try
            {
                Word.Paragraphs paragraphs = selection?.Range?.Paragraphs;
                if (paragraphs == null)
                {
                    return;
                }

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
            }
            catch
            {
            }
        }

        private static void ApplyLinkedStyleToParagraph(Word.Paragraph paragraph)
        {
            if (paragraph?.Range == null)
            {
                return;
            }

            int level = GetParagraphOutlineLevel(paragraph);
            if (level == 0)
            {
                try
                {
                    level = paragraph.OutlineLevel == Word.WdOutlineLevel.wdOutlineLevelBodyText ? 10 : 0;
                }
                catch
                {
                    level = 0;
                }
            }

            string linkedStyle = GetLinkedStyleForOutlineLevel(level);
            if (string.IsNullOrWhiteSpace(linkedStyle))
            {
                return;
            }

            Word.WdOutlineLevel outlineLevel = level == 10
                ? Word.WdOutlineLevel.wdOutlineLevelBodyText
                : (Word.WdOutlineLevel)level;

            if (TrySetStyle(paragraph.Range, linkedStyle))
            {
                paragraph.OutlineLevel = outlineLevel;
                paragraph.Range.ParagraphFormat.OutlineLevel = outlineLevel;
            }
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

        internal static void ResetOperationCancellation()
        {
            operationCancelRequested = false;
            EnsureStopShortcutHook();
        }

        internal static void RequestOperationCancellation()
        {
            operationCancelRequested = true;
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
