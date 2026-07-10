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
        private bool updatingStyleGallery;
        private bool updatingOutlineLevel;
        private string styleGalleryDocumentKey;
        private readonly Dictionary<string, string> styleGalleryStyleNames =
            new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
        private static readonly Dictionary<int, string> outlineLevelStyleLinks = new Dictionary<int, string>();
        private static Dictionary<int, StyleDefinitionRequest> styleDefinitions;
        private static OutlineNumberPattern outlineNumberPattern = OutlineNumberPattern.Decimal;
        private static int outlineNumberTextSpacing = 1;
        private static volatile bool operationCancelRequested;
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

            RegisterInstance();
            EnsureStopShortcutHook();
            ApplyFeatureAvailability();
            UpdateHelpVersionLabel();
            InitializeRibbonToolTips();
            InitializeOutlineLevelDropDown();
            button9.Click += button9_Click;
            button10.Click += button10_Click;
            button11.Click += button11_Click;
            button12.Click += button12_Click;

            InitializeStyleGalleriesLightweight();
            // 刷新按钮高亮状态
            RefreshCurrentStyleIndicator();
        }

        private void InitializeRibbonToolTips()
        {
            SetTip(btnBatchReplace, "批量替换", "按规则批量查找并替换多个文档中的内容。");
            SetTip(btnStyleBrush, "格式刷", "调用 Word 原生格式刷，操作逻辑与 Word 自带格式刷一致。");
            SetTip(btnSwitchWindows, "切换窗口", "调用 Word 原生“视图 > 切换窗口”功能。");

            SetTip(styleGalleryDropDown, "当前样式", "显示当前文档中的样式，选择后应用到当前选区。");
            SetTip(outlineLevelDropDown, "大纲级别", "显示当前段落的大纲级别，选择后应用到当前选区。");
            SetTip(btnStyleBinding, "样式绑定", "设置大纲级别、样式和多级列表的绑定关系。");
            SetTip(btnCreateCustomStyles, "创建自定义样式", "配置并在当前文档中创建“通用1级标题”到“正文”等常用样式。");

            SetTip(button14, "插入图片题注", "在当前光标位置插入“图+自动编号域”。");
            SetTip(button13, "插入表格题注", "在当前光标位置插入表格题注。");
            SetTip(splitButtonReferenceCaption, "引用题注", "默认引用下一个题注，也可从下拉菜单选择上一个或自定义题注。");
            SetTip(button31, "引用自定义题注", "从当前文档题注列表中选择一个题注并插入动态引用。");
            SetTip(button28, "引用上一个题注", "在当前位置插入上一个题注的动态引用。");
            SetTip(button29, "引用下一个题注", "在当前位置插入下一个题注的动态引用。");

            SetTip(splitButton2, "选择", "选择文档中的指定对象。");
            SetTip(splitButtonClean, "清理", "集中执行格式、手工编号和空白页清理。");
            SetTip(btnInsertTotalPages, "插入总页码", "在当前位置插入文档总页码字段。");
            SetTip(button8, "插入编号", "调用 Word 原生编号列表，回车后自动继续下一编号。");
            SetTip(btnApplyHeitiXiaosi, "黑体小四", "将当前选区或输入点设置为黑体、小四字号。");
            SetTip(btnApplySongtiXiaosi, "宋体小四", "将当前选区或输入点设置为宋体、小四字号。");
            SetTip(button32, "图片单倍行距", "将当前文档所有图片所在段落设置为单倍行距。");
            SetTip(btnClearFormatting, "一键清除格式", "将当前选区段落恢复为普通正文文本。");
            SetTip(btnClearManualHeadingNumbers, "清除标题前的手工编号", "清除当前文档所有标题段落前的手工阿拉伯数字编号。");
            SetTip(btnCleanBlankPages, "清理空白页", "删除没有文字、表格或图片的空白页面。");
            SetTip(splitButtonUpdate, "更新", "集中执行目录、总页码、题注和章节号更新。");
            SetTip(button26, "更新目录", "重新更新当前文档中的目录内容和页码。");
            SetTip(btnUpdateCaptions, "更新题注", "同步更新当前文档中的图片题注和表格题注。");
            SetTip(btnUpdateOutlineList, "更新所选章节号", "更新光标所在段落、当前选区或通过选择器选中的标题段落章节号。");
            SetTip(button7, "更新总页码", "更新文档中的总页码字段。");

            SetTip(splitButton4, "窗格显示", "打开书签、题注或标识窗格。");
            SetTip(button9, "书签窗格", "显示当前文档书签列表并支持定位。");
            SetTip(button10, "题注窗格", "显示图注/表注列表并支持定位。");
            SetTip(button11, "标识窗格", "按文档类型显示标识列表并支持定位。");

            SetTip(button12, "需求追踪", "打开需求追踪控制台并建立映射关系。");
            SetTip(btnRequirementExtraction, "需求提取", "提取当前文档中的需求名称、需求标识和章节号，也可从选区手动添加。");

            SetTip(chkNonBodyBlankLine, "章节标题为空", "检查大纲级别 1-9 的章节标题是否没有文字或只有手工编号。");
            SetTip(chkCaptionContinuity, "题注连续性", "检查图题注和表题注编号是否连续。");
            SetTip(chkListContinuity, "多级列表连续性", "检查标题多级列表编号是否按大纲级别连续。");
            SetTip(chkBrokenReferences, "未更新域", "检查“错误！未找到引用源”等失效引用结果。");
            SetTip(btnStartDocumentCheck, "开始检查", "按勾选项扫描当前文档，并在右侧检查结果窗格中逐个跳转问题位置。");
            SetTip(btnSoftwareDocumentCheck, "软件文档检查", "识别软件需求规格说明或软件设计说明，并按 GJB 438C 检查明确章节标题是否齐全。");

            SetTip(menuHelp, "关于", "查看插件版本、更新内容、作者并打开本地帮助文档。");
            SetTip(btnHelpVersion, "插件版本号", "点击查看当前版本更新内容。");
            SetTip(btnOpenHelpDocument, "打开帮助文档", "打开本机随插件安装的功能说明和使用步骤。");
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
                if (!RequirementTrackingEnabled)
                {
                    SetTip(button12, "需求追踪", "功能暂未开放，当前不可用。");
                }
            }
        }

        private static bool IsCurrentDocumentSoftwareDocument()
        {
            try
            {
                Word.Document doc = Globals.ThisAddIn?.Application?.ActiveDocument;
                string name = ((doc?.Name ?? string.Empty) + " " + (doc?.FullName ?? string.Empty)).Trim();
                return name.IndexOf("软件", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static void SetTip(RibbonButton control, string screenTip, string superTip)
        {
            if (control == null) return;
            control.ScreenTip = screenTip;
            control.SuperTip = superTip;
        }

        private static void SetTip(RibbonToggleButton control, string screenTip, string superTip)
        {
            if (control == null) return;
            control.ScreenTip = screenTip;
            control.SuperTip = superTip;
        }

        private static void SetTip(RibbonSplitButton control, string screenTip, string superTip)
        {
            if (control == null) return;
            control.ScreenTip = screenTip;
            control.SuperTip = superTip;
        }

        private static void SetTip(RibbonDropDown control, string screenTip, string superTip)
        {
            if (control == null) return;
            control.ScreenTip = screenTip;
            control.SuperTip = superTip;
        }

        private static void SetTip(RibbonMenu control, string screenTip, string superTip)
        {
            if (control == null) return;
            control.ScreenTip = screenTip;
            control.SuperTip = superTip;
        }

        private static void SetTip(RibbonCheckBox control, string screenTip, string superTip)
        {
            if (control == null) return;
            control.ScreenTip = screenTip;
            control.SuperTip = superTip;
        }

        internal static void RefreshAllStyleIndicators()
        {
            foreach (Ribbon1 ribbon in LoadedInstances.ToArray())
            {
                ribbon?.RefreshCurrentStyleIndicator();
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

        // 刷新当前样式下拉框，保持和 Word 当前光标样式一致。
        internal void RefreshCurrentStyleIndicator()
        {
            string currentDocumentKey = GetActiveDocumentKey();
            if (styleGalleryDropDown == null || styleGalleryDropDown.Items.Count == 0)
            {
                InitializeStyleGalleriesLightweight();
            }

            if (!string.Equals(styleGalleryDocumentKey, currentDocumentKey, StringComparison.OrdinalIgnoreCase))
            {
                InitializeStyleGalleriesLightweight();
            }

            if (styleGalleryDropDown == null || styleGalleryDropDown.Items.Count == 0)
            {
                SelectOutlineLevelItem(GetCurrentSelectionOutlineLevel());
                return;
            }

            string currentStyleName = GetCurrentSelectionStyleName();
            SelectStyleGalleryItem(currentStyleName);
            SelectOutlineLevelItem(GetCurrentSelectionOutlineLevel());
        }

        private void InitializeStyleGalleriesLightweight()
        {
            RefreshDocumentStyleGallery();
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
                SelectOutlineLevelItem(GetCurrentSelectionOutlineLevel());
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

        private void RefreshDocumentStyleGallery()
        {
            if (styleGalleryDropDown == null)
            {
                return;
            }

            updatingStyleGallery = true;
            try
            {
                styleGalleryDropDown.Items.Clear();
                styleGalleryStyleNames.Clear();

                Word.Document doc = Globals.ThisAddIn?.Application?.ActiveDocument;
                if (doc == null)
                {
                    styleGalleryDocumentKey = string.Empty;
                    AddStyleDropDownItem("无活动文档", null);
                    return;
                }

                styleGalleryDocumentKey = GetDocumentKey(doc);
                string currentStyle = GetCurrentSelectionStyleName();
                RibbonDropDownItem currentItem = AddStyleDropDownItem(BuildCurrentStyleGalleryLabel(currentStyle), null);
                HashSet<string> usedLabels = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
                foreach (string styleName in GetDocumentStyleNames(doc))
                {
                    string displayName = GetUniqueStyleGalleryLabel(FormatStyleGalleryLabel(styleName), usedLabels);
                    AddStyleDropDownItem(displayName, styleName);
                }

                if (styleGalleryDropDown.Items.Count == 1)
                {
                    AddStyleDropDownItem("未找到样式", null);
                }

                styleGalleryDropDown.SelectedItem = currentItem;
            }
            catch
            {
            }
            finally
            {
                updatingStyleGallery = false;
            }
        }

        private void SelectStyleGalleryItem(string styleName)
        {
            if (styleGalleryDropDown == null || styleGalleryDropDown.Items.Count == 0)
            {
                return;
            }

            updatingStyleGallery = true;
            try
            {
                RibbonDropDownItem currentItem = styleGalleryDropDown.Items[0];
                currentItem.Label = BuildCurrentStyleGalleryLabel(styleName);
                currentItem.ScreenTip = styleName ?? string.Empty;
                currentItem.SuperTip = styleName ?? string.Empty;
                styleGalleryDropDown.SelectedItem = currentItem;
            }
            catch
            {
            }
            finally
            {
                updatingStyleGallery = false;
            }
        }

        private static string GetActiveDocumentKey()
        {
            try
            {
                return GetDocumentKey(Globals.ThisAddIn?.Application?.ActiveDocument);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetDocumentKey(Word.Document doc)
        {
            if (doc == null)
            {
                return string.Empty;
            }

            try
            {
                return string.IsNullOrWhiteSpace(doc.FullName) ? doc.Name : doc.FullName;
            }
            catch
            {
                return string.Empty;
            }
        }

        private RibbonDropDownItem AddStyleDropDownItem(string label, string fullStyleName)
        {
            RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
            item.Label = label;
            item.ScreenTip = fullStyleName ?? label;
            item.SuperTip = fullStyleName ?? label;
            styleGalleryDropDown.Items.Add(item);
            if (!string.IsNullOrWhiteSpace(fullStyleName))
            {
                styleGalleryStyleNames[label] = fullStyleName;
            }
            return item;
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

        private static string GetUniqueStyleGalleryLabel(string label, HashSet<string> usedLabels)
        {
            if (usedLabels == null || usedLabels.Add(label))
            {
                return label;
            }

            for (int i = 2; i < 1000; i++)
            {
                string suffix = " " + i;
                string candidate = label.Length + suffix.Length <= 24
                    ? label + suffix
                    : label.Substring(0, Math.Max(0, 24 - suffix.Length)) + suffix;
                if (usedLabels.Add(candidate))
                {
                    return candidate;
                }
            }

            return label;
        }

        private string ResolveStyleGalleryStyleName(string label)
        {
            return styleGalleryStyleNames.TryGetValue(label ?? string.Empty, out string fullName)
                ? fullName
                : label;
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
                        string name = styles[i]?.NameLocal;
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
        private string GetCurrentSelectionStyleName()
        {
            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Selection selection = app?.Selection;
                Word.Range range = selection?.Range;
                if (range == null)
                    return string.Empty;

                object styleObj = TryGetParagraphStyle(range)
                    ?? TryGetStyle(range)
                    ?? TryGetStyle(selection);

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

        private void styleGalleryDropDown_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (updatingStyleGallery || styleGalleryDropDown?.SelectedItem == null)
            {
                return;
            }

            string styleName = ResolveStyleGalleryStyleName(styleGalleryDropDown.SelectedItem.Label);
            if (string.IsNullOrWhiteSpace(styleName)
                || styleGalleryDropDown.SelectedItem == styleGalleryDropDown.Items[0]
                || styleName == "无活动文档"
                || styleName == "未找到样式")
            {
                return;
            }

            try
            {
                Word.Application app = Globals.ThisAddIn.Application;
                Word.Selection selection = app?.Selection;
                if (selection == null)
                {
                    return;
                }

                using (new WordPerformanceScope(app))
                {
                    if (!TrySetStyle(selection.Range, styleName))
                    {
                        TrySetStyle(selection, styleName);
                    }

                    ApplyLinkedOutlineLevelToSelection(selection, styleName);
                }

                SelectStyleGalleryItem(styleName);
                TryUpdateStatusBar(app, styleName);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式失败: {ex.Message}", "文档不加班");
            }
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
                    () => GetDocumentStyleNames(activeDoc),
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
                .Concat(new[] { GetCurrentSelectionStyleName() })
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
                    if (definitionsToCreate.Count == 0)
                    {
                        MessageBox.Show("未勾选需要创建的样式。", "文档不加班");
                        return;
                    }

                    using (new WordPerformanceScope(app))
                    {
                        foreach (StyleDefinitionRequest definition in definitionsToCreate)
                        {
                            CreateOrUpdateDocumentStyle(doc, definition);
                        }
                    }

                    RefreshDocumentStyleGallery();
                    TryUpdateStatusBar(app, "自定义样式已创建");
                    MessageBox.Show("自定义样式已创建到当前文档。", "文档不加班");
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
            if (RequirementTrackingEnabled)
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
            else
            {
                MessageBox.Show("需求追踪功能暂未开放，当前不可用。", "文档不加班");
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
