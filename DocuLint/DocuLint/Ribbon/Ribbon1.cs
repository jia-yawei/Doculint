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

        private sealed class WordWindowPickerForm : Form
        {
            internal IntPtr SelectedHandle { get; private set; }

            internal bool CloseRequested { get; private set; }

            internal WordWindowPickerForm(IReadOnlyList<WordWindowItem> windows)
            {
                Text = "切换窗口";
                ShowInTaskbar = false;
                FormBorderStyle = FormBorderStyle.FixedSingle;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.Manual;
                Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);
                BackColor = System.Drawing.Color.White;
                AutoScaleMode = AutoScaleMode.Dpi;

                int contentWidth = 420;
                using (System.Drawing.Graphics graphics = CreateGraphics())
                {
                    foreach (WordWindowItem item in windows)
                    {
                        int measured = (int)Math.Ceiling(graphics.MeasureString(item.Title ?? string.Empty, Font).Width) + 78;
                        contentWidth = Math.Max(contentWidth, Math.Min(measured, 820));
                    }
                }

                TableLayoutPanel list = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    RowCount = windows.Count,
                    AutoScroll = windows.Count > 10,
                    Padding = new Padding(6),
                    BackColor = System.Drawing.Color.White
                };

                for (int i = 0; i < windows.Count; i++)
                {
                    WordWindowItem window = windows[i];
                    list.RowStyles.Add(new RowStyle(SizeType.Absolute, 38F));
                    TableLayoutPanel row = new TableLayoutPanel
                    {
                        Dock = DockStyle.Top,
                        Height = 38,
                        ColumnCount = 2,
                        Margin = new Padding(0),
                        BackColor = i % 2 == 0
                            ? System.Drawing.Color.White
                            : System.Drawing.Color.FromArgb(248, 249, 251)
                    };
                    row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 38F));
                    row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                    Button closeButton = new Button
                    {
                        Dock = DockStyle.Fill,
                        Text = "×",
                        FlatStyle = FlatStyle.Flat,
                        Margin = new Padding(2),
                        ForeColor = System.Drawing.Color.FromArgb(90, 96, 106),
                        BackColor = System.Drawing.Color.Transparent,
                        Cursor = Cursors.Hand,
                        TabStop = false
                    };
                    closeButton.FlatAppearance.BorderSize = 0;
                    closeButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(250, 226, 226);
                    closeButton.Click += (_, __) => SelectWindow(window.Handle, true);

                    Button documentButton = new Button
                    {
                        Dock = DockStyle.Fill,
                        Text = (i + 1) + " " + window.Title,
                        TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                        AutoEllipsis = true,
                        FlatStyle = FlatStyle.Flat,
                        Margin = new Padding(0, 2, 2, 2),
                        BackColor = System.Drawing.Color.Transparent,
                        Cursor = Cursors.Hand
                    };
                    documentButton.FlatAppearance.BorderSize = 0;
                    documentButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(232, 241, 253);
                    documentButton.Click += (_, __) => SelectWindow(window.Handle, false);

                    row.Controls.Add(closeButton, 0, 0);
                    row.Controls.Add(documentButton, 1, 0);
                    list.Controls.Add(row, 0, i);
                }

                ClientSize = new System.Drawing.Size(contentWidth, Math.Min(windows.Count * 38 + 12, 400));
                Controls.Add(list);

                System.Drawing.Rectangle workArea = Screen.FromPoint(Cursor.Position).WorkingArea;
                int x = Math.Min(Cursor.Position.X, workArea.Right - Width);
                int y = Math.Min(Cursor.Position.Y, workArea.Bottom - Height);
                Location = new System.Drawing.Point(Math.Max(workArea.Left, x), Math.Max(workArea.Top, y));
            }

            private void SelectWindow(IntPtr handle, bool closeRequested)
            {
                SelectedHandle = handle;
                CloseRequested = closeRequested;
                DialogResult = DialogResult.OK;
                Close();
            }
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

        private void btnSwitchWindows_Click(object sender, RibbonControlEventArgs e)
        {
            List<WordWindowItem> windows = GetOpenWordWindows();
            if (windows.Count == 0)
            {
                MessageBox.Show("当前没有打开的 Word 文档。", "文档管理");
                return;
            }

            using (WordWindowPickerForm picker = new WordWindowPickerForm(windows))
            {
                if (picker.ShowDialog() != DialogResult.OK || picker.SelectedHandle == IntPtr.Zero)
                {
                    return;
                }

                if (picker.CloseRequested)
                {
                    CloseWordDocumentWindow(picker.SelectedHandle);
                }
                else
                {
                    ActivateWordDocumentWindow(picker.SelectedHandle);
                }
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

        // 刷新当前样式显示，保持和 Word 当前光标所在段落样式一致。
        internal void RefreshCurrentStyleIndicator()
        {
            if (styleGalleryDropDown == null)
            {
                return;
            }

            string currentStyleName = GetCurrentParagraphStyleName();
            styleGalleryDropDown.Label = "当前样式：" + BuildCurrentStyleGalleryLabel(currentStyleName);
            styleGalleryDropDown.ScreenTip = "当前样式";
            styleGalleryDropDown.SuperTip = string.IsNullOrWhiteSpace(currentStyleName)
                ? "当前段落未读取到样式。"
                : currentStyleName.Trim();
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
            const int maxDisplayWidth = 20;
            const string ellipsis = "...";
            if (string.IsNullOrWhiteSpace(styleName))
            {
                return styleName;
            }

            string normalizedName = styleName.Trim();
            int displayWidth = 0;
            int endIndex = 0;
            while (endIndex < normalizedName.Length)
            {
                int characterWidth = normalizedName[endIndex] <= 0x7f ? 1 : 2;
                if (displayWidth + characterWidth > maxDisplayWidth)
                {
                    break;
                }

                displayWidth += characterWidth;
                endIndex++;
            }

            if (endIndex == normalizedName.Length)
            {
                return normalizedName;
            }

            int contentWidth = maxDisplayWidth - ellipsis.Length;
            displayWidth = 0;
            endIndex = 0;
            while (endIndex < normalizedName.Length)
            {
                int characterWidth = normalizedName[endIndex] <= 0x7f ? 1 : 2;
                if (displayWidth + characterWidth > contentWidth)
                {
                    break;
                }

                displayWidth += characterWidth;
                endIndex++;
            }

            return normalizedName.Substring(0, endIndex) + ellipsis;
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
