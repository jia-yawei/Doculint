using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    // Word 工具栏（Ribbon）功能类：快速应用文档样式
    public partial class Ribbon1
    {
        private static readonly List<Ribbon1> LoadedInstances = new List<Ribbon1>();
        internal static bool RequirementTrackingEnabled => false;
        private static readonly Dictionary<string, Dictionary<int, int>> ManualHeaderRowOverridesByDocument =
            new Dictionary<string, Dictionary<int, int>>(StringComparer.OrdinalIgnoreCase);
        private RibbonButton rememberedTableSelectionButton;
        private Action<object, RibbonControlEventArgs> rememberedTableSelectionAction;
        private RibbonButton rememberedInsertItemButton;
        private Action<object, RibbonControlEventArgs> rememberedInsertItemAction;
        private sealed class WordPerformanceScope : IDisposable
        {
            private readonly Word.Application app;
            private readonly object previousScreenUpdating;
            private readonly object previousDisplayAlerts;
            private readonly object previousPagination;

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

                TrySetComProperty(app, "ScreenUpdating", false);
                TrySetComProperty(app, "DisplayAlerts", Word.WdAlertLevel.wdAlertsNone);
                TrySetComProperty(options, "Pagination", false);
            }

            public void Dispose()
            {
                if (app == null)
                {
                    return;
                }

                object options = TryGetComProperty(app, "Options");
                TryRestoreComProperty(options, "Pagination", previousPagination);
                TryRestoreComProperty(app, "DisplayAlerts", previousDisplayAlerts);
                TryRestoreComProperty(app, "ScreenUpdating", previousScreenUpdating);
            }
        }

        // 定义5个常用样式名称（可直接改这里）
        private readonly string[] myStyles = new string[]
        {
            "1级标题-105模板",
            "2级标题-105模板",
            "3级标题-105模板",
            "正文-105",
            "题注"
        };

        // 工具栏加载时：初始化按钮文字 + 刷新样式选中状态
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            RegisterInstance();
            InitializeDocumentGroupMenu();
            ApplyFeatureAvailability();
            InitializeRibbonToolTips();
            button9.Click += button9_Click;
            button10.Click += button10_Click;
            button11.Click += button11_Click;
            button12.Click += button12_Click;
            splitButton2.Click += splitButton2_ButtonClick;
            splitButton6.Click += splitButton6_ButtonClick;
            InitializeSplitButtonPrimaryActions();

            // 给5个按钮设置显示文字
            if (myStyles.Length >= 5)
            {
                btnStyle1.Label = myStyles[0];
                btnStyle2.Label = myStyles[1];
                btnStyle3.Label = myStyles[2];
                btnStyle4.Label = myStyles[3];
                btnStyle5.Label = myStyles[4];
            }

            // 刷新按钮高亮状态
            RefreshCurrentStyleIndicator();
        }

        private void InitializeRibbonToolTips()
        {
            SetTip(splitButton1, "文档组操作", "管理文档组、快速加入当前文档并维护活动文档组。");
            SetTip(button27, "当前活动组", "显示当前生效的活动文档组。");
            SetTip(btnBatchReplace, "批量替换", "按规则批量查找并替换多个文档中的内容。");
            SetTip(button2, "抓取管理", "查看、重命名和删除当前文档组中的抓取内容。");
            SetTip(button3, "内容抓取", "将当前选区按原格式抓取到当前文档组。");
            SetTip(button30, "内容注入", "从当前文档组选择抓取内容并注入到光标处。");

            SetTip(btnStyle1, "应用样式", "将当前选区应用为“样式1”。");
            SetTip(btnStyle2, "应用样式", "将当前选区应用为“样式2”。");
            SetTip(btnStyle3, "应用样式", "将当前选区应用为“样式3”。");
            SetTip(btnStyle4, "应用样式", "将当前选区应用为“样式4”。");
            SetTip(btnStyle5, "应用样式", "将当前选区应用为“题注”样式。");

            SetTip(button14, "插入图片题注", "在当前光标位置插入“图+自动编号域”。");
            SetTip(button13, "插入表格题注", "在当前光标位置插入表格题注。");
            SetTip(button17, "更新图片题注", "按规则刷新当前文档中的图片题注。");
            SetTip(button16, "更新表格题注", "按规则刷新当前文档中的表格题注。");
            SetTip(button28, "引用上一个题注", "在当前位置插入上一个题注的动态引用。");
            SetTip(button29, "引用下一个题注", "在当前位置插入下一个题注的动态引用。");

            SetTip(splitButton6, "插入项", "插入总页码或在表格中执行自动序号。");
            SetTip(button8, "自动序号", "在当前表格列中按列表编号自动向下填充序号。");
            SetTip(btnRebuildOutlineList, "自动章节号", "按标题结构重建章节号。");
            SetTip(button7, "更新总页码", "更新文档中的总页码字段。");

            SetTip(splitButton4, "窗格显示", "打开书签、题注或标识窗格。");
            SetTip(button9, "书签窗格", "显示当前文档书签列表并支持定位。");
            SetTip(button10, "题注窗格", "显示图注/表注列表并支持定位。");
            SetTip(button11, "标识窗格", "按文档类型显示标识列表并支持定位。");

            SetTip(button12, "需求追踪", "打开需求追踪控制台并建立映射关系。");
        }

        private void ApplyFeatureAvailability()
        {
            if (button12 == null)
            {
                return;
            }

            button12.Enabled = RequirementTrackingEnabled;
            if (!RequirementTrackingEnabled)
            {
                SetTip(button12, "需求追踪", "功能暂未开放，当前不可用。");
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

        private void InitializeSplitButtonPrimaryActions()
        {
            RememberTableSelectionAction(button22, button22_Click);
            RememberInsertItemAction(btnInsertTotalPages, btnInsertTotalPages_Click);
        }

        private void splitButton2_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            rememberedTableSelectionAction?.Invoke(rememberedTableSelectionButton, e);
        }

        private void splitButton6_ButtonClick(object sender, RibbonControlEventArgs e)
        {
            rememberedInsertItemAction?.Invoke(rememberedInsertItemButton, e);
        }

        internal void RememberTableSelectionAction(RibbonButton button, Action<object, RibbonControlEventArgs> action)
        {
            rememberedTableSelectionButton = button;
            rememberedTableSelectionAction = action;
            ApplySplitButtonPrimaryVisual(splitButton2, button, "选择表格范围", "快速定位续表、跨页表格或全部表格。");
        }

        internal void RememberInsertItemAction(RibbonButton button, Action<object, RibbonControlEventArgs> action)
        {
            rememberedInsertItemButton = button;
            rememberedInsertItemAction = action;
            ApplySplitButtonPrimaryVisual(splitButton6, button, "插入项", "插入总页码或在表格中执行自动序号。");
        }

        private static void ApplySplitButtonPrimaryVisual(
            RibbonSplitButton splitButton,
            RibbonButton button,
            string defaultScreenTip,
            string defaultSuperTip)
        {
            if (splitButton == null || button == null)
            {
                return;
            }

            splitButton.Label = button.Label;
            splitButton.Image = button.Image;
            splitButton.ScreenTip = string.IsNullOrWhiteSpace(button.ScreenTip) ? defaultScreenTip : button.ScreenTip;
            splitButton.SuperTip = string.IsNullOrWhiteSpace(button.SuperTip) ? defaultSuperTip : button.SuperTip;
        }

        internal static void RefreshAllStyleIndicators()
        {
            foreach (Ribbon1 ribbon in LoadedInstances.ToArray())
            {
                ribbon?.RefreshCurrentStyleIndicator();
            }
        }

        // 刷新：哪个样式被选中，哪个按钮就高亮
        internal void RefreshCurrentStyleIndicator()
        {
            string currentStyleName = GetCurrentSelectionStyleName();

            SetStyleButtonChecked(btnStyle1, currentStyleName, GetStyleName(0));
            SetStyleButtonChecked(btnStyle2, currentStyleName, GetStyleName(1));
            SetStyleButtonChecked(btnStyle3, currentStyleName, GetStyleName(2));
            SetStyleButtonChecked(btnStyle4, currentStyleName, GetStyleName(3));
            SetStyleButtonChecked(btnStyle5, currentStyleName, GetStyleName(4));
        }

        // 根据编号获取样式名（防止越界报错）
        private string GetStyleName(int index)
        {
            if (index < 0 || index >= myStyles.Length)
                return string.Empty;

            return myStyles[index] ?? string.Empty;
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

        // 设置按钮高亮：当前样式 = 按钮样式 → 选中
        private static void SetStyleButtonChecked(RibbonToggleButton button, string currentStyleName, string expectedStyle)
        {
            if (button == null) return;

            button.Checked = !string.IsNullOrWhiteSpace(expectedStyle)
                && string.Equals(currentStyleName, expectedStyle, StringComparison.OrdinalIgnoreCase);
        }

        // 给Word选中文字应用指定样式
        private void ApplyMyStyle(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName)) return;

            try
            {
                Word.Selection selection = Globals.ThisAddIn.Application?.Selection;
                Word.Range range = selection?.Range;
                if (range == null) return;

                bool applied = TrySetStyle(range, styleName);
                if (!applied)
                {
                    applied = TrySetStyle(selection, styleName);
                }

                if (!applied)
                {
                    throw new InvalidOperationException("当前 Word 选区不支持设置该样式。");
                }

                RefreshAllStyleIndicators();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用样式失败: {ex.Message}", "文档不加班");
            }
        }

        private void btnStyle1_Click_1(object sender, RibbonControlEventArgs e)
        {
            ApplyMyStyle(GetStyleName(0));
        }

        private void btnStyle2_Click_1(object sender, RibbonControlEventArgs e)
        {
            ApplyMyStyle(GetStyleName(1));
        }

        private void btnStyle3_Click_1(object sender, RibbonControlEventArgs e)
        {
            ApplyMyStyle(GetStyleName(2));
        }

        private void btnStyle4_Click_1(object sender, RibbonControlEventArgs e)
        {
            ApplyMyStyle(GetStyleName(3));
        }

        private void button24_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteMergeContinuationTableAction();
        }

        private void group2_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            ShowTablesAndFiguresFormattingSettingsDialog();
        }

        private void button21_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteNormalizeTablesAction();
        }

        private void button25_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteNormalizeSelectedTableAction();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteNormalizeAllImagesAction();
        }

        private void button17_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteRefreshImageCaptions();
        }

        private void button14_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertImageCaption();
        }

        private void button13_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteInsertTableCaption();
        }

        private void button16_Click(object sender, RibbonControlEventArgs e)
        {
            ExecuteRefreshTableCaptions();
        }

        private void btnStyle5_Click_1(object sender, RibbonControlEventArgs e)
        {
            ApplyMyStyle(GetStyleName(4));
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            GroupCapturedContentService.OpenCaptureManager(Globals.ThisAddIn.Application, null);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            GroupCapturedContentService.CaptureCurrentSelection(Globals.ThisAddIn.Application, null);
        }

        private void button30_Click(object sender, RibbonControlEventArgs e)
        {
            GroupCapturedContentService.InjectCapturedContent(Globals.ThisAddIn.Application, null);
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

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {

        }

        private void button12_Click_1(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
