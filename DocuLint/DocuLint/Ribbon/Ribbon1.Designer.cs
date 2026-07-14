namespace DocuLint
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            if (disposing)
            {
                DisposeRuntimeResources();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupDocumentManage = this.Factory.CreateRibbonGroup();
            this.btnSwitchWindows = this.Factory.CreateRibbonMenu();
            this.btnOpenCurrentFolder = this.Factory.CreateRibbonButton();
            this.btnSaveAllDocuments = this.Factory.CreateRibbonButton();
            this.btnCloseOtherDocuments = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnBatchReplace = this.Factory.CreateRibbonButton();
            this.btnStyleBrush = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.styleGalleryDropDown = this.Factory.CreateRibbonDropDown();
            this.btnCreateCustomStyles = this.Factory.CreateRibbonButton();
            this.outlineLevelDropDown = this.Factory.CreateRibbonDropDown();
            this.btnStyleBinding = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.splitButton2 = this.Factory.CreateRibbonSplitButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.btnInsertTotalPages = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.btnApplyHeitiXiaosi = this.Factory.CreateRibbonButton();
            this.btnApplySongtiXiaosi = this.Factory.CreateRibbonButton();
            this.btnTogglePageWhitespace = this.Factory.CreateRibbonToggleButton();
            this.button32 = this.Factory.CreateRibbonButton();
            this.splitButtonClean = this.Factory.CreateRibbonSplitButton();
            this.btnClearFormatting = this.Factory.CreateRibbonButton();
            this.btnClearManualHeadingNumbers = this.Factory.CreateRibbonButton();
            this.btnCleanBlankPages = this.Factory.CreateRibbonButton();
            this.splitButtonUpdate = this.Factory.CreateRibbonSplitButton();
            this.button26 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.btnUpdateCaptions = this.Factory.CreateRibbonButton();
            this.btnUpdateOutlineList = this.Factory.CreateRibbonButton();
            this.button18 = this.Factory.CreateRibbonButton();
            this.button25 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.splitButtonReferenceCaption = this.Factory.CreateRibbonSplitButton();
            this.button29 = this.Factory.CreateRibbonButton();
            this.button28 = this.Factory.CreateRibbonButton();
            this.button31 = this.Factory.CreateRibbonButton();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.splitButton4 = this.Factory.CreateRibbonSplitButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.btnRequirementExtraction = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.groupDocumentCheck = this.Factory.CreateRibbonGroup();
            this.chkNonBodyBlankLine = this.Factory.CreateRibbonCheckBox();
            this.chkCaptionContinuity = this.Factory.CreateRibbonCheckBox();
            this.chkListContinuity = this.Factory.CreateRibbonCheckBox();
            this.chkStyleConsistency = this.Factory.CreateRibbonCheckBox();
            this.chkBrokenReferences = this.Factory.CreateRibbonCheckBox();
            this.btnStartDocumentCheck = this.Factory.CreateRibbonButton();
            this.groupSoftwareTools = this.Factory.CreateRibbonGroup();
            this.btnSoftwareDocumentCheck = this.Factory.CreateRibbonButton();
            this.group8 = this.Factory.CreateRibbonGroup();
            this.menuHelp = this.Factory.CreateRibbonMenu();
            this.btnHelpVersion = this.Factory.CreateRibbonButton();
            this.btnOpenHelpDocument = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupDocumentManage.SuspendLayout();
            this.group4.SuspendLayout();
            this.group1.SuspendLayout();
            this.group6.SuspendLayout();
            this.group3.SuspendLayout();
            this.group7.SuspendLayout();
            this.groupDocumentCheck.SuspendLayout();
            this.groupSoftwareTools.SuspendLayout();
            this.group8.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.groupDocumentManage);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Groups.Add(this.groupDocumentCheck);
            this.tab1.Groups.Add(this.groupSoftwareTools);
            this.tab1.Groups.Add(this.group8);
            this.tab1.Label = "搞快点";
            this.tab1.Name = "tab1";
            // 
            // groupDocumentManage
            // 
            this.groupDocumentManage.Items.Add(this.btnSwitchWindows);
            this.groupDocumentManage.Items.Add(this.btnOpenCurrentFolder);
            this.groupDocumentManage.Items.Add(this.btnSaveAllDocuments);
            this.groupDocumentManage.Items.Add(this.btnCloseOtherDocuments);
            this.groupDocumentManage.Label = "文档管理";
            this.groupDocumentManage.Name = "groupDocumentManage";
            // 
            // btnSwitchWindows
            // 
            this.btnSwitchWindows.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSwitchWindows.Dynamic = true;
            this.btnSwitchWindows.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnSwitchWindows.Label = "切换窗口";
            this.btnSwitchWindows.Name = "btnSwitchWindows";
            this.btnSwitchWindows.OfficeImageId = "WindowSwitchWindowsMenuWord";
            this.btnSwitchWindows.ShowImage = true;
            this.btnSwitchWindows.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSwitchWindows_ItemsLoading);
            // 
            // btnOpenCurrentFolder
            // 
            this.btnOpenCurrentFolder.Label = "打开所在文件夹";
            this.btnOpenCurrentFolder.Name = "btnOpenCurrentFolder";
            this.btnOpenCurrentFolder.OfficeImageId = "FileOpen";
            this.btnOpenCurrentFolder.ShowImage = true;
            this.btnOpenCurrentFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenCurrentFolder_Click);
            // 
            // btnSaveAllDocuments
            // 
            this.btnSaveAllDocuments.Label = "保存所有文档";
            this.btnSaveAllDocuments.Name = "btnSaveAllDocuments";
            this.btnSaveAllDocuments.OfficeImageId = "FileSaveAll";
            this.btnSaveAllDocuments.ShowImage = true;
            this.btnSaveAllDocuments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveAllDocuments_Click);
            // 
            // btnCloseOtherDocuments
            // 
            this.btnCloseOtherDocuments.Label = "关闭其他文档";
            this.btnCloseOtherDocuments.Name = "btnCloseOtherDocuments";
            this.btnCloseOtherDocuments.OfficeImageId = "WindowClose";
            this.btnCloseOtherDocuments.ShowImage = true;
            this.btnCloseOtherDocuments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCloseOtherDocuments_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnBatchReplace);
            this.group4.Items.Add(this.btnStyleBrush);
            this.group4.Label = "批量处理";
            this.group4.Name = "group4";
            // 
            // btnBatchReplace
            // 
            this.btnBatchReplace.Label = "批量替换";
            this.btnBatchReplace.Name = "btnBatchReplace";
            this.btnBatchReplace.OfficeImageId = "ReplaceDialog";
            this.btnBatchReplace.ShowImage = true;
            this.btnBatchReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBatchReplace_Click);
            // 
            // btnStyleBrush
            // 
            this.btnStyleBrush.Label = "格式刷";
            this.btnStyleBrush.Name = "btnStyleBrush";
            this.btnStyleBrush.OfficeImageId = "FormatPainter";
            this.btnStyleBrush.ShowImage = true;
            this.btnStyleBrush.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyleBrush_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.styleGalleryDropDown);
            this.group1.Items.Add(this.outlineLevelDropDown);
            this.group1.Items.Add(this.btnCreateCustomStyles);
            this.group1.Items.Add(this.btnStyleBinding);
            this.group1.Label = "样式管理";
            this.group1.Name = "group1";
            // 
            // styleGalleryDropDown
            // 
            this.styleGalleryDropDown.Label = "当前样式";
            this.styleGalleryDropDown.Name = "styleGalleryDropDown";
            this.styleGalleryDropDown.OfficeImageId = "GroupShapeStyles";
            this.styleGalleryDropDown.ShowImage = true;
            this.styleGalleryDropDown.SizeString = "000000000000";
            this.styleGalleryDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.styleGalleryDropDown_SelectionChanged);
            // 
            // btnCreateCustomStyles
            // 
            this.btnCreateCustomStyles.Label = "创建自定义样式";
            this.btnCreateCustomStyles.Name = "btnCreateCustomStyles";
            this.btnCreateCustomStyles.OfficeImageId = "GroupStyles";
            this.btnCreateCustomStyles.ShowImage = true;
            this.btnCreateCustomStyles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateCustomStyles_Click);
            // 
            // outlineLevelDropDown
            // 
            this.outlineLevelDropDown.Label = "大纲级别";
            this.outlineLevelDropDown.Name = "outlineLevelDropDown";
            this.outlineLevelDropDown.OfficeImageId = "GroupOutline";
            this.outlineLevelDropDown.ShowImage = true;
            this.outlineLevelDropDown.SizeString = "000000";
            this.outlineLevelDropDown.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.outlineLevelDropDown_SelectionChanged);
            // 
            // btnStyleBinding
            // 
            this.btnStyleBinding.Label = "样式绑定";
            this.btnStyleBinding.Name = "btnStyleBinding";
            this.btnStyleBinding.OfficeImageId = "ChangeBinding";
            this.btnStyleBinding.ShowImage = true;
            this.btnStyleBinding.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyleBinding_Click);
            // 
            // group6
            // 
            this.group6.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group6.Items.Add(this.splitButton2);
            this.group6.Items.Add(this.splitButtonUpdate);
            this.group6.Items.Add(this.btnInsertTotalPages);
            this.group6.Items.Add(this.button8);
            this.group6.Items.Add(this.btnApplyHeitiXiaosi);
            this.group6.Items.Add(this.btnApplySongtiXiaosi);
            this.group6.Items.Add(this.btnTogglePageWhitespace);
            this.group6.Items.Add(this.button32);
            this.group6.Items.Add(this.button25);
            this.group6.Items.Add(this.button18);
            this.group6.Items.Add(this.splitButtonClean);
            this.group6.Label = "快速工具";
            this.group6.Name = "group6";
            this.group6.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group2_DialogLauncherClick);
            // 
            // splitButton2
            // 
            this.splitButton2.Items.Add(this.button5);
            this.splitButton2.Label = "选择";
            this.splitButton2.Name = "splitButton2";
            this.splitButton2.OfficeImageId = "TableSelectMenuPowerPoint";
            // 
            // button5
            // 
            this.button5.Label = "下个跨页表格";
            this.button5.Name = "button5";
            this.button5.OfficeImageId = "TableSelectMenuPowerPoint";
            this.button5.ScreenTip = "选择下个跨页表格";
            this.button5.ShowImage = true;
            this.button5.SuperTip = "从当前光标位置开始，选中下一个跨页表格。";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // btnInsertTotalPages
            // 
            this.btnInsertTotalPages.Label = "插入总页码";
            this.btnInsertTotalPages.Name = "btnInsertTotalPages";
            this.btnInsertTotalPages.OfficeImageId = "CustomPageNumberTopGallery";
            this.btnInsertTotalPages.ScreenTip = "插入文档总页码";
            this.btnInsertTotalPages.ShowImage = true;
            this.btnInsertTotalPages.SuperTip = "在当前位置插入文档总页码字段。";
            this.btnInsertTotalPages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertTotalPages_Click);
            // 
            // button8
            // 
            this.button8.Label = "插入编号";
            this.button8.Name = "button8";
            this.button8.OfficeImageId = "Numbering";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            //
            // btnApplyHeitiXiaosi
            //
            this.btnApplyHeitiXiaosi.Label = "黑体小四";
            this.btnApplyHeitiXiaosi.Name = "btnApplyHeitiXiaosi";
            this.btnApplyHeitiXiaosi.OfficeImageId = "FontDialog";
            this.btnApplyHeitiXiaosi.ShowImage = true;
            this.btnApplyHeitiXiaosi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApplyHeitiXiaosi_Click);
            //
            // btnApplySongtiXiaosi
            //
            this.btnApplySongtiXiaosi.Label = "宋体小四";
            this.btnApplySongtiXiaosi.Name = "btnApplySongtiXiaosi";
            this.btnApplySongtiXiaosi.OfficeImageId = "FontDialog";
            this.btnApplySongtiXiaosi.ShowImage = true;
            this.btnApplySongtiXiaosi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApplySongtiXiaosi_Click);
            // 
            // btnTogglePageWhitespace
            // 
            this.btnTogglePageWhitespace.Label = "页面间空白";
            this.btnTogglePageWhitespace.Name = "btnTogglePageWhitespace";
            this.btnTogglePageWhitespace.OfficeImageId = "PageScaleToFitHeight";
            this.btnTogglePageWhitespace.ScreenTip = "显示或隐藏页面间空白";
            this.btnTogglePageWhitespace.ShowImage = true;
            this.btnTogglePageWhitespace.SuperTip = "在页面视图中切换页面之间的空白显示。";
            this.btnTogglePageWhitespace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTogglePageWhitespace_Click);
            // 
            // button32
            // 
            this.button32.Label = "图片单倍行距";
            this.button32.Name = "button32";
            this.button32.OfficeImageId = "LineSpacing";
            this.button32.ShowImage = true;
            this.button32.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button32_Click);
            // splitButtonClean
            // 
            this.splitButtonClean.Items.Add(this.btnClearFormatting);
            this.splitButtonClean.Items.Add(this.btnClearManualHeadingNumbers);
            this.splitButtonClean.Items.Add(this.btnCleanBlankPages);
            this.splitButtonClean.Label = "清理";
            this.splitButtonClean.Name = "splitButtonClean";
            this.splitButtonClean.OfficeImageId = "RelationshipsClearLayout";
            // 
            // btnClearFormatting
            // 
            this.btnClearFormatting.Label = "一键清除格式";
            this.btnClearFormatting.Name = "btnClearFormatting";
            this.btnClearFormatting.OfficeImageId = "ClearFormatting";
            this.btnClearFormatting.ShowImage = true;
            this.btnClearFormatting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearFormatting_Click);
            // 
            // btnClearManualHeadingNumbers
            // 
            this.btnClearManualHeadingNumbers.Label = "清除标题前的手工编号";
            this.btnClearManualHeadingNumbers.Name = "btnClearManualHeadingNumbers";
            this.btnClearManualHeadingNumbers.OfficeImageId = "Numbering";
            this.btnClearManualHeadingNumbers.ShowImage = true;
            this.btnClearManualHeadingNumbers.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClearManualHeadingNumbers_Click);
            // 
            // btnCleanBlankPages
            // 
            this.btnCleanBlankPages.Label = "清理空白页";
            this.btnCleanBlankPages.Name = "btnCleanBlankPages";
            this.btnCleanBlankPages.OfficeImageId = "Delete";
            this.btnCleanBlankPages.ShowImage = true;
            this.btnCleanBlankPages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanBlankPages_Click);
            // 
            // splitButtonUpdate
            // 
            this.splitButtonUpdate.Items.Add(this.button26);
            this.splitButtonUpdate.Items.Add(this.button7);
            this.splitButtonUpdate.Items.Add(this.btnUpdateCaptions);
            this.splitButtonUpdate.Items.Add(this.btnUpdateOutlineList);
            this.splitButtonUpdate.Label = "更新";
            this.splitButtonUpdate.Name = "splitButtonUpdate";
            this.splitButtonUpdate.OfficeImageId = "AccessRefreshAllLists";
            // 
            // button26
            // 
            this.button26.Label = "更新目录";
            this.button26.Name = "button26";
            this.button26.OfficeImageId = "AccessRefreshAllLists";
            this.button26.ScreenTip = "更新文档目录";
            this.button26.ShowImage = true;
            this.button26.SuperTip = "重新更新当前文档中的目录内容和页码。";
            this.button26.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button26_Click);
            // 
            // button7
            // 
            this.button7.Label = "更新总页码";
            this.button7.Name = "button7";
            this.button7.OfficeImageId = "CustomPageNumberTopGallery";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // btnUpdateCaptions
            // 
            this.btnUpdateCaptions.Label = "更新题注";
            this.btnUpdateCaptions.Name = "btnUpdateCaptions";
            this.btnUpdateCaptions.OfficeImageId = "CitationInsert";
            this.btnUpdateCaptions.ShowImage = true;
            this.btnUpdateCaptions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateCaptions_Click);
            // 
            // btnUpdateOutlineList
            // 
            this.btnUpdateOutlineList.Label = "更新所选章节号";
            this.btnUpdateOutlineList.Name = "btnUpdateOutlineList";
            this.btnUpdateOutlineList.OfficeImageId = "Numbering";
            this.btnUpdateOutlineList.ShowImage = true;
            this.btnUpdateOutlineList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRebuildOutlineList_Click);
            // 
            // button18
            // 
            this.button18.Label = "拆分表格";
            this.button18.Name = "button18";
            this.button18.OfficeImageId = "TableSplitTable";
            this.button18.ScreenTip = "拆分表格";
            this.button18.ShowImage = true;
            this.button18.SuperTip = "按 Word 原生拆分表格方式拆分当前表格，并在续表上方插入续表题注。";
            this.button18.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button18_Click);
            // 
            // button25
            // 
            this.button25.Label = "快速表格样式";
            this.button25.Name = "button25";
            this.button25.OfficeImageId = "TableAutoFormat";
            this.button25.ScreenTip = "快速表格样式";
            this.button25.ShowImage = true;
            this.button25.SuperTip = "只对当前选中的表格应用统一规范样式。";
            this.button25.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button25_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button14);
            this.group3.Items.Add(this.button13);
            this.group3.Items.Add(this.splitButtonReferenceCaption);
            this.group3.Label = "题注";
            this.group3.Name = "group3";
            // 
            // button14
            // 
            this.button14.Label = "插入图片题注";
            this.button14.Name = "button14";
            this.button14.OfficeImageId = "PictureInsertFromFile";
            this.button14.ShowImage = true;
            this.button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button14_Click);
            // 
            // button13
            // 
            this.button13.Label = "插入表格题注";
            this.button13.Name = "button13";
            this.button13.OfficeImageId = "TableInsert";
            this.button13.ShowImage = true;
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // splitButtonReferenceCaption
            // 
            this.splitButtonReferenceCaption.Items.Add(this.button29);
            this.splitButtonReferenceCaption.Items.Add(this.button28);
            this.splitButtonReferenceCaption.Items.Add(this.button31);
            this.splitButtonReferenceCaption.Label = "引用题注";
            this.splitButtonReferenceCaption.Name = "splitButtonReferenceCaption";
            this.splitButtonReferenceCaption.OfficeImageId = "CitationInsert";
            this.splitButtonReferenceCaption.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button29_Click);
            // 
            // button29
            // 
            this.button29.Label = "引用下一个题注";
            this.button29.Name = "button29";
            this.button29.OfficeImageId = "CitationInsert";
            this.button29.ShowImage = true;
            this.button29.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button29_Click);
            // 
            // button28
            // 
            this.button28.Label = "引用上一个题注";
            this.button28.Name = "button28";
            this.button28.OfficeImageId = "CitationInsert";
            this.button28.ShowImage = true;
            this.button28.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button28_Click);
            // 
            // button31
            // 
            this.button31.Label = "引用自定义题注";
            this.button31.Name = "button31";
            this.button31.OfficeImageId = "CitationInsert";
            this.button31.ShowImage = true;
            this.button31.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button31_Click);
            // 
            // group7
            // 
            this.group7.Items.Add(this.splitButton4);
            this.group7.Label = "显示";
            this.group7.Name = "group7";
            // 
            // splitButton4
            // 
            this.splitButton4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton4.Items.Add(this.button9);
            this.splitButton4.Items.Add(this.button10);
            this.splitButton4.Items.Add(this.button11);
            this.splitButton4.Label = "窗格显示";
            this.splitButton4.Name = "splitButton4";
            this.splitButton4.OfficeImageId = "SelectionPane";
            // 
            // button9
            // 
            this.button9.Label = "书签窗格";
            this.button9.Name = "button9";
            this.button9.OfficeImageId = "BookmarkInsert";
            this.button9.ShowImage = true;
            // 
            // button10
            // 
            this.button10.Label = "题注窗格";
            this.button10.Name = "button10";
            this.button10.OfficeImageId = "CitationInsert";
            this.button10.ShowImage = true;
            // 
            // button11
            // 
            this.button11.Label = "标识窗格";
            this.button11.Name = "button11";
            this.button11.OfficeImageId = "DocumentInspect";
            this.button11.ShowImage = true;
            // 
            // button12
            // 
            this.button12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button12.Label = "需求追踪";
            this.button12.Name = "button12";
            this.button12.OfficeImageId = "MailMergeSelectRecipients";
            this.button12.ShowImage = true;
            // 
            // groupDocumentCheck
            // 
            this.groupDocumentCheck.Items.Add(this.chkNonBodyBlankLine);
            this.groupDocumentCheck.Items.Add(this.chkCaptionContinuity);
            this.groupDocumentCheck.Items.Add(this.chkListContinuity);
            this.groupDocumentCheck.Items.Add(this.chkStyleConsistency);
            this.groupDocumentCheck.Items.Add(this.chkBrokenReferences);
            this.groupDocumentCheck.Items.Add(this.btnStartDocumentCheck);
            this.groupDocumentCheck.Label = "文档检查";
            this.groupDocumentCheck.Name = "groupDocumentCheck";
            // 
            // chkNonBodyBlankLine
            // 
            this.chkNonBodyBlankLine.Label = "章节标题为空";
            this.chkNonBodyBlankLine.Name = "chkNonBodyBlankLine";
            // 
            // chkCaptionContinuity
            // 
            this.chkCaptionContinuity.Label = "题注连续性";
            this.chkCaptionContinuity.Name = "chkCaptionContinuity";
            // 
            // chkListContinuity
            // 
            this.chkListContinuity.Label = "多级列表连续性";
            this.chkListContinuity.Name = "chkListContinuity";
            // 
            // chkStyleConsistency
            // 
            this.chkStyleConsistency.Label = "样式一致性";
            this.chkStyleConsistency.Name = "chkStyleConsistency";
            // 
            // chkBrokenReferences
            // 
            this.chkBrokenReferences.Label = "未更新域";
            this.chkBrokenReferences.Name = "chkBrokenReferences";
            // 
            // btnStartDocumentCheck
            // 
            this.btnStartDocumentCheck.Label = "开始检查";
            this.btnStartDocumentCheck.Name = "btnStartDocumentCheck";
            this.btnStartDocumentCheck.OfficeImageId = "SpellingAndGrammar";
            this.btnStartDocumentCheck.ShowImage = true;
            this.btnStartDocumentCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStartDocumentCheck_Click);
            // 
            // groupSoftwareTools
            // 
            this.groupSoftwareTools.Items.Add(this.btnRequirementExtraction);
            this.groupSoftwareTools.Items.Add(this.button12);
            this.groupSoftwareTools.Items.Add(this.btnSoftwareDocumentCheck);
            this.groupSoftwareTools.Label = "软件专用";
            this.groupSoftwareTools.Name = "groupSoftwareTools";
            // 
            // btnRequirementExtraction
            // 
            this.btnRequirementExtraction.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRequirementExtraction.Label = "需求提取";
            this.btnRequirementExtraction.Name = "btnRequirementExtraction";
            this.btnRequirementExtraction.OfficeImageId = "TableInsert";
            this.btnRequirementExtraction.ShowImage = true;
            this.btnRequirementExtraction.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRequirementExtraction_Click);
            // 
            // btnSoftwareDocumentCheck
            // 
            this.btnSoftwareDocumentCheck.Label = "软件文档检查";
            this.btnSoftwareDocumentCheck.Name = "btnSoftwareDocumentCheck";
            this.btnSoftwareDocumentCheck.OfficeImageId = "DocumentInspect";
            this.btnSoftwareDocumentCheck.ShowImage = true;
            this.btnSoftwareDocumentCheck.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSoftwareDocumentCheck_Click);
            // 
            // group8
            // 
            this.group8.Items.Add(this.menuHelp);
            this.group8.Label = "关于";
            this.group8.Name = "group8";
            // 
            // menuHelp
            // 
            this.menuHelp.Items.Add(this.btnHelpVersion);
            this.menuHelp.Items.Add(this.btnOpenHelpDocument);
            this.menuHelp.Label = "关于";
            this.menuHelp.Name = "menuHelp";
            this.menuHelp.OfficeImageId = "Help";
            // 
            // btnHelpVersion
            // 
            this.btnHelpVersion.Enabled = false;
            this.btnHelpVersion.Label = "版本号：0.0.1.2";
            this.btnHelpVersion.Name = "btnHelpVersion";
            this.btnHelpVersion.OfficeImageId = "Info";
            this.btnHelpVersion.ShowImage = true;
            // 
            // btnOpenHelpDocument
            // 
            this.btnOpenHelpDocument.Label = "打开帮助文档";
            this.btnOpenHelpDocument.Name = "btnOpenHelpDocument";
            this.btnOpenHelpDocument.OfficeImageId = "Help";
            this.btnOpenHelpDocument.ShowImage = true;
            this.btnOpenHelpDocument.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenHelpDocument_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupDocumentManage.ResumeLayout(false);
            this.groupDocumentManage.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.groupDocumentCheck.ResumeLayout(false);
            this.groupDocumentCheck.PerformLayout();
            this.groupSoftwareTools.ResumeLayout(false);
            this.groupSoftwareTools.PerformLayout();
            this.group8.ResumeLayout(false);
            this.group8.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown styleGalleryDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown outlineLevelDropDown;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStyleBinding;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateCustomStyles;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDocumentManage;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBatchReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStyleBrush;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu btnSwitchWindows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenCurrentFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveAllDocuments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCloseOtherDocuments;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertTotalPages;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonClean;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button18;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTogglePageWhitespace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button32;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearFormatting;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClearManualHeadingNumbers;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanBlankPages;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateCaptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateOutlineList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button25;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApplyHeitiXiaosi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApplySongtiXiaosi;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButtonReferenceCaption;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button28;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button29;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button31;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRequirementExtraction;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button26;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDocumentCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkNonBodyBlankLine;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkCaptionContinuity;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkListContinuity;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkStyleConsistency;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkBrokenReferences;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStartDocumentCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSoftwareTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSoftwareDocumentCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group8;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menuHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelpVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenHelpDocument;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
