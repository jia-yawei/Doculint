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
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl2 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.button27 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btnBatchReplace = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button30 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnStyle1 = this.Factory.CreateRibbonToggleButton();
            this.btnStyle2 = this.Factory.CreateRibbonToggleButton();
            this.btnStyle3 = this.Factory.CreateRibbonToggleButton();
            this.btnStyle4 = this.Factory.CreateRibbonToggleButton();
            this.btnStyle5 = this.Factory.CreateRibbonToggleButton();
            this.btnStyle6 = this.Factory.CreateRibbonToggleButton();
            this.btnStyleBody = this.Factory.CreateRibbonToggleButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.splitButton2 = this.Factory.CreateRibbonSplitButton();
            this.button22 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button19 = this.Factory.CreateRibbonButton();
            this.splitButton7 = this.Factory.CreateRibbonSplitButton();
            this.button21 = this.Factory.CreateRibbonButton();
            this.button25 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button18 = this.Factory.CreateRibbonButton();
            this.button24 = this.Factory.CreateRibbonButton();
            this.button20 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.button17 = this.Factory.CreateRibbonButton();
            this.button28 = this.Factory.CreateRibbonButton();
            this.button29 = this.Factory.CreateRibbonButton();
            this.button16 = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.splitButton6 = this.Factory.CreateRibbonSplitButton();
            this.btnInsertTotalPages = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.btnTogglePageWhitespace = this.Factory.CreateRibbonToggleButton();
            this.btnRebuildOutlineList = this.Factory.CreateRibbonButton();
            this.button26 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.group7 = this.Factory.CreateRibbonGroup();
            this.splitButton4 = this.Factory.CreateRibbonSplitButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.button12 = this.Factory.CreateRibbonButton();
            this.group8 = this.Factory.CreateRibbonGroup();
            this.tab1.SuspendLayout();
            this.group4.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group6.SuspendLayout();
            this.group7.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Groups.Add(this.group7);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group8);
            this.tab1.Label = "搞快点";
            this.tab1.Name = "tab1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.splitButton1);
            this.group4.Items.Add(this.btnBatchReplace);
            this.group4.Items.Add(this.button2);
            this.group4.Items.Add(this.button3);
            this.group4.Items.Add(this.button30);
            this.group4.Label = "批量功能";
            this.group4.Name = "group4";
            // 
            // splitButton1
            // 
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Image = ((System.Drawing.Image)(resources.GetObject("splitButton1.Image")));
            this.splitButton1.Items.Add(this.button27);
            this.splitButton1.Items.Add(this.button4);
            this.splitButton1.Items.Add(this.button1);
            this.splitButton1.Label = "文档组";
            this.splitButton1.Name = "splitButton1";
            // 
            // button27
            // 
            this.button27.Enabled = false;
            this.button27.Label = "当前活动组：未设置";
            this.button27.Name = "button27";
            this.button27.ShowImage = true;
            // 
            // button4
            // 
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "当前文档添加到组";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "文档组管理";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // btnBatchReplace
            // 
            this.btnBatchReplace.Image = ((System.Drawing.Image)(resources.GetObject("btnBatchReplace.Image")));
            this.btnBatchReplace.Label = "批量替换";
            this.btnBatchReplace.Name = "btnBatchReplace";
            this.btnBatchReplace.ShowImage = true;
            this.btnBatchReplace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnBatchReplace_Click);
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "抓取管理";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "内容抓取";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button30
            // 
            this.button30.Image = ((System.Drawing.Image)(resources.GetObject("button30.Image")));
            this.button30.Label = "内容注入";
            this.button30.Name = "button30";
            this.button30.ShowImage = true;
            this.button30.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button30_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnStyle1);
            this.group1.Items.Add(this.btnStyle2);
            this.group1.Items.Add(this.btnStyle3);
            this.group1.Items.Add(this.btnStyle4);
            this.group1.Items.Add(this.btnStyle5);
            this.group1.Items.Add(this.btnStyle6);
            this.group1.Items.Add(this.btnStyleBody);
            this.group1.Items.Add(this.btnRebuildOutlineList);
            this.group1.DialogLauncher = ribbonDialogLauncherImpl2;
            this.group1.Label = "常用样式库";
            this.group1.Name = "group1";
            this.group1.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group1_DialogLauncherClick);
            // 
            // btnStyle1
            // 
            this.btnStyle1.Image = ((System.Drawing.Image)(resources.GetObject("btnStyle1.Image")));
            this.btnStyle1.Label = "一级标题";
            this.btnStyle1.Name = "btnStyle1";
            this.btnStyle1.ShowImage = true;
            this.btnStyle1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyle1_Click_1);
            // 
            // btnStyle2
            // 
            this.btnStyle2.Image = ((System.Drawing.Image)(resources.GetObject("btnStyle2.Image")));
            this.btnStyle2.Label = "二级标题";
            this.btnStyle2.Name = "btnStyle2";
            this.btnStyle2.ShowImage = true;
            this.btnStyle2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyle2_Click_1);
            // 
            // btnStyle3
            // 
            this.btnStyle3.Image = ((System.Drawing.Image)(resources.GetObject("btnStyle3.Image")));
            this.btnStyle3.Label = "三级标题";
            this.btnStyle3.Name = "btnStyle3";
            this.btnStyle3.ShowImage = true;
            this.btnStyle3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyle3_Click_1);
            // 
            // btnStyle4
            // 
            this.btnStyle4.Image = ((System.Drawing.Image)(resources.GetObject("btnStyle4.Image")));
            this.btnStyle4.Label = "四级标题";
            this.btnStyle4.Name = "btnStyle4";
            this.btnStyle4.ShowImage = true;
            this.btnStyle4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyle4_Click_1);
            // 
            // btnStyle5
            // 
            this.btnStyle5.Image = ((System.Drawing.Image)(resources.GetObject("btnStyle5.Image")));
            this.btnStyle5.Label = "五级标题";
            this.btnStyle5.Name = "btnStyle5";
            this.btnStyle5.ShowImage = true;
            this.btnStyle5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyle5_Click_1);
            // 
            // btnStyle6
            // 
            this.btnStyle6.Label = "六级标题";
            this.btnStyle6.Name = "btnStyle6";
            this.btnStyle6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyle6_Click);
            // 
            // btnStyleBody
            // 
            this.btnStyleBody.Label = "正文";
            this.btnStyleBody.Name = "btnStyleBody";
            this.btnStyleBody.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStyleBody_Click);
            // 
            // group2
            // 
            this.group2.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group2.Items.Add(this.splitButton2);
            this.group2.Items.Add(this.splitButton7);
            this.group2.Items.Add(this.button18);
            this.group2.Items.Add(this.button24);
            this.group2.Items.Add(this.button20);
            this.group2.Label = "图表";
            this.group2.Name = "group2";
            this.group2.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.group2_DialogLauncherClick);
            // 
            // splitButton2
            // 
            this.splitButton2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton2.Image = ((System.Drawing.Image)(resources.GetObject("splitButton2.Image")));
            this.splitButton2.Items.Add(this.button22);
            this.splitButton2.Items.Add(this.button5);
            this.splitButton2.Items.Add(this.button19);
            this.splitButton2.Label = "表格选取";
            this.splitButton2.Name = "splitButton2";
            this.splitButton2.ScreenTip = "选择表格范围";
            this.splitButton2.SuperTip = "快速定位续表、跨页表格或全部表格。";
            // 
            // button22
            // 
            this.button22.Image = ((System.Drawing.Image)(resources.GetObject("button22.Image")));
            this.button22.Label = "下个续表";
            this.button22.Name = "button22";
            this.button22.ScreenTip = "选择下个续表";
            this.button22.ShowImage = true;
            this.button22.SuperTip = "从当前光标位置开始，选中下一个续表。";
            this.button22.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button22_Click);
            // 
            // button5
            // 
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Label = "下个跨页表格";
            this.button5.Name = "button5";
            this.button5.ScreenTip = "选择下个跨页表格";
            this.button5.ShowImage = true;
            this.button5.SuperTip = "从当前光标位置开始，选中下一个跨页表格。";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button19
            // 
            this.button19.Image = ((System.Drawing.Image)(resources.GetObject("button19.Image")));
            this.button19.Label = "全部表格";
            this.button19.Name = "button19";
            this.button19.ScreenTip = "选择全部表格";
            this.button19.ShowImage = true;
            this.button19.SuperTip = "选中文档中的全部表格。";
            this.button19.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button19_Click);
            // 
            // splitButton7
            // 
            this.splitButton7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton7.Image = ((System.Drawing.Image)(resources.GetObject("splitButton7.Image")));
            this.splitButton7.Items.Add(this.button21);
            this.splitButton7.Items.Add(this.button25);
            this.splitButton7.Items.Add(this.button6);
            this.splitButton7.Label = "图表规范";
            this.splitButton7.Name = "splitButton7";
            this.splitButton7.ScreenTip = "统一规范图表";
            this.splitButton7.SuperTip = "批量规范表格和图片的字体、宽度、边框等样式。";
            // 
            // button21
            // 
            this.button21.Image = ((System.Drawing.Image)(resources.GetObject("button21.Image")));
            this.button21.Label = "规范所有表格";
            this.button21.Name = "button21";
            this.button21.ScreenTip = "规范所有表格";
            this.button21.ShowImage = true;
            this.button21.SuperTip = "对文档中的所有表格应用统一规范。";
            this.button21.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button21_Click);
            // 
            // button25
            // 
            this.button25.Image = ((System.Drawing.Image)(resources.GetObject("button25.Image")));
            this.button25.Label = "规范选中的表格";
            this.button25.Name = "button25";
            this.button25.ScreenTip = "规范选中的表格";
            this.button25.ShowImage = true;
            this.button25.SuperTip = "只对当前选中的表格应用统一规范。";
            this.button25.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button25_Click);
            // 
            // button6
            // 
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Label = "规范全部图片";
            this.button6.Name = "button6";
            this.button6.ScreenTip = "规范全部图片";
            this.button6.ShowImage = true;
            this.button6.SuperTip = "对文档中的所有图片应用统一规范。";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button18
            // 
            this.button18.Image = ((System.Drawing.Image)(resources.GetObject("button18.Image")));
            this.button18.Label = "按续表拆分";
            this.button18.Name = "button18";
            this.button18.ScreenTip = "按续表拆分表格";
            this.button18.ShowImage = true;
            this.button18.SuperTip = "按跨页位置将当前表格拆成续表。";
            this.button18.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button18_Click);
            // 
            // button24
            // 
            this.button24.Image = ((System.Drawing.Image)(resources.GetObject("button24.Image")));
            this.button24.Label = "合并续表";
            this.button24.Name = "button24";
            this.button24.ScreenTip = "合并续表";
            this.button24.ShowImage = true;
            this.button24.SuperTip = "将续表和前面的正表合并为一个表格。";
            this.button24.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button24_Click);
            // 
            // button20
            // 
            this.button20.Image = ((System.Drawing.Image)(resources.GetObject("button20.Image")));
            this.button20.Label = "设置为表头";
            this.button20.Name = "button20";
            this.button20.ScreenTip = "设置为表头";
            this.button20.ShowImage = true;
            this.button20.SuperTip = "将当前选中的行标记为表头。";
            this.button20.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button20_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button14);
            this.group3.Items.Add(this.button13);
            this.group3.Items.Add(this.button17);
            this.group3.Items.Add(this.button28);
            this.group3.Items.Add(this.button29);
            this.group3.Items.Add(this.button16);
            this.group3.Label = "题注";
            this.group3.Name = "group3";
            // 
            // button14
            // 
            this.button14.Image = ((System.Drawing.Image)(resources.GetObject("button14.Image")));
            this.button14.Label = "插入图片题注";
            this.button14.Name = "button14";
            this.button14.ShowImage = true;
            this.button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button14_Click);
            // 
            // button13
            // 
            this.button13.Image = ((System.Drawing.Image)(resources.GetObject("button13.Image")));
            this.button13.Label = "插入表格题注";
            this.button13.Name = "button13";
            this.button13.ShowImage = true;
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // button17
            // 
            this.button17.Image = ((System.Drawing.Image)(resources.GetObject("button17.Image")));
            this.button17.Label = "更新图片题注";
            this.button17.Name = "button17";
            this.button17.ShowImage = true;
            this.button17.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button17_Click);
            // 
            // button28
            // 
            this.button28.Image = ((System.Drawing.Image)(resources.GetObject("button28.Image")));
            this.button28.Label = "引用上一个题注";
            this.button28.Name = "button28";
            this.button28.ShowImage = true;
            this.button28.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button28_Click);
            // 
            // button29
            // 
            this.button29.Image = ((System.Drawing.Image)(resources.GetObject("button29.Image")));
            this.button29.Label = "引用下一个题注";
            this.button29.Name = "button29";
            this.button29.ShowImage = true;
            this.button29.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button29_Click);
            // 
            // button16
            // 
            this.button16.Image = ((System.Drawing.Image)(resources.GetObject("button16.Image")));
            this.button16.Label = "更新表格题注";
            this.button16.Name = "button16";
            this.button16.ShowImage = true;
            this.button16.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button16_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.splitButton6);
            this.group6.Items.Add(this.btnTogglePageWhitespace);
            this.group6.Items.Add(this.button26);
            this.group6.Items.Add(this.button7);
            this.group6.Label = "快速工具";
            this.group6.Name = "group6";
            // 
            // splitButton6
            // 
            this.splitButton6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton6.Image = ((System.Drawing.Image)(resources.GetObject("splitButton6.Image")));
            this.splitButton6.Items.Add(this.btnInsertTotalPages);
            this.splitButton6.Items.Add(this.button8);
            this.splitButton6.Label = "插入项";
            this.splitButton6.Name = "splitButton6";
            // 
            // btnInsertTotalPages
            // 
            this.btnInsertTotalPages.Image = ((System.Drawing.Image)(resources.GetObject("btnInsertTotalPages.Image")));
            this.btnInsertTotalPages.Label = "总页码";
            this.btnInsertTotalPages.Name = "btnInsertTotalPages";
            this.btnInsertTotalPages.ScreenTip = "插入文档总页码";
            this.btnInsertTotalPages.ShowImage = true;
            this.btnInsertTotalPages.SuperTip = "在当前位置插入文档总页码字段。";
            this.btnInsertTotalPages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertTotalPages_Click);
            // 
            // button8
            // 
            this.button8.Image = ((System.Drawing.Image)(resources.GetObject("button8.Image")));
            this.button8.Label = "自动序号";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // btnTogglePageWhitespace
            // 
            this.btnTogglePageWhitespace.Image = ((System.Drawing.Image)(resources.GetObject("btnTogglePageWhitespace.Image")));
            this.btnTogglePageWhitespace.Label = "页面间空白";
            this.btnTogglePageWhitespace.Name = "btnTogglePageWhitespace";
            this.btnTogglePageWhitespace.ScreenTip = "显示或隐藏页面间空白";
            this.btnTogglePageWhitespace.ShowImage = true;
            this.btnTogglePageWhitespace.SuperTip = "在页面视图中切换页面之间的空白显示。";
            this.btnTogglePageWhitespace.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTogglePageWhitespace_Click);
            // 
            // btnRebuildOutlineList
            // 
            this.btnRebuildOutlineList.Image = ((System.Drawing.Image)(resources.GetObject("btnRebuildOutlineList.Image")));
            this.btnRebuildOutlineList.Label = "更新全部章节号";
            this.btnRebuildOutlineList.Name = "btnRebuildOutlineList";
            this.btnRebuildOutlineList.ShowImage = true;
            this.btnRebuildOutlineList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRebuildOutlineList_Click);
            // 
            // button26
            // 
            this.button26.Image = ((System.Drawing.Image)(resources.GetObject("button26.Image")));
            this.button26.Label = "刷新目录";
            this.button26.Name = "button26";
            this.button26.ScreenTip = "刷新文档目录";
            this.button26.ShowImage = true;
            this.button26.SuperTip = "重新更新当前文档中的目录内容和页码。";
            this.button26.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button26_Click);
            // 
            // button7
            // 
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.Label = "更新总页码";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
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
            this.splitButton4.Image = ((System.Drawing.Image)(resources.GetObject("splitButton4.Image")));
            this.splitButton4.Items.Add(this.button9);
            this.splitButton4.Items.Add(this.button10);
            this.splitButton4.Items.Add(this.button11);
            this.splitButton4.Label = "窗格显示";
            this.splitButton4.Name = "splitButton4";
            // 
            // button9
            // 
            this.button9.Image = ((System.Drawing.Image)(resources.GetObject("button9.Image")));
            this.button9.Label = "书签窗格";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            // 
            // button10
            // 
            this.button10.Image = ((System.Drawing.Image)(resources.GetObject("button10.Image")));
            this.button10.Label = "题注窗格";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            // 
            // button11
            // 
            this.button11.Image = ((System.Drawing.Image)(resources.GetObject("button11.Image")));
            this.button11.Label = "标识窗格";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            // 
            // group5
            // 
            this.group5.Items.Add(this.button12);
            this.group5.Label = "追踪";
            this.group5.Name = "group5";
            // 
            // button12
            // 
            this.button12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button12.Image = ((System.Drawing.Image)(resources.GetObject("button12.Image")));
            this.button12.Label = "需求追踪";
            this.button12.Name = "button12";
            this.button12.ShowImage = true;
            this.button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button12_Click_1);
            // 
            // group8
            // 
            this.group8.Label = "帮助";
            this.group8.Name = "group8";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group7.ResumeLayout(false);
            this.group7.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnStyle1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnStyle2;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnStyle3;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnStyle4;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnStyle5;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnStyle6;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnStyleBody;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnBatchReplace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRebuildOutlineList;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertTotalPages;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button18;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button19;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button21;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button20;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button22;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button17;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button16;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton6;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTogglePageWhitespace;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button24;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button25;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button28;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button29;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button26;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button27;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button30;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group8;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
