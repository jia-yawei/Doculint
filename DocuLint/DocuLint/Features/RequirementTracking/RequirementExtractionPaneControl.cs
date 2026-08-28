using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal sealed class RequirementExtractionPaneControl : UserControl
    {
        private const string SavedRequirementsNamespace = "urn:doculint:requirement-extraction";
        private readonly Func<Word.Application> applicationAccessor;
        private readonly DataGridView grid;
        private readonly FlowLayoutPanel normalModeActions;
        private readonly FlowLayoutPanel customModeActions;
        private readonly ComboBox extractionModeSelector;
        private readonly Button btnRefresh;
        private readonly Button btnCustomBatchExtraction;
        private readonly Button btnCustomSave;
        private readonly ComboBox cmbCustomExtractionMode;
        private readonly Button btnExtractionEnabled;
        private readonly Button btnBatchExtraction;
        private readonly Button btnClearSelected;
        private readonly Button btnClearAll;
        private readonly Button btnDelete;
        private readonly Button btnSave;
        private readonly Label lblStatus;
        private readonly ToolTip toolTip;
        private readonly ContextMenuStrip gridContextMenu;
        private readonly ToolStripMenuItem insertCustomRowAboveMenuItem;
        private readonly ToolStripMenuItem insertCustomRowBelowMenuItem;
        private Word.Document currentDocument;
        private readonly List<RequirementItem> requirements = new List<RequirementItem>();
        private readonly HashSet<string> excludedRequirementIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool extractionEnabled;
        private bool batchExtractionEnabled;
        private bool hasUnsavedChanges;
        private bool suppressGridValueChanged;
        private bool customMode;
        private bool customBatchExtractionEnabled;
        private bool customAlternatingExpectsIdentifier = true;
        private bool updatingExtractionViewMode;
        private List<string> extractionMarkerTemplates = new List<string>();
        private bool extractionUseCustomTemplates;
        private DocumentMarkerDocumentType extractionPresetTemplateType = DocumentMarkerDocumentType.Unknown;
        private bool extractionMarkerPageRangeEnabled;
        private int extractionMarkerStartPage = 1;
        private int extractionMarkerEndPage = 1;
        private int extractionSectionNumberLevel;
        private bool extractionScanFieldResults;

        private enum CustomExtractionMode
        {
            IdentifierOnly,
            NameOnly,
            Alternating
        }

        internal event Action<int> RequirementActivated;
        internal event Action<bool> BatchExtractionModeChanged;
        internal event Action<bool> ExtractionModeChanged;

        internal RequirementExtractionPaneControl(Func<Word.Application> applicationAccessor)
        {
            this.applicationAccessor = applicationAccessor;
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            toolTip = new ToolTip
            {
                InitialDelay = 350,
                ReshowDelay = 100,
                AutoPopDelay = 5000,
                ShowAlways = true
            };
            DoubleBuffered = true;
            SetStyle(
                ControlStyles.AllPaintingInWmPaint |
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw,
                true);
            UpdateStyles();

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };
            EnableDoubleBuffering(layout);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            TableLayoutPanel toolbar = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 3,
                Margin = new Padding(0, 0, 0, 8)
            };
            toolbar.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            toolbar.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            toolbar.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            EnableDoubleBuffering(toolbar);

            FlowLayoutPanel viewModeSelector = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(0, 0, 0, 4),
                Margin = Padding.Empty
            };
            extractionModeSelector = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei UI", 9F),
                Width = 156,
                Height = 32,
                Margin = new Padding(0, 0, 8, 0)
            };
            extractionModeSelector.Items.AddRange(new object[] { "常规提取", "自定义提取" });
            extractionModeSelector.SelectedIndex = 0;
            extractionModeSelector.SelectedIndexChanged += (_, __) =>
            {
                if (!updatingExtractionViewMode)
                {
                    SetCustomMode(extractionModeSelector.SelectedIndex == 1, true);
                }
            };
            viewModeSelector.Controls.Add(extractionModeSelector);

            normalModeActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(0, 38),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(0, 2, 0, 4),
                Margin = Padding.Empty
            };
            customModeActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                MinimumSize = new Size(0, 38),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(0, 2, 0, 4),
                Margin = Padding.Empty,
                Visible = false
            };
            toolbar.Controls.Add(viewModeSelector, 0, 0);
            toolbar.Controls.Add(normalModeActions, 0, 1);
            toolbar.Controls.Add(customModeActions, 0, 2);

            btnRefresh = CreateButton("加载标识");
            btnRefresh.Click += (_, __) => ReloadCurrentDocumentRequirements();
            cmbCustomExtractionMode = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Microsoft YaHei UI", 9F),
                Width = 128,
                Height = 32,
                Margin = new Padding(0, 1, 8, 1)
            };
            cmbCustomExtractionMode.Items.AddRange(new object[] { "仅标识", "仅名称", "交替" });
            cmbCustomExtractionMode.SelectedIndex = (int)CustomExtractionMode.Alternating;
            cmbCustomExtractionMode.SelectedIndexChanged += (_, __) =>
            {
                customAlternatingExpectsIdentifier = true;
                if (customMode)
                {
                    SetStatus("已切换提取方式；交替模式将从标识开始");
                }
            };
            btnCustomBatchExtraction = CreateButton("开始批量提取");
            btnCustomBatchExtraction.Click += (_, __) => SetCustomBatchExtractionEnabled(!customBatchExtractionEnabled);
            btnCustomSave = CreateButton("保存");
            btnCustomSave.Click += (_, __) => SaveToCurrentDocument();
            btnExtractionEnabled = CreateButton("开始批量提取");
            btnBatchExtraction = CreateButton("开始批量提取");
            btnBatchExtraction.Click += (_, __) => SetRegularBatchExtraction(!BatchExtractionEnabled);
            btnClearSelected = CreateButton("清除");
            btnClearSelected.BackColor = Color.FromArgb(177, 63, 63);
            btnClearSelected.FlatAppearance.MouseOverBackColor = Color.FromArgb(196, 76, 76);
            btnClearSelected.FlatAppearance.MouseDownBackColor = Color.FromArgb(151, 49, 49);
            btnClearSelected.Click += (_, __) => ClearSelectedRequirementContent();
            btnClearAll = CreateButton("全部清空");
            btnClearAll.BackColor = Color.FromArgb(177, 63, 63);
            btnClearAll.FlatAppearance.MouseOverBackColor = Color.FromArgb(196, 76, 76);
            btnClearAll.FlatAppearance.MouseDownBackColor = Color.FromArgb(151, 49, 49);
            btnClearAll.Click += (_, __) => ClearAllRequirements();
            btnDelete = CreateButton("删除");
            btnDelete.BackColor = Color.FromArgb(177, 63, 63);
            btnDelete.FlatAppearance.MouseOverBackColor = Color.FromArgb(196, 76, 76);
            btnDelete.FlatAppearance.MouseDownBackColor = Color.FromArgb(151, 49, 49);
            btnDelete.Click += (_, __) => DeleteSavedResults();
            btnSave = CreateButton("保存");
            btnSave.Click += (_, __) => SaveToCurrentDocument();
            ApplySecondaryButtonStyle(btnExtractionEnabled);
            ApplySecondaryButtonStyle(btnBatchExtraction);
            // 选项卡行只保留模式切换，加载和编辑操作统一放到下一行，避免窄窗格裁切。
            PlaceToolbarButton(btnRefresh, normalModeActions, 122);
            PlaceToolbarButton(btnBatchExtraction, normalModeActions, 146);
            PlaceToolbarButton(btnClearSelected, normalModeActions, 72);
            PlaceToolbarButton(btnClearAll, normalModeActions, 124);
            PlaceToolbarButton(btnDelete, normalModeActions, 64);
            PlaceToolbarButton(btnSave, normalModeActions, 64);
            customModeActions.Controls.Add(cmbCustomExtractionMode);
            PlaceToolbarButton(btnCustomBatchExtraction, customModeActions, 160);
            PlaceToolbarButton(btnCustomSave, customModeActions, 78);
            toolTip.SetToolTip(btnRefresh, "按当前文档类型加载或更新标识，并保留已提取的内容");
            toolTip.SetToolTip(btnBatchExtraction, "点击后开启批量提取；选择文档文字会自动写入当前需求行");
            toolTip.SetToolTip(btnClearSelected, "清除当前选中需求的名称和章节号，保留需求标识");
            toolTip.SetToolTip(btnClearAll, "清空全部需求名称和章节号，保留需求标识");
            toolTip.SetToolTip(btnDelete, "删除全部需求提取结果；删除后需重新点击加载标识");
            toolTip.SetToolTip(btnSave, "将当前需求表保存到 Word 文档");
            toolTip.SetToolTip(extractionModeSelector, "切换常规提取和自定义提取界面");
            toolTip.SetToolTip(cmbCustomExtractionMode, "选择选区写入需求标识、需求名称，或按标识与名称交替写入");
            toolTip.SetToolTip(btnCustomBatchExtraction, "开启后，选择文档文字会按当前方式自动写入需求表");

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeColumns = true,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                ColumnHeadersHeight = 34,
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(230, 235, 242),
                ReadOnly = false,
                RowHeadersVisible = false,
                RowTemplate = { Height = 30 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(238, 242, 247);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(32, 45, 64);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 238, 255);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            EnableDoubleBuffering(grid);
            grid.Columns.Add(CreateTextColumn("Id", "需求标识", true, 160));
            grid.Columns.Add(CreateTextColumn("Name", "需求名称", false, 120));
            grid.Columns.Add(CreateTextColumn("SectionNumber", "章节号", false, 80));
            grid.CellDoubleClick += (_, e) =>
            {
                if (e.RowIndex >= 0 && e.RowIndex < requirements.Count)
                {
                    if (customMode || e.ColumnIndex == 1 || e.ColumnIndex == 2)
                    {
                        grid.CurrentCell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        grid.BeginEdit(true);
                    }
                    else
                    {
                        NavigateToRequirement(requirements[e.RowIndex]);
                    }
                }
            };
            grid.CellClick += (_, e) =>
            {
                if (!customMode &&
                    e.RowIndex >= 0 &&
                    e.RowIndex < requirements.Count &&
                    e.ColumnIndex == 0)
                {
                    NavigateToRequirement(requirements[e.RowIndex]);
                }
            };
            grid.CellValueChanged += Grid_CellValueChanged;
            grid.CellMouseDown += Grid_CellMouseDown;
            grid.KeyDown += Grid_KeyDown;

            gridContextMenu = new ContextMenuStrip();
            ToolStripMenuItem deleteRequirementMenuItem = new ToolStripMenuItem("删除选中标识");
            deleteRequirementMenuItem.Click += (_, __) => DeleteSelectedRequirement();
            insertCustomRowAboveMenuItem = new ToolStripMenuItem("在上方插入行");
            insertCustomRowAboveMenuItem.Click += (_, __) => InsertCustomRow(false);
            insertCustomRowBelowMenuItem = new ToolStripMenuItem("在下方插入行");
            insertCustomRowBelowMenuItem.Click += (_, __) => InsertCustomRow(true);
            gridContextMenu.Items.Add(deleteRequirementMenuItem);
            gridContextMenu.Items.Add(new ToolStripSeparator());
            gridContextMenu.Items.Add(insertCustomRowAboveMenuItem);
            gridContextMenu.Items.Add(insertCustomRowBelowMenuItem);
            grid.ContextMenuStrip = gridContextMenu;

            lblStatus = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Fill,
                Text = "未提取",
                Font = new Font("Microsoft YaHei UI", 9F),
                ForeColor = Color.FromArgb(96, 96, 96)
            };

            layout.Controls.Add(toolbar, 0, 0);
            layout.Controls.Add(grid, 0, 1);
            layout.Controls.Add(lblStatus, 0, 2);
            Controls.Add(layout);
        }

        internal void LoadSavedRequirementsFromCurrentDocument()
        {
            Word.Document doc = applicationAccessor?.Invoke()?.ActiveDocument;
            if (doc == null)
            {
                SetStatus("当前没有活动文档");
                return;
            }

            currentDocument = doc;
            // 每次打开需求提取窗格默认显示常规提取；已保存的自定义结果仍可通过模式下拉框查看。
            SetCustomMode(false, false);
            requirements.Clear();
            excludedRequirementIds.Clear();
            requirements.AddRange(LoadSavedRequirementItems(doc));
            RestoreRequirementPositions(doc);
            excludedRequirementIds.UnionWith(LoadExcludedRequirementIds(doc));
            RenderRows();
            hasUnsavedChanges = false;
            UpdateExtractionActionLabel(doc);
            ApplyBatchExtractionState(batchExtractionEnabled);
            if (requirements.Count > 0)
            {
                SetStatus($"已显示文档中保存的 {requirements.Count} 个需求标识；点击“更新提取”可同步文档新增标识");
                SelectRow(0);
            }
            else
            {
                SetStatus("当前文档没有已保存的需求提取结果；点击“加载标识”扫描标识");
            }
        }

        internal void ReloadCurrentDocumentRequirements()
        {
            if (customMode)
            {
                SetStatus("自定义提取不自动加载标识；请开启批量提取并在文档中选择文字");
                return;
            }

            Word.Application app = applicationAccessor?.Invoke();
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                SetStatus("当前没有活动文档");
                return;
            }

            currentDocument = doc;
            bool hasExistingResults = requirements.Any(item => item != null && !string.IsNullOrWhiteSpace(item.Id));
            Office.CustomXMLPart savedPart = FindSavedRequirementPart(doc);
            if (!hasExistingResults && savedPart != null && ContainsSavedRequirement(savedPart))
            {
                hasExistingResults = true;
            }

            if (hasExistingResults)
            {
                DialogResult result = MessageBox.Show(
                    "更新提取后再次保存会覆盖当前提取结果，原有需求名称和章节号会尽量按标识变化迁移，是否继续？",
                    "更新提取",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (result != DialogResult.Yes)
                {
                    SetStatus("已取消更新提取，当前提取结果未改变");
                    return;
                }
            }

            SetStatus(hasExistingResults ? "正在更新提取结果..." : "正在加载需求标识...");
            try
            {
                DocumentMarkerCollectionResult markerResult = DocumentMarkerService.CollectMarkers(
                    doc,
                    extractionUseCustomTemplates ? GetConfiguredMarkerTemplates() : Enumerable.Empty<string>(),
                    extractionMarkerPageRangeEnabled ? extractionMarkerStartPage : 0,
                    extractionMarkerPageRangeEnabled ? extractionMarkerEndPage : 0,
                    GetEffectivePresetTemplateType(doc),
                    extractionScanFieldResults);
                List<RequirementItem> saved = requirements
                    .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                    .Select(item => new RequirementItem
                    {
                        Id = item.Id,
                        Name = item.Name,
                        SectionNumber = item.SectionNumber
                    })
                    .ToList();
                if (saved.Count == 0)
                {
                    saved = LoadSavedRequirementItems(doc);
                }
                requirements.Clear();
                excludedRequirementIds.Clear();
                requirements.AddRange((markerResult.Entries ?? new List<NavigationPaneEntry>())
                    .Where(entry => entry != null &&
                                    !string.IsNullOrWhiteSpace(entry.Text))
                    .Select(entry => new RequirementItem
                    {
                        Id = entry.Text,
                        Name = string.Empty,
                        SectionNumber = string.Empty,
                        Start = entry.Start,
                        BookmarkOrRange = entry.Start
                    }));
                MergeSavedRequirements(saved);
                RenderRows();
                hasUnsavedChanges = true;
                UpdateExtractionActionLabel(doc);
                ApplyBatchExtractionState(batchExtractionEnabled);
                SetStatus($"已更新 {requirements.Count} 个需求标识；右键或按 Delete 可删除不需要的标识");
                SelectRowAndNavigate(FindNextUnfilledRow(0));
            }
            catch (Exception ex)
            {
                SetStatus("需求提取失败");
                MessageBox.Show($"需求提取失败: {ex.Message}", "需求提取");
            }
        }

        internal void AddSelectionAsRequirement()
        {
            if (!extractionEnabled)
            {
                MessageBox.Show("请先点击“开始批量提取”。", "需求提取");
                return;
            }

            TryAddSelectionAsRequirement(false);
        }

        internal bool TryAutoAddSelection()
        {
            if (customMode)
            {
                return customBatchExtractionEnabled && TryAddSelectionAsCustomRequirement();
            }

            return extractionEnabled && TryAddSelectionAsRequirement(true);
        }

        internal bool ExtractionEnabled => customMode ? customBatchExtractionEnabled : extractionEnabled;

        internal bool BatchExtractionEnabled => customMode
            ? customBatchExtractionEnabled
            : extractionEnabled && batchExtractionEnabled;

        private void SetCustomMode(bool enabled, bool userInitiated)
        {
            if (customMode == enabled)
            {
                ApplyCustomModeState();
                return;
            }

            if (customMode)
            {
                SetCustomBatchExtractionEnabled(false);
            }
            else if (extractionEnabled)
            {
                SetExtractionEnabled(false);
            }

            customMode = enabled;
            customAlternatingExpectsIdentifier = true;
            ApplyCustomModeState();
            if (!userInitiated)
            {
                return;
            }

            SetStatus(customMode
                ? "自定义提取：选择“交替”时将先提取标识，再提取名称"
                : "已退出自定义模式；可点击“更新提取”同步文档标识");
        }

        private void ApplyCustomModeState()
        {
            normalModeActions.Visible = !customMode;
            customModeActions.Visible = customMode;
            btnRefresh.Visible = !customMode;
            btnSave.Visible = !customMode;

            updatingExtractionViewMode = true;
            try
            {
                extractionModeSelector.SelectedIndex = customMode ? 1 : 0;
            }
            finally
            {
                updatingExtractionViewMode = false;
            }

            if (grid != null)
            {
                grid.ReadOnly = false;
                if (grid.Columns.Contains("Id"))
                {
                    grid.Columns["Id"].ReadOnly = !customMode;
                }

                if (grid.Columns.Contains("Name"))
                {
                    grid.Columns["Name"].ReadOnly = false;
                }

                if (grid.Columns.Contains("SectionNumber"))
                {
                    grid.Columns["SectionNumber"].ReadOnly = false;
                }
            }

            insertCustomRowAboveMenuItem.Enabled = customMode;
            insertCustomRowBelowMenuItem.Enabled = customMode;

        }

        private void SetCustomBatchExtractionEnabled(bool enabled)
        {
            if (customBatchExtractionEnabled == enabled)
            {
                return;
            }

            customBatchExtractionEnabled = enabled;
            btnCustomBatchExtraction.Text = enabled ? "结束批量提取" : "开始批量提取";
            btnCustomBatchExtraction.BackColor = enabled ? Color.FromArgb(35, 138, 79) : Color.FromArgb(43, 108, 176);
            BatchExtractionModeChanged?.Invoke(enabled);
            ExtractionModeChanged?.Invoke(enabled);
            SetStatus(enabled
                ? "自定义批量提取已开启：在文档中选择文字即可写入需求表"
                : "自定义批量提取已关闭");
        }

        private void InsertCustomRow(bool belowCurrentRow)
        {
            if (!customMode)
            {
                return;
            }

            int currentRow = GetCurrentRowIndex();
            if (currentRow < 0 || currentRow >= requirements.Count)
            {
                SetStatus("请先选中一行，再选择插入位置");
                return;
            }

            int insertIndex = belowCurrentRow ? currentRow + 1 : currentRow;
            requirements.Insert(insertIndex, new RequirementItem());
            hasUnsavedChanges = true;
            RenderRows();
            SelectRow(insertIndex);
            grid.CurrentCell = grid.Rows[insertIndex].Cells[0];
            grid.BeginEdit(true);
            SetStatus(belowCurrentRow ? "已在当前行下方插入行" : "已在当前行上方插入行");
        }

        private void SetExtractionEnabled(bool enabled)
        {
            if (extractionEnabled == enabled)
            {
                return;
            }

            extractionEnabled = enabled;
            if (!enabled)
            {
                SetBatchExtractionMode(false);
            }

            ExtractionModeChanged?.Invoke(enabled);
            ApplyExtractionButtonState();

            SetStatus(enabled
                ? "常规提取已开启：可在文档中选择文字后填入当前行"
                : "常规提取已关闭");
        }

        private void SetRegularBatchExtraction(bool enabled)
        {
            if (enabled == (extractionEnabled && batchExtractionEnabled))
            {
                return;
            }

            if (enabled)
            {
                extractionEnabled = true;
                batchExtractionEnabled = true;
                ExtractionModeChanged?.Invoke(true);
                BatchExtractionModeChanged?.Invoke(true);
            }
            else
            {
                batchExtractionEnabled = false;
                extractionEnabled = false;
                BatchExtractionModeChanged?.Invoke(false);
                ExtractionModeChanged?.Invoke(false);
            }

            ApplyExtractionButtonState();
            SetStatus(enabled
                ? "批量提取已开启：在文档中选择文字会自动写入当前行"
                : "批量提取已关闭");
        }

        internal void SetBatchExtractionMode(bool enabled)
        {
            if (enabled && !extractionEnabled)
            {
                return;
            }

            if (batchExtractionEnabled == enabled)
            {
                return;
            }

            batchExtractionEnabled = enabled;
            ApplyBatchExtractionState(enabled);
            BatchExtractionModeChanged?.Invoke(enabled);
        }

        private void ApplyBatchExtractionState(bool enabled)
        {
            ApplyExtractionButtonState();

            SetStatus(enabled
                ? "批量提取已开启：在左侧文档中选中文字后会自动写入当前行"
                : "批量提取已关闭");
        }

        private bool TryAddSelectionAsRequirement(bool autoMode)
        {
            Word.Application app = applicationAccessor?.Invoke();
            Word.Selection selection = app?.Selection;
            Word.Range range = selection?.Range;
            if (range == null || range.Start == range.End)
            {
                if (!autoMode)
                {
                    MessageBox.Show("请先在文档中选中要添加为需求名称的文字。", "需求提取");
                }
                return false;
            }

            string name = CleanText(range.Text);
            if (string.IsNullOrWhiteSpace(name))
            {
                if (!autoMode)
                {
                    MessageBox.Show("选中的文字为空，无法添加。", "需求提取");
                }
                return false;
            }

            Word.Document doc = selection.Document ?? currentDocument;
            currentDocument = doc;
            string sectionNumber = RequirementTrackingWordService.ResolveCurrentHeadingNumber(selection);
            if (string.IsNullOrWhiteSpace(sectionNumber))
            {
                sectionNumber = RequirementTrackingWordService.ResolveNearestSectionNumber(doc, range.Start);
            }
            sectionNumber = TruncateSectionNumber(sectionNumber, extractionSectionNumberLevel);
            int rowIndex = GetCurrentRowIndex();
            if (rowIndex < 0)
            {
                if (!autoMode)
                {
                    MessageBox.Show("请先在需求窗格中选中要填写的需求标识行。", "需求提取");
                }
                return false;
            }

            RequirementItem item = requirements[rowIndex];
            item.Name = name;
            item.SectionNumber = sectionNumber;
            item.BookmarkOrRange = range.Start;
            hasUnsavedChanges = true;
            int nextRowIndex = rowIndex + 1;
            RenderRows();
            SelectRow(Math.Min(rowIndex, requirements.Count - 1));
            SetStatus(autoMode ? $"已自动添加到 {item.Id}：{name}" : $"已添加到 {item.Id}：{name}");
            SelectRowAndNavigate(nextRowIndex);
            return true;
        }

        private bool TryAddSelectionAsCustomRequirement()
        {
            Word.Application app = applicationAccessor?.Invoke();
            Word.Selection selection = app?.Selection;
            Word.Range range = selection?.Range;
            if (range == null || range.Start == range.End)
            {
                return false;
            }

            string text = CleanText(range.Text);
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            Word.Document doc = selection.Document ?? currentDocument;
            currentDocument = doc;
            string sectionNumber = RequirementTrackingWordService.ResolveCurrentHeadingNumber(selection);
            if (string.IsNullOrWhiteSpace(sectionNumber))
            {
                sectionNumber = RequirementTrackingWordService.ResolveNearestSectionNumber(doc, range.Start);
            }
            sectionNumber = TruncateSectionNumber(sectionNumber, extractionSectionNumberLevel);

            CustomExtractionMode mode = (CustomExtractionMode)Math.Max(0, cmbCustomExtractionMode.SelectedIndex);
            RequirementItem item;
            switch (mode)
            {
                case CustomExtractionMode.IdentifierOnly:
                    item = CreateCustomRequirement(range.Start, sectionNumber);
                    item.Id = text;
                    SetStatus($"已提取标识：{text}");
                    break;
                case CustomExtractionMode.NameOnly:
                    item = requirements.FirstOrDefault(candidate =>
                        candidate != null &&
                        !string.IsNullOrWhiteSpace(candidate.Id) &&
                        string.IsNullOrWhiteSpace(candidate.Name));
                    if (item == null)
                    {
                        SetStatus("没有待填写名称的标识；请先提取标识或使用交替方式");
                        return false;
                    }

                    item.Name = text;
                    item.SectionNumber = sectionNumber;
                    item.Start = range.Start;
                    item.BookmarkOrRange = range.Start;
                    SetStatus($"已提取名称：{text}");
                    break;
                default:
                    if (customAlternatingExpectsIdentifier)
                    {
                        item = CreateCustomRequirement(range.Start, sectionNumber);
                        item.Id = text;
                        customAlternatingExpectsIdentifier = false;
                        SetStatus($"已提取标识：{text}；下一次将提取名称");
                    }
                    else
                    {
                        item = requirements.LastOrDefault(candidate =>
                            candidate != null &&
                            !string.IsNullOrWhiteSpace(candidate.Id) &&
                            string.IsNullOrWhiteSpace(candidate.Name));
                        if (item == null)
                        {
                            item = CreateCustomRequirement(range.Start, sectionNumber);
                        }

                        item.Name = text;
                        item.SectionNumber = sectionNumber;
                        item.Start = range.Start;
                        item.BookmarkOrRange = range.Start;
                        customAlternatingExpectsIdentifier = true;
                        SetStatus($"已提取名称：{text}；下一次将提取标识");
                    }
                    break;
            }

            hasUnsavedChanges = true;
            RenderRows();
            SelectRow(requirements.IndexOf(item));
            return true;
        }

        private RequirementItem CreateCustomRequirement(int start, string sectionNumber)
        {
            RequirementItem item = new RequirementItem
            {
                Start = start,
                BookmarkOrRange = start,
                SectionNumber = sectionNumber ?? string.Empty
            };
            requirements.Add(item);
            return item;
        }

        private void DeleteSelectedRequirement()
        {
            int rowIndex = GetCurrentRowIndex();
            if (rowIndex < 0 || rowIndex >= requirements.Count)
            {
                SetStatus("请先选中要删除的需求标识");
                return;
            }

            RequirementItem item = requirements[rowIndex];
            string removedId = item.Id ?? string.Empty;
            if (!customMode && !string.IsNullOrWhiteSpace(removedId))
            {
                excludedRequirementIds.Add(removedId);
            }

            requirements.RemoveAt(rowIndex);
            hasUnsavedChanges = true;
            RenderRows();

            if (requirements.Count > 0)
            {
                SelectRow(Math.Min(rowIndex, requirements.Count - 1));
            }

            SetStatus($"已删除需求标识 {removedId}，点击“保存”后生效");
        }

        private void ClearAllRequirements()
        {
            if (requirements.Count == 0)
            {
                SetStatus("当前没有需求标识");
                return;
            }

            SyncGridToRequirements();
            int populatedCount = requirements.Count(item =>
                item != null &&
                (!string.IsNullOrWhiteSpace(item.Name) ||
                 !string.IsNullOrWhiteSpace(item.SectionNumber)));
            if (populatedCount == 0)
            {
                SetStatus("当前没有需要清空的提取内容");
                return;
            }

            DialogResult result = MessageBox.Show(
                $"确认清空 {populatedCount} 条需求的名称和章节号吗？需求标识会保留，点击“保存”后同步写入当前文档。",
                "全部清空",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
            {
                return;
            }

            foreach (RequirementItem item in requirements)
            {
                if (item != null)
                {
                    item.Name = string.Empty;
                    item.SectionNumber = string.Empty;
                }
            }

            RenderRows();
            hasUnsavedChanges = true;
            SelectRow(0);
            SetStatus($"已清空 {populatedCount} 条需求的名称和章节号，需求标识已保留；点击“保存”后写入当前文档");
        }

        private void ClearSelectedRequirementContent()
        {
            SyncGridToRequirements();
            int rowIndex = GetCurrentRowIndex();
            if (rowIndex < 0 || rowIndex >= requirements.Count)
            {
                SetStatus("请先选择需要清除的需求行");
                return;
            }

            RequirementItem item = requirements[rowIndex];
            if (item == null ||
                (string.IsNullOrWhiteSpace(item.Name) && string.IsNullOrWhiteSpace(item.SectionNumber)))
            {
                SetStatus("当前选中需求没有可清除的名称或章节号");
                return;
            }

            item.Name = string.Empty;
            item.SectionNumber = string.Empty;
            RenderRows();
            hasUnsavedChanges = true;
            SelectRow(rowIndex);
            SetStatus($"已清除需求标识 {item.Id} 的名称和章节号；点击“保存”后写入当前文档");
        }

        private void DeleteSavedResults()
        {
            Word.Document doc = currentDocument ?? applicationAccessor?.Invoke()?.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("当前没有可操作的 Word 文档。", "需求提取");
                return;
            }

            List<string> savedPartIds = FindSavedRequirementPartIds(doc);
            if (savedPartIds.Count == 0 && requirements.Count == 0)
            {
                SetStatus("当前没有可删除的需求提取结果");
                return;
            }

            DialogResult result = MessageBox.Show(
                "确认删除全部需求提取结果吗？需求标识、名称和章节号都会被删除，之后必须重新点击“加载标识”。",
                "删除",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
            {
                return;
            }

            try
            {
                SetExtractionEnabled(false);
                DeleteSavedRequirementParts(doc);
                int remainingCount = FindSavedRequirementPartIds(doc).Count;
                if (remainingCount > 0)
                {
                    MessageBox.Show(
                        $"仍有 {remainingCount} 份需求提取数据未能删除，请关闭其他需求窗格后重试。",
                        "删除",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    SetStatus("删除需求提取结果失败");
                    return;
                }

                currentDocument = doc;
                requirements.Clear();
                excludedRequirementIds.Clear();
                RenderRows();
                hasUnsavedChanges = false;
                UpdateExtractionActionLabel(doc);
                SetStatus("已删除全部需求提取结果；请重新点击“加载标识”");
            }
            catch (Exception ex)
            {
                SetStatus("删除需求提取结果失败");
                MessageBox.Show(
                    "删除需求提取结果失败：\r\n" + ex.Message,
                    "删除",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        internal void OpenExtractionSettings()
        {
            // Type detection inspects document paragraphs, so defer it until extraction actually runs.
            DocumentMarkerDocumentType currentDocumentType = extractionPresetTemplateType == DocumentMarkerDocumentType.Unknown
                ? DocumentMarkerDocumentType.RequirementSpecification
                : extractionPresetTemplateType;
            using (RequirementExtractionSettingsForm form = new RequirementExtractionSettingsForm(
                currentDocumentType,
                extractionPresetTemplateType,
                extractionUseCustomTemplates,
                GetConfiguredMarkerTemplates(),
                extractionMarkerPageRangeEnabled,
                extractionMarkerStartPage,
                extractionMarkerEndPage,
                extractionSectionNumberLevel,
                extractionScanFieldResults,
                Globals.ThisAddIn?.PreserveForwardMappingsWhenReverseTracing ?? true))
            {
                if (form.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                extractionMarkerTemplates = form.Templates.ToList();
                extractionUseCustomTemplates = form.UseCustomTemplates;
                extractionPresetTemplateType = form.SelectedPresetTemplateType;
                extractionMarkerPageRangeEnabled = form.LimitPages;
                extractionMarkerStartPage = form.StartPage;
                extractionMarkerEndPage = form.EndPage;
                extractionSectionNumberLevel = form.SectionTruncationLevel;
                extractionScanFieldResults = form.ScanFieldResults;
                if (Globals.ThisAddIn != null)
                {
                    Globals.ThisAddIn.PreserveForwardMappingsWhenReverseTracing =
                        form.PreserveForwardMappingsWhenReverseTracing;
                }
                SetStatus("提取设置仅在当前窗格打开期间生效；点击“加载标识”或“更新提取”后生效");
            }
        }

        internal void ResetOneTimeExtractionSettings()
        {
            extractionMarkerTemplates.Clear();
            extractionUseCustomTemplates = false;
            extractionPresetTemplateType = DocumentMarkerDocumentType.Unknown;
            extractionMarkerPageRangeEnabled = false;
            extractionMarkerStartPage = 1;
            extractionMarkerEndPage = 1;
            extractionSectionNumberLevel = 0;
            extractionScanFieldResults = false;
        }

        private List<string> GetConfiguredMarkerTemplates()
        {
            return extractionMarkerTemplates.ToList();
        }

        private DocumentMarkerDocumentType GetEffectivePresetTemplateType(Word.Document doc)
        {
            if (extractionPresetTemplateType == DocumentMarkerDocumentType.SystemSpecification ||
                extractionPresetTemplateType == DocumentMarkerDocumentType.RequirementSpecification ||
                extractionPresetTemplateType == DocumentMarkerDocumentType.SoftwareDesign)
            {
                return extractionPresetTemplateType;
            }

            DocumentMarkerDocumentType detectedType = DocumentMarkerService.DetectDocumentType(doc);
            return detectedType == DocumentMarkerDocumentType.SystemSpecification ||
                   detectedType == DocumentMarkerDocumentType.SoftwareDesign
                ? detectedType
                : DocumentMarkerDocumentType.RequirementSpecification;
        }

        private static string TruncateSectionNumber(string sectionNumber, int maxLevel)
        {
            string normalized = (sectionNumber ?? string.Empty).Trim();
            if (maxLevel <= 0 || string.IsNullOrWhiteSpace(normalized))
            {
                return normalized;
            }

            string[] levels = normalized
                .TrimEnd('.', '。', '、')
                .Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            if (levels.Length <= maxLevel || levels.Any(level => string.IsNullOrWhiteSpace(level)))
            {
                return normalized;
            }

            return string.Join(".", levels.Take(maxLevel));
        }

        internal void SaveToCurrentDocument()
        {
            Word.Document doc = currentDocument ?? applicationAccessor?.Invoke()?.ActiveDocument;
            if (doc == null)
            {
                MessageBox.Show("当前没有可保存的文档。", "需求提取");
                return;
            }

            SyncGridToRequirements();
            try
            {
                DeleteSavedRequirementParts(doc);
                doc.CustomXMLParts.Add(BuildSavedRequirementsXml());
                hasUnsavedChanges = false;
                UpdateExtractionActionLabel(doc);
                SetStatus("需求提取结果已保存到当前文档，请保存 Word 文档");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败: {ex.Message}", "需求提取");
            }
        }

        private void RenderRows()
        {
            suppressGridValueChanged = true;
            try
            {
                grid.Rows.Clear();
                foreach (RequirementItem item in requirements)
                {
                    grid.Rows.Add(item.Id ?? string.Empty, item.Name ?? string.Empty, item.SectionNumber ?? string.Empty);
                }
            }
            finally
            {
                suppressGridValueChanged = false;
            }
        }

        private void SyncGridToRequirements()
        {
            grid.EndEdit();
            for (int i = 0; i < requirements.Count && i < grid.Rows.Count; i++)
            {
                DataGridViewRow row = grid.Rows[i];
                requirements[i].Id = Convert.ToString(row.Cells[0].Value) ?? string.Empty;
                requirements[i].Name = Convert.ToString(row.Cells[1].Value) ?? string.Empty;
                requirements[i].SectionNumber = Convert.ToString(row.Cells[2].Value) ?? string.Empty;
            }
        }

        private void MergeSavedRequirements(IEnumerable<RequirementItem> savedItems)
        {
            List<RequirementItem> saved = (savedItems ?? Enumerable.Empty<RequirementItem>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                .ToList();
            HashSet<RequirementItem> matchedSaved = new HashSet<RequirementItem>();
            HashSet<RequirementItem> matchedCurrent = new HashSet<RequirementItem>();

            // First retain data for markers that did not change at all.
            Dictionary<string, Queue<RequirementItem>> savedById = saved
                .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => new Queue<RequirementItem>(group),
                    StringComparer.OrdinalIgnoreCase);
            foreach (RequirementItem current in requirements.Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id)))
            {
                Queue<RequirementItem> candidates;
                if (!savedById.TryGetValue(current.Id, out candidates) || candidates.Count == 0)
                {
                    continue;
                }

                RequirementItem previous = candidates.Dequeue();
                CopySavedRequirementData(previous, current);
                matchedSaved.Add(previous);
                matchedCurrent.Add(current);
            }

            // When inserted markers shift sequence numbers, preserve the old data by
            // matching the remaining entries within the same marker prefix.
            List<RequirementItem> remainingSaved = saved.Where(item => !matchedSaved.Contains(item)).ToList();
            List<RequirementItem> remainingCurrent = requirements
                .Where(item => item != null && !matchedCurrent.Contains(item) && !string.IsNullOrWhiteSpace(item.Id))
                .ToList();
            foreach (IGrouping<string, RequirementItem> savedGroup in remainingSaved
                .GroupBy(item => GetMarkerPrefix(item.Id), StringComparer.OrdinalIgnoreCase))
            {
                string prefix = savedGroup.Key;
                List<RequirementItem> currentGroup = remainingCurrent
                    .Where(item => string.Equals(GetMarkerPrefix(item.Id), prefix, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (string.IsNullOrWhiteSpace(prefix) || currentGroup.Count == 0)
                {
                    continue;
                }

                foreach (RequirementItem previous in savedGroup)
                {
                    RequirementItem current = SelectClosestMarkerNumber(previous, currentGroup);
                    if (current == null)
                    {
                        break;
                    }

                    CopySavedRequirementData(previous, current);
                    currentGroup.Remove(current);
                }
            }
        }

        private void UpdateExtractionActionLabel(Word.Document doc)
        {
            bool hasExtractedMarkers = requirements.Any(item =>
                item != null && !string.IsNullOrWhiteSpace(item.Id));
            if (!hasExtractedMarkers && doc != null)
            {
                Office.CustomXMLPart savedPart = FindSavedRequirementPart(doc);
                hasExtractedMarkers = savedPart != null && ContainsSavedRequirement(savedPart);
            }

            btnRefresh.Text = hasExtractedMarkers ? "更新提取" : "加载标识";
            toolTip.SetToolTip(
                btnRefresh,
                hasExtractedMarkers
                    ? "按当前文档类型更新标识，并保留已提取的内容"
                    : "按当前文档类型扫描并加载标识");
        }

        private static void CopySavedRequirementData(RequirementItem source, RequirementItem target)
        {
            if (source == null || target == null)
            {
                return;
            }

            target.Name = source.Name ?? string.Empty;
            target.SectionNumber = source.SectionNumber ?? string.Empty;
        }

        private static RequirementItem SelectClosestMarkerNumber(
            RequirementItem previous,
            IEnumerable<RequirementItem> candidates)
        {
            string previousPrefix;
            long previousNumber;
            bool hasPreviousNumber = TryGetMarkerNumber(previous?.Id, out previousPrefix, out previousNumber);
            return (candidates ?? Enumerable.Empty<RequirementItem>())
                .Select((candidate, index) => new
                {
                    Candidate = candidate,
                    Index = index,
                    Distance = hasPreviousNumber && TryGetMarkerNumber(candidate?.Id, out _, out long candidateNumber)
                        ? Math.Abs(candidateNumber - previousNumber)
                        : long.MaxValue
                })
                .OrderBy(item => item.Distance)
                .ThenBy(item => item.Index)
                .Select(item => item.Candidate)
                .FirstOrDefault();
        }

        private static string GetMarkerPrefix(string markerId)
        {
            string prefix;
            long number;
            return TryGetMarkerNumber(markerId, out prefix, out number) ? prefix : string.Empty;
        }

        private static bool TryGetMarkerNumber(string markerId, out string prefix, out long number)
        {
            prefix = string.Empty;
            number = 0;
            string value = (markerId ?? string.Empty).Trim().TrimStart('/');
            int start = value.Length;
            while (start > 0 && char.IsDigit(value[start - 1]))
            {
                start--;
            }

            if (start == value.Length || start == 0 || !long.TryParse(value.Substring(start), out number))
            {
                return false;
            }

            prefix = value.Substring(0, start).TrimEnd('-').Trim();
            return prefix.Length > 0;
        }

        private string BuildSavedRequirementsXml()
        {
            XElement root = new XElement(XName.Get("requirements", SavedRequirementsNamespace));
            foreach (RequirementItem item in requirements.Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id)))
            {
                root.Add(new XElement(XName.Get("requirement", SavedRequirementsNamespace),
                    new XAttribute("id", item.Id ?? string.Empty),
                    new XAttribute("name", item.Name ?? string.Empty),
                    new XAttribute("section", item.SectionNumber ?? string.Empty)));
            }

            foreach (string id in excludedRequirementIds.OrderBy(id => id, StringComparer.OrdinalIgnoreCase))
            {
                root.Add(new XElement(XName.Get("excluded", SavedRequirementsNamespace),
                    new XAttribute("id", id)));
            }

            return root.ToString(SaveOptions.DisableFormatting);
        }

        private static HashSet<string> LoadExcludedRequirementIds(Word.Document doc)
        {
            HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Office.CustomXMLPart part = FindSavedRequirementPart(doc);
            if (part == null)
            {
                return result;
            }

            try
            {
                XDocument document = XDocument.Parse(part.XML);
                XName excludedName = XName.Get("excluded", SavedRequirementsNamespace);
                foreach (XElement element in document.Descendants(excludedName))
                {
                    string id = (string)element.Attribute("id") ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(id))
                    {
                        result.Add(id);
                    }
                }
            }
            catch
            {
            }

            return result;
        }

        internal static List<RequirementItem> LoadSavedRequirementItems(Word.Document doc)
        {
            List<RequirementItem> result = new List<RequirementItem>();
            Office.CustomXMLPart part = FindSavedRequirementPart(doc);
            if (part == null)
            {
                return result;
            }

            try
            {
                XDocument document = XDocument.Parse(part.XML);
                XName itemName = XName.Get("requirement", SavedRequirementsNamespace);
                foreach (XElement element in document.Descendants(itemName))
                {
                    string id = (string)element.Attribute("id") ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(id))
                    {
                        continue;
                    }

                    result.Add(new RequirementItem
                    {
                        Id = id,
                        Name = (string)element.Attribute("name") ?? string.Empty,
                        SectionNumber = (string)element.Attribute("section") ?? string.Empty
                    });
                }
            }
            catch
            {
                result.Clear();
            }

            return result;
        }

        private void RestoreRequirementPositions(Word.Document doc)
        {
            if (doc == null || requirements.Count == 0)
            {
                return;
            }

            try
            {
                Dictionary<string, Queue<int>> positionsById =
                    new Dictionary<string, Queue<int>>(StringComparer.OrdinalIgnoreCase);
                DocumentMarkerCollectionResult markerResult = DocumentMarkerService.CollectMarkers(doc);
                foreach (NavigationPaneEntry entry in markerResult?.Entries ?? new List<NavigationPaneEntry>())
                {
                    if (entry == null || string.IsNullOrWhiteSpace(entry.Text) || entry.Start < 0)
                    {
                        continue;
                    }

                    string key = NormalizeRequirementIdKey(entry.Text);
                    if (!positionsById.TryGetValue(key, out Queue<int> positions))
                    {
                        positions = new Queue<int>();
                        positionsById[key] = positions;
                    }

                    positions.Enqueue(entry.Start);
                }

                foreach (RequirementItem item in requirements)
                {
                    if (item == null || string.IsNullOrWhiteSpace(item.Id))
                    {
                        continue;
                    }

                    string key = NormalizeRequirementIdKey(item.Id);
                    if (positionsById.TryGetValue(key, out Queue<int> positions) && positions.Count > 0)
                    {
                        item.Start = positions.Dequeue();
                        item.BookmarkOrRange = item.Start;
                        continue;
                    }

                    int fallbackStart = FindRequirementPosition(doc, item.Id);
                    if (fallbackStart >= 0)
                    {
                        item.Start = fallbackStart;
                        item.BookmarkOrRange = fallbackStart;
                    }
                }
            }
            catch
            {
                // 定位恢复失败不应阻止需求提取窗格显示；单项查找仍可在后续补救。
            }
        }

        private static int FindRequirementPosition(Word.Document doc, string requirementId)
        {
            if (doc?.Content == null || string.IsNullOrWhiteSpace(requirementId))
            {
                return -1;
            }

            foreach (string candidate in new[]
            {
                requirementId.Trim(),
                requirementId.Trim().TrimStart('/')
            }.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    Word.Range searchRange = doc.Content.Duplicate;
                    Word.Find find = searchRange.Find;
                    find.ClearFormatting();
                    find.Replacement.ClearFormatting();
                    find.Text = candidate;
                    find.Forward = true;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    find.Format = false;
                    find.MatchCase = false;
                    find.MatchWholeWord = false;
                    if (find.Execute())
                    {
                        return searchRange.Start;
                    }
                }
                catch
                {
                }
            }

            return -1;
        }

        private static string NormalizeRequirementIdKey(string value)
        {
            return (value ?? string.Empty).Trim().TrimStart('/');
        }

        private static void DeleteSavedRequirementParts(Word.Document doc)
        {
            List<string> partIds = FindSavedRequirementPartIds(doc);
            foreach (string partId in partIds)
            {
                try
                {
                    Office.CustomXMLPart part = FindSavedRequirementPartById(doc, partId);
                    part?.Delete();
                }
                catch
                {
                }
            }

            // Word can refresh the COM collection after a delete. Repeat a bounded
            // cleanup pass so a stale duplicate cannot survive the next save.
            for (int attempt = 0; attempt < 3; attempt++)
            {
                List<string> remaining = FindSavedRequirementPartIds(doc);
                if (remaining.Count == 0)
                {
                    return;
                }

                foreach (string partId in remaining)
                {
                    try
                    {
                        FindSavedRequirementPartById(doc, partId)?.Delete();
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static List<string> FindSavedRequirementPartIds(Word.Document doc)
        {
            return FindSavedRequirementParts(doc)
                .Select(part =>
                {
                    try
                    {
                        return part?.Id ?? string.Empty;
                    }
                    catch
                    {
                        return string.Empty;
                    }
                })
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static Office.CustomXMLPart FindSavedRequirementPartById(Word.Document doc, string partId)
        {
            if (doc == null || string.IsNullOrWhiteSpace(partId))
            {
                return null;
            }

            foreach (Office.CustomXMLPart part in FindSavedRequirementParts(doc))
            {
                try
                {
                    if (string.Equals(part.Id, partId, StringComparison.OrdinalIgnoreCase))
                    {
                        return part;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static Office.CustomXMLPart FindSavedRequirementPart(Word.Document doc)
        {
            List<Office.CustomXMLPart> parts = FindSavedRequirementParts(doc);
            for (int index = parts.Count - 1; index >= 0; index--)
            {
                if (ContainsSavedRequirement(parts[index]))
                {
                    return parts[index];
                }
            }

            return parts.Count > 0 ? parts[parts.Count - 1] : null;
        }

        private static bool ContainsSavedRequirement(Office.CustomXMLPart part)
        {
            try
            {
                XDocument document = XDocument.Parse(part?.XML ?? string.Empty);
                XName itemName = XName.Get("requirement", SavedRequirementsNamespace);
                return document.Descendants(itemName).Any(element =>
                    !string.IsNullOrWhiteSpace((string)element.Attribute("id")) ||
                    !string.IsNullOrWhiteSpace((string)element.Attribute("name")));
            }
            catch
            {
                return false;
            }
        }

        private static List<Office.CustomXMLPart> FindSavedRequirementParts(Word.Document doc)
        {
            List<Office.CustomXMLPart> result = new List<Office.CustomXMLPart>();
            HashSet<string> partIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (doc == null)
            {
                return result;
            }

            try
            {
                Office.CustomXMLParts parts = doc.CustomXMLParts.SelectByNamespace(SavedRequirementsNamespace);
                AddMatchingSavedParts(parts, result, partIds);
            }
            catch
            {
            }

            // Some Office builds do not reliably return custom XML parts through
            // SelectByNamespace. Enumerate all parts as a compatibility fallback.
            try
            {
                AddMatchingSavedParts(doc.CustomXMLParts, result, partIds);
            }
            catch
            {
            }

            return result;
        }

        private static void AddMatchingSavedParts(
            Office.CustomXMLParts parts,
            ICollection<Office.CustomXMLPart> result,
            ISet<string> partIds)
        {
            if (parts == null)
            {
                return;
            }

            for (int index = 1; index <= parts.Count; index++)
            {
                Office.CustomXMLPart part = null;
                try
                {
                    part = parts[index];
                    if (!IsSavedRequirementPart(part))
                    {
                        continue;
                    }

                    string id = part.Id ?? string.Empty;
                    if (id.Length == 0 || partIds.Add(id))
                    {
                        result.Add(part);
                    }
                }
                catch
                {
                }
            }
        }

        private static bool IsSavedRequirementPart(Office.CustomXMLPart part)
        {
            try
            {
                string xml = part?.XML;
                if (string.IsNullOrWhiteSpace(xml))
                {
                    return false;
                }

                XDocument document = XDocument.Parse(xml);
                XElement root = document.Root;
                return root != null &&
                       string.Equals(root.Name.LocalName, "requirements", StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(root.Name.NamespaceName, SavedRequirementsNamespace, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }

        private int GetCurrentRowIndex()
        {
            if (grid.CurrentRow != null && grid.CurrentRow.Index >= 0 && grid.CurrentRow.Index < requirements.Count)
            {
                return grid.CurrentRow.Index;
            }

            return grid.SelectedRows.Count > 0 ? grid.SelectedRows[0].Index : -1;
        }

        private void Grid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right || e.RowIndex < 0 || e.RowIndex >= grid.Rows.Count)
            {
                return;
            }

            grid.ClearSelection();
            grid.Rows[e.RowIndex].Selected = true;
            grid.CurrentCell = grid.Rows[e.RowIndex].Cells[Math.Max(0, e.ColumnIndex)];
        }

        private void Grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (suppressGridValueChanged ||
                e.RowIndex < 0 ||
                e.RowIndex >= requirements.Count ||
                (e.ColumnIndex != 0 && e.ColumnIndex != 1 && e.ColumnIndex != 2))
            {
                return;
            }

            RequirementItem item = requirements[e.RowIndex];
            if (e.ColumnIndex == 0)
            {
                if (!customMode)
                {
                    return;
                }

                item.Id = Convert.ToString(grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value) ?? string.Empty;
            }
            else if (e.ColumnIndex == 1)
            {
                item.Name = Convert.ToString(grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value) ?? string.Empty;
            }
            else
            {
                item.SectionNumber = Convert.ToString(grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value) ?? string.Empty;
            }

            hasUnsavedChanges = true;
            SetStatus($"已修改 {item.Id}，点击“保存”后写入当前文档");
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Delete)
            {
                return;
            }

            DeleteSelectedRequirement();
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        private int FindNextUnfilledRow(int start)
        {
            for (int i = Math.Max(0, start); i < requirements.Count; i++)
            {
                if (string.IsNullOrWhiteSpace(requirements[i].Name))
                {
                    return i;
                }
            }

            return -1;
        }

        private void SelectRowAndNavigate(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= requirements.Count)
            {
                SetStatus("需求名称已全部填写完成");
                return;
            }

            SelectRow(rowIndex);
            NavigateToRequirement(requirements[rowIndex]);
        }

        private void NavigateToRequirement(RequirementItem item)
        {
            if (item == null)
            {
                return;
            }

            try
            {
                currentDocument?.Activate();
            }
            catch
            {
            }

            int position = item.Start;
            try
            {
                int contentStart = currentDocument?.Content?.Start ?? 0;
                int contentEnd = currentDocument?.Content?.End ?? 0;
                if (position <= contentStart || position >= contentEnd)
                {
                    int recoveredPosition = FindRequirementPosition(currentDocument, item.Id);
                    if (recoveredPosition >= 0)
                    {
                        position = recoveredPosition;
                        item.Start = recoveredPosition;
                        item.BookmarkOrRange = recoveredPosition;
                    }
                }
            }
            catch
            {
            }

            RequirementActivated?.Invoke(NormalizeDocumentPosition(currentDocument, position));
        }

        private static int NormalizeDocumentPosition(Word.Document doc, int position)
        {
            if (doc == null)
            {
                return Math.Max(0, position);
            }

            try
            {
                int start = doc.Content.Start;
                int end = Math.Max(start, doc.Content.End - 1);
                return Math.Max(start, Math.Min(position, end));
            }
            catch
            {
                return Math.Max(0, position);
            }
        }

        private void SelectRow(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= grid.Rows.Count)
            {
                return;
            }

            grid.ClearSelection();
            grid.Rows[rowIndex].Selected = true;
            grid.CurrentCell = grid.Rows[rowIndex].Cells[0];
            try
            {
                grid.FirstDisplayedScrollingRowIndex = Math.Max(0, rowIndex - 2);
            }
            catch
            {
            }
        }

        private void SetStatus(string message)
        {
            lblStatus.Text = string.IsNullOrWhiteSpace(message) ? string.Empty : message;
            try
            {
                Application.DoEvents();
            }
            catch
            {
            }
        }

        private static Button CreateButton(string text)
        {
            Button button = new Button
            {
                AutoSize = true,
                BackColor = Color.FromArgb(43, 108, 176),
                FlatStyle = FlatStyle.Flat,
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                Text = text,
                Margin = new Padding(0, 0, 8, 8),
                Padding = Padding.Empty,
                TextAlign = ContentAlignment.MiddleCenter,
                UseCompatibleTextRendering = false
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private static Label CreateToolbarLabel(string text)
        {
            return new Label
            {
                AutoSize = true,
                Text = text ?? string.Empty,
                Margin = new Padding(4, 8, 4, 0),
                ForeColor = Color.FromArgb(70, 78, 90),
                Font = new Font("Microsoft YaHei UI", 9F)
            };
        }

        internal void ConfirmSaveBeforeClosing()
        {
            if (!hasUnsavedChanges)
            {
                return;
            }

            DialogResult result = MessageBox.Show(
                "需求提取结果尚未保存，是否现在保存到当前 Word 文档？",
                "需求提取",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                SaveToCurrentDocument();
            }
            else
            {
                // 放弃窗格中的未保存加载或编辑结果，恢复到文档中已保存的内容。
                LoadSavedRequirementsFromCurrentDocument();
            }
        }

        private void ApplyExtractionButtonState()
        {
            btnExtractionEnabled.Text = extractionEnabled ? "结束批量提取" : "开始批量提取";
            btnExtractionEnabled.BackColor = extractionEnabled ? Color.FromArgb(35, 138, 79) : Color.FromArgb(235, 243, 255);
            btnExtractionEnabled.ForeColor = extractionEnabled ? Color.White : Color.FromArgb(38, 90, 166);
            btnBatchExtraction.Text = batchExtractionEnabled ? "结束批量提取" : "开始批量提取";
            btnBatchExtraction.BackColor = batchExtractionEnabled ? Color.FromArgb(35, 138, 79) : Color.FromArgb(235, 243, 255);
            btnBatchExtraction.ForeColor = batchExtractionEnabled ? Color.White : Color.FromArgb(38, 90, 166);
        }

        private static void ApplySecondaryButtonStyle(Button button)
        {
            button.BackColor = Color.FromArgb(235, 243, 255);
            button.ForeColor = Color.FromArgb(38, 90, 166);
            button.FlatAppearance.BorderColor = Color.FromArgb(176, 203, 240);
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(223, 236, 255);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(209, 227, 253);
        }

        private static void PlaceToolbarButton(Button button, TableLayoutPanel container, int column, int row)
        {
            button.AutoSize = false;
            button.Dock = DockStyle.Fill;
            button.Margin = new Padding(4, 4, 4, 4);
            container.Controls.Add(button, column, row);
        }

        private static void PlaceToolbarButton(Button button, FlowLayoutPanel container, int width)
        {
            button.AutoSize = false;
            button.Size = new Size(width, 32);
            button.Margin = new Padding(0, 1, 8, 1);
            container.Controls.Add(button);
        }

        private static string GetCurrentDocumentDisplayName(Word.Document doc)
        {
            try
            {
                string sample = (doc?.Name ?? string.Empty) + " " + GetDocumentOpeningText(doc);
                if (ContainsAny(sample, "软件需求规格说明", "需求规格说明", "SRS"))
                {
                    return "软件需求规格说明";
                }

                if (ContainsAny(sample, "软件概要设计说明", "概要设计说明", "SDS"))
                {
                    return "软件概要设计说明";
                }

                if (ContainsAny(sample, "软件详细设计说明", "详细设计说明", "SDD"))
                {
                    return "软件详细设计说明";
                }

                if (ContainsAny(sample, "软件设计说明", "软件设计描述"))
                {
                    return "软件设计说明";
                }

                if (ContainsAny(sample, "软件测试说明", "软件测试描述", "测试说明", "STD", "STS"))
                {
                    return "软件测试说明";
                }

                if (ContainsAny(sample, "软件研制任务书", "研制任务书"))
                {
                    return "软件研制任务书";
                }

                if (ContainsAny(sample, "系统规格说明", "系统/子系统规格说明", "系统子系统规格说明", "SSS"))
                {
                    return "系统规格说明";
                }
            }
            catch
            {
            }

            return "软件需求规格说明";
        }

        private static string GetDocumentOpeningText(Word.Document doc)
        {
            if (doc == null || doc.Content == null)
            {
                return string.Empty;
            }

            int end = Math.Min(doc.Content.End, 3000);
            if (end <= 0)
            {
                return string.Empty;
            }

            return doc.Range(0, end).Text ?? string.Empty;
        }

        private static bool ContainsAny(string text, params string[] tokens)
        {
            return !string.IsNullOrWhiteSpace(text) && tokens != null &&
                   tokens.Any(token => !string.IsNullOrWhiteSpace(token) &&
                                       text.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static DataGridViewTextBoxColumn CreateTextColumn(
            string propertyName,
            string headerText,
            bool readOnly = false,
            int minimumWidth = 80)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = propertyName,
                HeaderText = headerText,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 1f,
                MinimumWidth = minimumWidth,
                ReadOnly = readOnly,
                Resizable = DataGridViewTriState.True,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private static void EnableDoubleBuffering(Control control)
        {
            try
            {
                control?.GetType()
                    .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    ?.SetValue(control, true, null);
            }
            catch
            {
            }
        }

        private static string CleanText(string text)
        {
            return (text ?? string.Empty).Replace("\r", string.Empty).Replace("\a", string.Empty).Trim();
        }
    }
}
