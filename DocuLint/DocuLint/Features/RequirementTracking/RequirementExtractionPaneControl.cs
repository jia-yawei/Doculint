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
        private readonly Button btnRefresh;
        private readonly Button btnExtractionEnabled;
        private readonly Button btnBatchExtraction;
        private readonly Button btnClearCurrent;
        private readonly Button btnClearAll;
        private readonly Button btnSave;
        private readonly Button btnExportTraceTable;
        private readonly Label lblStatus;
        private readonly ToolTip toolTip;
        private Word.Document currentDocument;
        private readonly List<RequirementItem> requirements = new List<RequirementItem>();
        private bool extractionEnabled;
        private bool batchExtractionEnabled;
        private bool hasUnsavedChanges;

        internal event Action<int> RequirementActivated;
        internal event Action<bool> BatchExtractionModeChanged;
        internal event Action<bool> ExtractionModeChanged;

        internal RequirementExtractionPaneControl(Func<Word.Application> applicationAccessor)
        {
            this.applicationAccessor = applicationAccessor;
            Dock = DockStyle.Fill;
            BackColor = Color.White;
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
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 92f));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Panel toolbar = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = false
            };
            EnableDoubleBuffering(toolbar);

            TableLayoutPanel primaryActions = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 42,
                BackColor = Color.FromArgb(245, 247, 250),
                ColumnCount = 4,
                RowCount = 1
            };
            primaryActions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 108f));
            primaryActions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            primaryActions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 80f));
            primaryActions.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 128f));

            FlowLayoutPanel modeActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 38,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 2, 0, 2)
            };
            toolbar.Controls.Add(primaryActions);
            toolbar.Controls.Add(modeActions);

            btnRefresh = CreateButton("加载标识");
            btnRefresh.Click += (_, __) => LoadCurrentDocumentRequirements();
            btnExtractionEnabled = CreateButton("开始提取");
            btnExtractionEnabled.Click += (_, __) => SetExtractionEnabled(!extractionEnabled);
            btnBatchExtraction = CreateButton("批量提取");
            btnBatchExtraction.Click += (_, __) => SetBatchExtractionMode(!batchExtractionEnabled);
            btnClearCurrent = CreateButton("清除选中");
            btnClearCurrent.Click += (_, __) => ClearCurrentRequirementName();
            btnClearAll = CreateButton("一键清空");
            btnClearAll.BackColor = Color.FromArgb(177, 63, 63);
            btnClearAll.FlatAppearance.MouseOverBackColor = Color.FromArgb(196, 76, 76);
            btnClearAll.FlatAppearance.MouseDownBackColor = Color.FromArgb(151, 49, 49);
            btnClearAll.Click += (_, __) => ClearAllRequirements();
            btnSave = CreateButton("保存");
            btnSave.Click += (_, __) => SaveToCurrentDocument();
            btnExportTraceTable = CreateButton("导出追踪表");
            btnExportTraceTable.Click += (_, __) => ExportTraceTable();
            ApplySecondaryButtonStyle(btnExtractionEnabled);
            ApplySecondaryButtonStyle(btnBatchExtraction);
            ApplySecondaryButtonStyle(btnClearCurrent);
            ApplySecondaryButtonStyle(btnExportTraceTable);
            PlaceToolbarButton(btnRefresh, primaryActions, 0, 0);
            PlaceToolbarButton(btnSave, primaryActions, 2, 0);
            PlaceToolbarButton(btnExportTraceTable, primaryActions, 3, 0);
            PlaceToolbarButton(btnExtractionEnabled, modeActions, 100);
            PlaceToolbarButton(btnBatchExtraction, modeActions, 156);
            PlaceToolbarButton(btnClearCurrent, modeActions, 100);
            PlaceToolbarButton(btnClearAll, modeActions, 100);
            toolTip.SetToolTip(btnRefresh, "从当前 Word 文档加载需求标识");
            toolTip.SetToolTip(btnExtractionEnabled, "开启后，在文档中选中文字可进行提取");
            toolTip.SetToolTip(btnBatchExtraction, "开启后，选中文字会自动写入当前需求行");
            toolTip.SetToolTip(btnClearCurrent, "清除当前选中行的需求名称和章节号");
            toolTip.SetToolTip(btnClearAll, "清空当前需求表中的全部内容");
            toolTip.SetToolTip(btnSave, "将当前需求表保存到 Word 文档");
            toolTip.SetToolTip(btnExportTraceTable, "在当前 Word 光标位置插入需求追踪表");

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeColumns = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                ColumnHeadersHeight = 34,
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(230, 235, 242),
                ReadOnly = true,
                RowHeadersVisible = false,
                RowTemplate = { Height = 30 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(238, 242, 247);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(32, 45, 64);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 238, 255);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            EnableDoubleBuffering(grid);
            grid.Columns.Add(CreateTextColumn("Id", "需求标识"));
            grid.Columns.Add(CreateTextColumn("Name", "需求名称"));
            grid.Columns.Add(CreateTextColumn("SectionNumber", "章节号"));
            grid.CellDoubleClick += (_, e) =>
            {
                if (e.RowIndex >= 0 && e.RowIndex < requirements.Count)
                {
                    NavigateToRequirement(requirements[e.RowIndex]);
                }
            };

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

        internal void LoadCurrentDocumentRequirements()
        {
            Word.Application app = applicationAccessor?.Invoke();
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                SetStatus("当前没有活动文档");
                return;
            }

            currentDocument = doc;
            SetStatus("正在提取需求标识...");
            try
            {
                DocumentMarkerCollectionResult markerResult = DocumentMarkerService.CollectMarkers(doc);
                requirements.Clear();
                requirements.AddRange((markerResult.Entries ?? new List<NavigationPaneEntry>())
                    .Where(entry => entry != null && !string.IsNullOrWhiteSpace(entry.Text))
                    .Select(entry => new RequirementItem
                    {
                        Id = entry.Text,
                        Name = string.Empty,
                        SectionNumber = string.Empty,
                        Start = entry.Start,
                        BookmarkOrRange = entry.Start
                    }));
                MergeSavedRequirements(doc);
                RenderRows();
                hasUnsavedChanges = true;
                btnRefresh.Text = "重新加载";
                ApplyBatchExtractionState(batchExtractionEnabled);
                SetStatus($"已加载 {requirements.Count} 个需求标识，勾选“开始提取”后可填写需求名称");
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
                MessageBox.Show("请先勾选“开始提取”。", "需求提取");
                return;
            }

            TryAddSelectionAsRequirement(false);
        }

        internal bool TryAutoAddSelection()
        {
            return extractionEnabled && TryAddSelectionAsRequirement(true);
        }

        internal bool ExtractionEnabled => extractionEnabled;

        internal bool BatchExtractionEnabled => extractionEnabled && batchExtractionEnabled;

        private void SetExtractionEnabled(bool enabled)
        {
            if (extractionEnabled == enabled)
            {
                return;
            }

            extractionEnabled = enabled;
            btnBatchExtraction.Enabled = enabled;
            if (!enabled)
            {
                SetBatchExtractionMode(false);
            }

            ExtractionModeChanged?.Invoke(enabled);
            ApplyExtractionButtonState();

            SetStatus(enabled
                ? "开始提取已开启：在左侧文档中选中文字后可点击“提取”"
                : "开始提取已关闭");
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
                : "批量提取已关闭：选中文字后可手动点击“提取”");
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

        internal void ClearCurrentRequirementName()
        {
            int rowIndex = GetCurrentRowIndex();
            if (rowIndex < 0)
            {
                MessageBox.Show("请先在需求窗格中选中要清除的行。", "需求提取");
                return;
            }

            RequirementItem item = requirements[rowIndex];
            item.Name = string.Empty;
            item.SectionNumber = string.Empty;
            hasUnsavedChanges = true;
            RenderRows();
            SelectRow(rowIndex);
            SetStatus($"已清除 {item.Id} 的需求名称和章节号");
        }

        private void ClearAllRequirements()
        {
            if (requirements.Count == 0)
            {
                SetStatus("当前没有可清空的需求");
                return;
            }

            DialogResult result = MessageBox.Show(
                "确认清空当前窗格中的全部需求吗？点击“保存”后会同步写入当前文档。",
                "全部清空",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
            {
                return;
            }

            requirements.Clear();
            RenderRows();
            hasUnsavedChanges = true;
            SetStatus("已清空全部需求，点击“保存”后写入当前文档");
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
                SetStatus("需求提取结果已保存到当前文档，请保存 Word 文档");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败: {ex.Message}", "需求提取");
            }
        }

        private void ExportTraceTable()
        {
            Word.Application app = applicationAccessor?.Invoke();
            Word.Selection selection = app?.Selection;
            Word.Document doc = app?.ActiveDocument ?? currentDocument;
            if (selection == null || doc == null)
            {
                MessageBox.Show("当前没有可插入表格的 Word 光标位置。", "需求提取");
                return;
            }

            SyncGridToRequirements();
            IList<RequirementTraceExportRow> rows = RequirementTraceTableExporter.BuildSourceOnlyRows(requirements);
            if (rows.Count == 0)
            {
                MessageBox.Show("当前没有可导出的需求。请先提取需求标识。", "需求提取");
                return;
            }

            DialogResult result = MessageBox.Show(
                "确认把需求追踪表导出到当前 Word 光标位置？",
                "导出追踪表",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            if (result != DialogResult.OK)
            {
                return;
            }

            currentDocument = doc;
            try
            {
                RequirementTraceTableExporter.InsertTraceTable(
                    doc,
                    selection.Range,
                    rows,
                    GetCurrentDocumentDisplayName(doc),
                    string.Empty,
                    false);
                SetStatus("需求追踪表已导出到当前光标位置");
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出需求追踪表失败：\r\n" + ex.Message, "需求提取");
            }
        }

        private void RenderRows()
        {
            grid.Rows.Clear();
            foreach (RequirementItem item in requirements)
            {
                grid.Rows.Add(item.Id ?? string.Empty, item.Name ?? string.Empty, item.SectionNumber ?? string.Empty);
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

        private void MergeSavedRequirements(Word.Document doc)
        {
            Dictionary<string, RequirementItem> saved = LoadSavedRequirements(doc);
            foreach (RequirementItem item in requirements)
            {
                if (item == null || string.IsNullOrWhiteSpace(item.Id))
                {
                    continue;
                }

                RequirementItem savedItem;
                if (saved.TryGetValue(item.Id, out savedItem))
                {
                    item.Name = savedItem.Name;
                    item.SectionNumber = savedItem.SectionNumber;
                }
            }
        }

        private string BuildSavedRequirementsXml()
        {
            XElement root = new XElement(XName.Get("requirements", SavedRequirementsNamespace),
                requirements
                    .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                    .Select(item =>
                        new XElement(XName.Get("requirement", SavedRequirementsNamespace),
                            new XAttribute("id", item.Id ?? string.Empty),
                            new XAttribute("name", item.Name ?? string.Empty),
                            new XAttribute("section", item.SectionNumber ?? string.Empty))));

            return root.ToString(SaveOptions.DisableFormatting);
        }

        private static Dictionary<string, RequirementItem> LoadSavedRequirements(Word.Document doc)
        {
            Dictionary<string, RequirementItem> result = new Dictionary<string, RequirementItem>(StringComparer.OrdinalIgnoreCase);
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

                    result[id] = new RequirementItem
                    {
                        Id = id,
                        Name = (string)element.Attribute("name") ?? string.Empty,
                        SectionNumber = (string)element.Attribute("section") ?? string.Empty
                    };
                }
            }
            catch
            {
            }

            return result;
        }

        internal static List<RequirementItem> LoadSavedRequirementItems(Word.Document doc)
        {
            return LoadSavedRequirements(doc)
                .Values
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                .OrderBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void DeleteSavedRequirementParts(Word.Document doc)
        {
            Office.CustomXMLPart part;
            while ((part = FindSavedRequirementPart(doc)) != null)
            {
                part.Delete();
            }
        }

        private static Office.CustomXMLPart FindSavedRequirementPart(Word.Document doc)
        {
            if (doc == null)
            {
                return null;
            }

            try
            {
                Office.CustomXMLParts parts = doc.CustomXMLParts.SelectByNamespace(SavedRequirementsNamespace);
                return parts != null && parts.Count > 0 ? parts[1] : null;
            }
            catch
            {
                return null;
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

            RequirementActivated?.Invoke(NormalizeDocumentPosition(currentDocument, item.Start));
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
        }

        private void ApplyExtractionButtonState()
        {
            btnExtractionEnabled.Text = extractionEnabled ? "结束提取" : "开始提取";
            btnExtractionEnabled.BackColor = extractionEnabled ? Color.FromArgb(35, 138, 79) : Color.FromArgb(235, 243, 255);
            btnExtractionEnabled.ForeColor = extractionEnabled ? Color.White : Color.FromArgb(38, 90, 166);
            btnBatchExtraction.Text = batchExtractionEnabled ? "批量提取：开" : "批量提取";
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

        private static DataGridViewTextBoxColumn CreateTextColumn(string propertyName, string headerText, bool readOnly = false)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = propertyName,
                HeaderText = headerText,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = 1f,
                ReadOnly = readOnly,
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
