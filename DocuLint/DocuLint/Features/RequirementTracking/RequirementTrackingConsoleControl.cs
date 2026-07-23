using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace DocuLint
{
    internal sealed class RequirementTrackingConsoleControl : UserControl
    {
        private const string SavedTraceMappingsNamespace = "urn:doculint:requirement-tracking";

        private enum FilterMode
        {
            All,
            Mapped,
            Unmapped
        }

        private readonly Func<Word.Application> applicationAccessor;
        private readonly Button btnImportTarget;
        private readonly Button btnExportTable;
        private readonly Button btnClearTraceMappings;
        private readonly ComboBox cmbTraceTemplate;
        private readonly Label lblTraceTemplate;
        private readonly Label lblCustomSourceTitle;
        private readonly TextBox txtCustomSourceTitle;
        private readonly Label lblCustomTargetTitle;
        private readonly TextBox txtCustomTargetTitle;
        private readonly Button btnPreviousSource;
        private readonly Button btnNextSource;
        private readonly Button btnNextUnmappedSource;
        private readonly ComboBox cmbSourceFilter;
        private readonly ComboBox cmbTargetFilter;
        private readonly TextBox txtTargetSearch;
        private readonly DataGridView gridSource;
        private readonly DataGridView gridTargetRecommended;
        private readonly DataGridView gridTargetAll;
        private readonly Label lblSourceTitle;
        private readonly Label lblTargetTitle;
        private readonly Label lblPreviousSource;
        private readonly Label lblCurrentSource;
        private readonly Label lblNextSource;
        private readonly Label lblRecommendedTitle;
        private readonly Label lblAllTargetTitle;
        private readonly Panel sourceContentPanel;
        private readonly Panel targetContentPanel;
        private readonly Panel recommendedTargetPanel;
        private readonly Panel allTargetPanel;
        private readonly TableLayoutPanel compactSourcePanel;

        private readonly Dictionary<string, RequirementTraceMapping> mappingsBySourceId =
            new Dictionary<string, RequirementTraceMapping>(StringComparer.OrdinalIgnoreCase);
        private readonly List<RequirementItem> currentSourceViewItems = new List<RequirementItem>();

        private RequirementTrackingDocumentSnapshot sourceSnapshot;
        private RequirementTrackingDocumentSnapshot targetSnapshot;
        private string sourceDocumentFullName;
        private RequirementItem selectedSource;
        private string selectedTargetId;
        private bool suppressSourceChanged;
        private bool suppressTargetChanged;

        internal RequirementTrackingConsoleControl(Func<Word.Application> applicationAccessor)
        {
            this.applicationAccessor = applicationAccessor ?? throw new ArgumentNullException(nameof(applicationAccessor));

            Dock = DockStyle.Fill;
            BackColor = Color.White;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            AutoScaleMode = AutoScaleMode.Dpi;

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(6)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            FlowLayoutPanel header = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, 4)
            };
            btnImportTarget = CreatePrimaryButton("导入互追文档");
            btnImportTarget.Dock = DockStyle.None;
            btnImportTarget.Width = 168;
            btnImportTarget.Click += (_, __) => ImportTargetDocument();
            btnExportTable = CreatePrimaryButton("导出追踪表");
            btnExportTable.Dock = DockStyle.None;
            btnExportTable.Width = 120;
            btnExportTable.Click += (_, __) => ExportTraceTable();
            btnClearTraceMappings = CreatePrimaryButton("清空追踪关系");
            btnClearTraceMappings.Dock = DockStyle.None;
            btnClearTraceMappings.Width = 156;
            btnClearTraceMappings.Click += (_, __) => ClearTraceMappings();
            header.Controls.Add(btnImportTarget);
            header.Controls.Add(btnExportTable);
            header.Controls.Add(btnClearTraceMappings);
            lblTraceTemplate = CreateToolbarLabel("追踪模板");
            lblTraceTemplate.Margin = new Padding(0, 4, 4, 0);
            cmbTraceTemplate = CreateComboBox("SRS->SDS", "SDS->SRS", "SDS->SDD", "SDD->SDS", "自定义");
            cmbTraceTemplate.Width = 140;
            cmbTraceTemplate.SelectedIndexChanged += (_, __) => OnTraceTemplateChanged();
            FlowLayoutPanel secondHeader = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                WrapContents = false
            };
            secondHeader.Controls.Add(lblTraceTemplate);
            secondHeader.Controls.Add(cmbTraceTemplate);
            lblCustomSourceTitle = CreateToolbarLabel("左表头");
            txtCustomSourceTitle = new TextBox { Width = 130, Visible = false, Margin = new Padding(0, 4, 6, 0) };
            lblCustomTargetTitle = CreateToolbarLabel("右表头");
            txtCustomTargetTitle = new TextBox { Width = 130, Visible = false, Margin = new Padding(0, 4, 0, 0) };
            lblCustomSourceTitle.Visible = false;
            lblCustomTargetTitle.Visible = false;
            secondHeader.Controls.Add(lblCustomSourceTitle);
            secondHeader.Controls.Add(txtCustomSourceTitle);
            secondHeader.Controls.Add(lblCustomTargetTitle);
            secondHeader.Controls.Add(txtCustomTargetTitle);

            TableLayoutPanel split = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                Margin = new Padding(0)
            };
            split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
            split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

            lblSourceTitle = CreatePaneTitle("当前文档需求");
            lblTargetTitle = CreatePaneTitle("互追文档需求");
            lblPreviousSource = CreateSourceCardLabel();
            lblCurrentSource = CreateSourceCardLabel(true);
            lblNextSource = CreateSourceCardLabel();
            lblPreviousSource.Click += (_, __) => SelectAdjacentSource(-1);
            lblNextSource.Click += (_, __) => SelectAdjacentSource(1);

            cmbSourceFilter = CreateComboBox("全部", "已追踪", "未追踪");
            cmbTargetFilter = CreateComboBox("全部", "已追踪", "未追踪");
            cmbSourceFilter.SelectedIndexChanged += (_, __) => RefreshViewPreservingSelection();
            cmbTargetFilter.SelectedIndexChanged += (_, __) => RenderTargets();

            btnPreviousSource = CreateSmallButton("上一条");
            btnNextSource = CreateSmallButton("下一条");
            btnNextUnmappedSource = CreateSmallButton("下一条未追踪");
            btnPreviousSource.Click += (_, __) => SelectAdjacentSource(-1);
            btnNextSource.Click += (_, __) => SelectAdjacentSource(1);
            btnNextUnmappedSource.Click += (_, __) => SelectNextUnmappedSource();

            txtTargetSearch = new TextBox
            {
                Width = 220,
                Margin = new Padding(6, 4, 0, 0)
            };
            txtTargetSearch.TextChanged += (_, __) => RenderTargets();

            gridSource = CreateRequirementGrid(false);
            gridSource.SelectionChanged += GridSource_SelectionChanged;
            gridTargetRecommended = CreateRequirementGrid(true);
            gridTargetAll = CreateRequirementGrid(true);
            WireTargetGrid(gridTargetRecommended);
            WireTargetGrid(gridTargetAll);

            lblRecommendedTitle = CreatePaneTitle("候选推荐（最多 20 条）");
            lblAllTargetTitle = CreatePaneTitle("全部目标需求");
            compactSourcePanel = CreateCompactSourcePanel();
            sourceContentPanel = new Panel { Dock = DockStyle.Fill };
            sourceContentPanel.Controls.Add(gridSource);
            sourceContentPanel.Controls.Add(compactSourcePanel);
            recommendedTargetPanel = CreateTargetSectionPanel(lblRecommendedTitle, gridTargetRecommended);
            allTargetPanel = CreateTargetSectionPanel(lblAllTargetTitle, gridTargetAll);
            TableLayoutPanel targetLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2 };
            targetLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 40f));
            targetLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 60f));
            targetLayout.Controls.Add(recommendedTargetPanel, 0, 0);
            targetLayout.Controls.Add(allTargetPanel, 0, 1);
            targetContentPanel = new Panel { Dock = DockStyle.Fill };
            targetContentPanel.Controls.Add(targetLayout);

            split.Controls.Add(CreateSourcePane(), 0, 0);
            split.Controls.Add(CreateTargetPane(), 1, 0);

            root.Controls.Add(header, 0, 0);
            root.Controls.Add(secondHeader, 0, 1);
            root.Controls.Add(split, 0, 2);
            Controls.Add(root);
        }

        internal void LoadCurrentDocumentAsSource()
        {
            Word.Document document = GetApplication()?.ActiveDocument;
            if (document == null)
            {
                lblSourceTitle.Text = "当前没有活动文档";
                lblTargetTitle.Text = "互追文档需求";
                ClearAllGrids();
                return;
            }

            sourceSnapshot = BuildSavedSnapshot(document);
            sourceDocumentFullName = sourceSnapshot.FullName;
            LoadTraceMappingsFromSourceDocument();
            selectedSource = sourceSnapshot.Requirements.FirstOrDefault();
            RenderSources(selectedSource?.Id);
            RenderTargets();
            UpdateTitlesAndStatus();
        }

        private void ImportTargetDocument()
        {
            Word.Application app = GetApplication();
            Word.Document document = PickTargetDocument(app);
            if (document == null)
            {
                return;
            }

            targetSnapshot = BuildSavedSnapshot(document);
            RenderTargets();
            UpdateTitlesAndStatus();
        }

        private Word.Document PickTargetDocument(Word.Application app)
        {
            if (app == null)
            {
                MessageBox.Show(this, "当前没有可用的 Word 文档。", "需求追踪");
                return null;
            }

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择互追文档";
                dialog.Filter = "Word 文档|*.doc;*.docx;*.docm;*.dot;*.dotx;*.dotm|所有文件|*.*";
                dialog.Multiselect = false;

                string sourceDirectory = Path.GetDirectoryName(sourceDocumentFullName ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(sourceDirectory) && Directory.Exists(sourceDirectory))
                {
                    dialog.InitialDirectory = sourceDirectory;
                }

                DialogResult result = dialog.ShowDialog();
                if (result != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
                {
                    return null;
                }

                string selectedPath = NormalizeFilePath(dialog.FileName);
                foreach (Word.Document openDocument in app.Documents)
                {
                    if (string.Equals(NormalizeFilePath(openDocument?.FullName), selectedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        return openDocument;
                    }
                }

                try
                {
                    return app.Documents.Open(dialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, "无法打开所选互追文档：\r\n" + ex.Message, "需求追踪");
                    return null;
                }
            }
        }

        private void ExportTraceTable()
        {
            if (targetSnapshot == null)
            {
                MessageBox.Show(this, "请先导入待追踪文档。", "需求追踪");
                return;
            }

            List<RequirementTraceExportRow> rows = BuildTraceExportRows();
            if (rows.Count == 0)
            {
                MessageBox.Show(this, "当前模板下没有可导出的互追需求。", "需求追踪");
                return;
            }

            DialogResult result = MessageBox.Show(
                this,
                "确认把需求追踪表导出到当前 Word 光标位置？",
                "导出需求追踪表",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Question);
            if (result != DialogResult.OK)
            {
                return;
            }

            Word.Application app = GetApplication();
            Word.Selection selection = app?.Selection;
            Word.Document document = app?.ActiveDocument;
            if (selection == null || document == null)
            {
                MessageBox.Show(this, "当前没有可插入表格的 Word 光标位置。", "需求追踪");
                return;
            }

            try
            {
                InsertTraceTable(document, selection.Range, rows);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "导出需求追踪表失败：\r\n" + ex.Message, "需求追踪");
            }
        }

        private void ClearTraceMappings()
        {
            if (mappingsBySourceId.Count == 0)
            {
                MessageBox.Show(this, "当前没有可清空的追踪关系。", "需求追踪");
                return;
            }

            DialogResult result = MessageBox.Show(
                this,
                "确认清空当前文档中的全部追踪关系吗？该操作会立即保存。",
                "清空追踪关系",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
            {
                return;
            }

            mappingsBySourceId.Clear();
            selectedTargetId = null;
            SaveTraceMappingsToSourceDocument();
            RenderSources(selectedSource?.Id);
            RenderTargets();
        }

        private List<RequirementTraceExportRow> BuildTraceExportRows()
        {
            Dictionary<string, RequirementItem> targetById = (targetSnapshot?.Requirements ?? new List<RequirementItem>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.Id))
                .GroupBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);

            List<RequirementTraceExportRow> rows = new List<RequirementTraceExportRow>();
            foreach (RequirementItem source in (sourceSnapshot?.Requirements ?? new List<RequirementItem>()).Where(item => item != null))
            {
                List<RequirementItem> targets = new List<RequirementItem>();
                if (!string.IsNullOrWhiteSpace(source.Id) &&
                    mappingsBySourceId.TryGetValue(source.Id, out RequirementTraceMapping mapping))
                {
                    foreach (string targetId in mapping.TargetRequirementIds.Where(item => !string.IsNullOrWhiteSpace(item)))
                    {
                        if (targetById.TryGetValue(targetId, out RequirementItem target))
                        {
                            targets.Add(target);
                        }
                    }
                }

                if (targets.Count == 0)
                {
                    rows.Add(new RequirementTraceExportRow { Source = source, SourceSpan = 1 });
                    continue;
                }

                for (int i = 0; i < targets.Count; i++)
                {
                    rows.Add(new RequirementTraceExportRow
                    {
                        Source = source,
                        Target = targets[i],
                        SourceSpan = i == 0 ? targets.Count : 0
                    });
                }
            }

            return rows;
        }

        private void InsertTraceTable(Word.Document document, Word.Range range, List<RequirementTraceExportRow> rows)
        {
            RequirementTraceTableExporter.InsertTraceTable(
                document,
                range,
                rows,
                GetTemplateSideDisplayName(true),
                GetTemplateSideDisplayName(false));
        }

        private RequirementTrackingDocumentSnapshot BuildSavedSnapshot(Word.Document document)
        {
            string fullName = NormalizeFilePath(document?.FullName);
            RequirementTrackingDocumentSnapshot snapshot = new RequirementTrackingDocumentSnapshot
            {
                FullName = fullName,
                DisplayName = string.IsNullOrWhiteSpace(fullName) ? "当前文档" : Path.GetFileName(fullName),
                Requirements = RequirementExtractionPaneControl.LoadSavedRequirementItems(document) ?? new List<RequirementItem>()
            };

            if (snapshot.Requirements.Count == 0)
            {
                MessageBox.Show(this, $"{snapshot.DisplayName} 中没有已保存的需求提取结果。\r\n请先在该文档中使用“需求提取”并点击“保存”。", "需求追踪");
            }

            return snapshot;
        }

        private Control CreateSourcePane()
        {
            TableLayoutPanel pane = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(8)
            };
            pane.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            pane.RowStyles.Add(new RowStyle(SizeType.Absolute, 82f));
            pane.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            pane.Controls.Add(lblSourceTitle, 0, 0);
            pane.Controls.Add(CreateSourceToolbar(), 0, 1);
            pane.Controls.Add(sourceContentPanel, 0, 2);
            return WrapPane(pane);
        }

        private Control CreateTargetPane()
        {
            TableLayoutPanel pane = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(8)
            };
            pane.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            pane.RowStyles.Add(new RowStyle(SizeType.Absolute, 38f));
            pane.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            pane.Controls.Add(lblTargetTitle, 0, 0);
            pane.Controls.Add(CreateTargetToolbar(), 0, 1);
            pane.Controls.Add(targetContentPanel, 0, 2);
            return WrapPane(pane);
        }

        private Control CreateSourceToolbar()
        {
            TableLayoutPanel panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0)
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 38f));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 38f));

            FlowLayoutPanel filters = CreateToolbarPanel();
            filters.Controls.Add(CreateToolbarLabel("状态"));
            filters.Controls.Add(cmbSourceFilter);

            FlowLayoutPanel nav = CreateToolbarPanel();
            nav.Controls.Add(btnPreviousSource);
            nav.Controls.Add(btnNextSource);
            nav.Controls.Add(btnNextUnmappedSource);

            panel.Controls.Add(filters, 0, 0);
            panel.Controls.Add(nav, 0, 1);
            return panel;
        }

        private Control CreateTargetToolbar()
        {
            FlowLayoutPanel panel = CreateToolbarPanel();
            panel.Controls.Add(CreateToolbarLabel("状态"));
            panel.Controls.Add(cmbTargetFilter);
            panel.Controls.Add(CreateToolbarLabel("名称搜索"));
            panel.Controls.Add(txtTargetSearch);
            return panel;
        }

        private static FlowLayoutPanel CreateToolbarPanel()
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0)
            };
        }

        private static TableLayoutPanel CreateCompactSourcePanel()
        {
            TableLayoutPanel panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(0)
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 72f));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 86f));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 72f));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            return panel;
        }

        private void RenderSources(string preferredId)
        {
            suppressSourceChanged = true;
            gridSource.SuspendLayout();
            try
            {
                currentSourceViewItems.Clear();
                currentSourceViewItems.AddRange(GetFilteredSourceRequirements());

                selectedSource = PickSelectedSource(preferredId);
                RenderCompactSourceCards();
                RenderDetailedSourceGrid();
                ApplyViewModeVisibility();
            }
            finally
            {
                gridSource.ResumeLayout();
                suppressSourceChanged = false;
            }
        }

        private void RenderCompactSourceCards()
        {
            compactSourcePanel.SuspendLayout();
            try
            {
                compactSourcePanel.Controls.Clear();
                int index = GetSelectedSourceIndex();
                RequirementItem previous = index > 0 ? currentSourceViewItems[index - 1] : null;
                RequirementItem current = index >= 0 ? currentSourceViewItems[index] : null;
                RequirementItem next = index >= 0 && index + 1 < currentSourceViewItems.Count ? currentSourceViewItems[index + 1] : null;

                BindSourceCard(lblPreviousSource, "上一条", previous);
                BindSourceCard(lblCurrentSource, "当前需求", current);
                BindSourceCard(lblNextSource, "下一条", next);
                compactSourcePanel.Controls.Add(lblPreviousSource, 0, 0);
                compactSourcePanel.Controls.Add(lblCurrentSource, 0, 1);
                compactSourcePanel.Controls.Add(lblNextSource, 0, 2);
            }
            finally
            {
                compactSourcePanel.ResumeLayout();
            }
        }

        private void RenderDetailedSourceGrid()
        {
            int firstDisplayedRowIndex = -1;
            try
            {
                firstDisplayedRowIndex = gridSource.FirstDisplayedScrollingRowIndex;
            }
            catch
            {
                firstDisplayedRowIndex = -1;
            }

            gridSource.Rows.Clear();
            foreach (RequirementItem item in currentSourceViewItems.Where(item => item != null))
            {
                int row = gridSource.Rows.Add(RequirementItem.GetDisplayRequirementId(item.Id), item.Name ?? string.Empty, item.SectionNumber ?? string.Empty);
                gridSource.Rows[row].Tag = item;
                if (HasMappedTargets(item.Id))
                {
                    gridSource.Rows[row].DefaultCellStyle.ForeColor = Color.FromArgb(22, 120, 70);
                }
            }

            int selected = FindRowByRequirementId(gridSource, selectedSource?.Id);
            if (selected >= 0)
            {
                gridSource.ClearSelection();
                gridSource.Rows[selected].Selected = true;
                gridSource.CurrentCell = gridSource.Rows[selected].Cells[0];
            }

            if (firstDisplayedRowIndex >= 0 && gridSource.Rows.Count > 0)
            {
                try
                {
                    gridSource.FirstDisplayedScrollingRowIndex = Math.Min(firstDisplayedRowIndex, gridSource.Rows.Count - 1);
                }
                catch
                {
                }
            }
        }

        private void RenderTargets()
        {
            suppressTargetChanged = true;
            try
            {
                RenderTargetGrid(gridTargetRecommended, GetRecommendedTargets());
                RenderTargetGrid(gridTargetAll, GetVisibleTargets());
            }
            finally
            {
                suppressTargetChanged = false;
            }

            ApplyViewModeVisibility();
            UpdateTitlesAndStatus();
        }

        private void RenderTargetGrid(DataGridView grid, IEnumerable<RequirementItem> targets)
        {
            RequirementItem source = GetSelectedSourceRequirement();
            int firstDisplayedRowIndex = -1;
            try
            {
                firstDisplayedRowIndex = grid.FirstDisplayedScrollingRowIndex;
            }
            catch
            {
                firstDisplayedRowIndex = -1;
            }

            grid.SuspendLayout();
            try
            {
                grid.Rows.Clear();
                foreach (RequirementItem target in targets.Where(item => item != null))
                {
                    bool mapped = source != null && IsMappedToCurrentSource(source.Id, target.Id);
                    int row = grid.Rows.Add(mapped, RequirementItem.GetDisplayRequirementId(target.Id), target.Name ?? string.Empty, target.SectionNumber ?? string.Empty);
                    grid.Rows[row].Tag = target;
                    if (IsTargetMappedToAnySource(target.Id))
                    {
                        grid.Rows[row].DefaultCellStyle.ForeColor = Color.FromArgb(22, 120, 70);
                    }
                }

                RestoreTargetSelection(grid);
                if (firstDisplayedRowIndex >= 0 && grid.Rows.Count > 0)
                {
                    try
                    {
                        grid.FirstDisplayedScrollingRowIndex = Math.Min(firstDisplayedRowIndex, grid.Rows.Count - 1);
                    }
                    catch
                    {
                    }
                }
            }
            finally
            {
                grid.ResumeLayout();
            }
        }

        private IEnumerable<RequirementItem> GetFilteredSourceRequirements()
        {
            IEnumerable<RequirementItem> items = sourceSnapshot?.Requirements ?? Enumerable.Empty<RequirementItem>();
            return items.Where(item => item != null)
                .Where(MatchesSourceFilter)
                .ToList();
        }

        private IEnumerable<RequirementItem> GetVisibleTargets()
        {
            IEnumerable<RequirementItem> items = targetSnapshot?.Requirements ?? Enumerable.Empty<RequirementItem>();
            return items.Where(item => item != null)
                .Where(MatchesTargetFilter)
                .ToList();
        }

        private IEnumerable<RequirementItem> GetRecommendedTargets()
        {
            RequirementItem source = GetSelectedSourceRequirement();
            if (source == null)
            {
                return Enumerable.Empty<RequirementItem>();
            }

            HashSet<string> keywords = ExtractKeywords(source.Name);
            if (keywords.Count == 0)
            {
                return Enumerable.Empty<RequirementItem>();
            }

            return GetVisibleTargets()
                .Select(target => new { Target = target, Score = GetNameScore(keywords, target.Name) })
                .Where(item => item.Score > 0)
                .OrderByDescending(item => item.Score)
                .ThenBy(item => item.Target.Id ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .Take(20)
                .Select(item => item.Target)
                .ToList();
        }

        private RequirementItem PickSelectedSource(string preferredId)
        {
            if (currentSourceViewItems.Count == 0)
            {
                return null;
            }

            RequirementItem preferred = null;
            if (!string.IsNullOrWhiteSpace(preferredId))
            {
                preferred = currentSourceViewItems.FirstOrDefault(item => string.Equals(item.Id, preferredId, StringComparison.OrdinalIgnoreCase));
            }

            if (preferred == null && selectedSource != null)
            {
                preferred = currentSourceViewItems.FirstOrDefault(item => string.Equals(item.Id, selectedSource.Id, StringComparison.OrdinalIgnoreCase));
            }

            return preferred ?? currentSourceViewItems[0];
        }

        private void ApplyViewModeVisibility()
        {
            compactSourcePanel.Visible = false;
            gridSource.Visible = true;
            recommendedTargetPanel.Visible = true;
            allTargetPanel.Visible = true;
            btnPreviousSource.Enabled = selectedSource != null;
            btnNextSource.Enabled = selectedSource != null;
            btnNextUnmappedSource.Enabled = selectedSource != null;
        }

        private void RefreshViewPreservingSelection()
        {
            string id = selectedSource?.Id;
            RenderSources(id);
            RenderTargets();
            UpdateTitlesAndStatus();
        }

        private void GridSource_SelectionChanged(object sender, EventArgs e)
        {
            if (suppressSourceChanged || !gridSource.Visible)
            {
                return;
            }

            RequirementItem item = gridSource.CurrentRow?.Tag as RequirementItem;
            if (item == null)
            {
                return;
            }

            if (selectedSource != null && string.Equals(selectedSource.Id, item.Id, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            selectedSource = item;
            RenderCompactSourceCards();
            RenderTargets();
        }

        private void GridTarget_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            DataGridView grid = sender as DataGridView;
            if (grid?.IsCurrentCellDirty == true)
            {
                grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void GridTarget_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (suppressTargetChanged || e.RowIndex < 0 || e.ColumnIndex != 0)
            {
                return;
            }

            DataGridView grid = sender as DataGridView;
            RequirementItem source = GetSelectedSourceRequirement();
            RequirementItem target = grid?.Rows[e.RowIndex].Tag as RequirementItem;
            if (source == null || target == null)
            {
                return;
            }

            selectedTargetId = target.Id;
            bool isMapped = Convert.ToBoolean(grid.Rows[e.RowIndex].Cells[0].Value ?? false);
            SetMappingState(source.Id, target.Id, isMapped);
            SaveTraceMappingsToSourceDocument();
            RenderSources(source.Id);
            RenderTargets();
        }

        private void GridTarget_SelectionChanged(object sender, EventArgs e)
        {
            if (suppressTargetChanged)
            {
                return;
            }

            DataGridView grid = sender as DataGridView;
            RequirementItem target = grid?.CurrentRow?.Tag as RequirementItem;
            if (target == null || string.IsNullOrWhiteSpace(target.Id))
            {
                return;
            }

            selectedTargetId = target.Id;
        }

        private void SelectAdjacentSource(int offset)
        {
            int index = GetSelectedSourceIndex();
            if (index < 0)
            {
                return;
            }

            int nextIndex = Math.Max(0, Math.Min(currentSourceViewItems.Count - 1, index + offset));
            SelectSource(currentSourceViewItems[nextIndex]);
        }

        private void SelectNextUnmappedSource()
        {
            List<RequirementItem> all = currentSourceViewItems
                .Where(item => item != null)
                .ToList();
            if (all.Count == 0)
            {
                return;
            }

            int currentIndex = all.FindIndex(item => string.Equals(item.Id, selectedSource?.Id, StringComparison.OrdinalIgnoreCase));
            int start = currentIndex < 0 ? 0 : currentIndex + 1;
            RequirementItem next = all.Skip(start).Concat(all.Take(start)).FirstOrDefault(item => !HasMappedTargets(item.Id));
            if (next != null)
            {
                SelectSource(next);
            }
        }

        private void SelectSource(RequirementItem item)
        {
            if (item == null)
            {
                return;
            }

            selectedSource = item;
            RenderSources(item.Id);
            RenderTargets();
        }

        private int GetSelectedSourceIndex()
        {
            if (selectedSource == null)
            {
                return -1;
            }

            return currentSourceViewItems.FindIndex(item => string.Equals(item.Id, selectedSource.Id, StringComparison.OrdinalIgnoreCase));
        }

        private RequirementItem GetSelectedSourceRequirement()
        {
            return selectedSource;
        }

        private bool MatchesSourceFilter(RequirementItem item)
        {
            switch (GetFilterMode(cmbSourceFilter))
            {
                case FilterMode.Mapped:
                    return HasMappedTargets(item.Id);
                case FilterMode.Unmapped:
                    return !HasMappedTargets(item.Id);
                default:
                    return true;
            }
        }

        private bool MatchesTargetFilter(RequirementItem item)
        {
            if (!MatchesTargetStatusFilter(item))
            {
                return false;
            }

            string search = (txtTargetSearch.Text ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(search) ||
                   (item.Name ?? string.Empty).IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool MatchesTargetStatusFilter(RequirementItem item)
        {
            switch (GetFilterMode(cmbTargetFilter))
            {
                case FilterMode.Mapped:
                    return IsTargetMappedToAnySource(item.Id);
                case FilterMode.Unmapped:
                    return !IsTargetMappedToAnySource(item.Id);
                default:
                    return true;
            }
        }

        private void SetMappingState(string sourceId, string targetId, bool isMapped)
        {
            if (string.IsNullOrWhiteSpace(sourceId) || string.IsNullOrWhiteSpace(targetId))
            {
                return;
            }

            if (!mappingsBySourceId.TryGetValue(sourceId, out RequirementTraceMapping mapping))
            {
                if (!isMapped)
                {
                    return;
                }

                mapping = new RequirementTraceMapping { SourceRequirementId = sourceId };
                mappingsBySourceId[sourceId] = mapping;
            }

            bool exists = mapping.TargetRequirementIds.Any(item => string.Equals(item, targetId, StringComparison.OrdinalIgnoreCase));
            if (isMapped && !exists)
            {
                mapping.TargetRequirementIds.Add(targetId);
            }
            else if (!isMapped && exists)
            {
                mapping.TargetRequirementIds = mapping.TargetRequirementIds
                    .Where(item => !string.Equals(item, targetId, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            if (mapping.TargetRequirementIds.Count == 0)
            {
                mappingsBySourceId.Remove(sourceId);
            }
        }

        private void LoadTraceMappings(Word.Document doc)
        {
            mappingsBySourceId.Clear();
            Office.CustomXMLPart part = FindSavedTraceMappingPart(doc);
            if (part == null)
            {
                return;
            }

            try
            {
                XDocument document = XDocument.Parse(part.XML);
                XName mappingName = XName.Get("mapping", SavedTraceMappingsNamespace);
                XName targetName = XName.Get("target", SavedTraceMappingsNamespace);
                string currentTemplate = GetCurrentTraceTemplate().ToString();
                bool hasTemplateTags = document.Descendants(mappingName).Any(element => element.Attribute("template") != null);
                foreach (XElement mappingElement in document.Descendants(mappingName))
                {
                    string template = (string)mappingElement.Attribute("template") ?? string.Empty;
                    if (hasTemplateTags && !string.Equals(template, currentTemplate, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string sourceId = (string)mappingElement.Attribute("sourceId") ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(sourceId))
                    {
                        continue;
                    }

                    RequirementTraceMapping mapping = new RequirementTraceMapping { SourceRequirementId = sourceId };
                    foreach (XElement targetElement in mappingElement.Elements(targetName))
                    {
                        string targetId = (string)targetElement.Attribute("id") ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(targetId))
                        {
                            mapping.TargetRequirementIds.Add(targetId);
                        }
                    }

                    if (mapping.TargetRequirementIds.Count > 0)
                    {
                        mappingsBySourceId[sourceId] = mapping;
                    }
                }
            }
            catch
            {
                mappingsBySourceId.Clear();
            }
        }

        private void LoadTraceMappingsFromSourceDocument()
        {
            Word.Document doc = GetSourceDocument();
            if (doc == null)
            {
                mappingsBySourceId.Clear();
                return;
            }

            try
            {
                LoadTraceMappings(doc);
            }
            finally
            {
                ReleaseComObject(doc);
            }
        }

        private void SaveTraceMappingsToSourceDocument()
        {
            Word.Document doc = GetSourceDocument();
            if (doc == null)
            {
                return;
            }

            try
            {
                Office.CustomXMLPart part = FindSavedTraceMappingPart(doc);
                string existingXml = part?.XML ?? string.Empty;
                string updatedXml = BuildTraceMappingsXml(existingXml);
                if (part != null)
                {
                    part.Delete();
                }

                doc.CustomXMLParts.Add(updatedXml);
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(doc);
            }
        }

        private string BuildTraceMappingsXml(string existingXml)
        {
            XElement root = null;
            try
            {
                if (!string.IsNullOrWhiteSpace(existingXml))
                {
                    root = XDocument.Parse(existingXml).Root;
                }
            }
            catch
            {
                root = null;
            }

            if (root == null)
            {
                root = new XElement(XName.Get("traceMappings", SavedTraceMappingsNamespace));
            }

            string currentTemplate = GetCurrentTraceTemplate().ToString();
            XName mappingName = XName.Get("mapping", SavedTraceMappingsNamespace);
            bool hasTemplateTags = root.Elements(mappingName).Any(element => element.Attribute("template") != null);
            if (hasTemplateTags)
            {
                root.Elements(mappingName)
                    .Where(element => string.Equals((string)element.Attribute("template") ?? string.Empty, currentTemplate, StringComparison.OrdinalIgnoreCase))
                    .Remove();
            }
            else
            {
                root.Elements(mappingName).Remove();
            }

            foreach (RequirementTraceMapping mapping in mappingsBySourceId.Values
                .Where(mapping => mapping != null &&
                                  !string.IsNullOrWhiteSpace(mapping.SourceRequirementId) &&
                                  mapping.TargetRequirementIds.Any(id => !string.IsNullOrWhiteSpace(id))))
            {
                root.Add(new XElement(mappingName,
                    new XAttribute("template", currentTemplate),
                    new XAttribute("sourceId", mapping.SourceRequirementId ?? string.Empty),
                    mapping.TargetRequirementIds
                        .Where(id => !string.IsNullOrWhiteSpace(id))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .Select(id =>
                            new XElement(XName.Get("target", SavedTraceMappingsNamespace),
                                new XAttribute("id", id)))));
            }

            return root.ToString(SaveOptions.DisableFormatting);
        }

        private static void DeleteSavedTraceMappingParts(Word.Document doc)
        {
            Office.CustomXMLPart part;
            while ((part = FindSavedTraceMappingPart(doc)) != null)
            {
                part.Delete();
            }
        }

        private static Office.CustomXMLPart FindSavedTraceMappingPart(Word.Document doc)
        {
            if (doc == null)
            {
                return null;
            }

            try
            {
                Office.CustomXMLParts parts = doc.CustomXMLParts.SelectByNamespace(SavedTraceMappingsNamespace);
                return parts != null && parts.Count > 0 ? parts[1] : null;
            }
            catch
            {
                return null;
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null || !System.Runtime.InteropServices.Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comObject);
            }
            catch
            {
            }
        }

        private Word.Document GetSourceDocument()
        {
            if (string.IsNullOrWhiteSpace(sourceDocumentFullName))
            {
                return GetApplication()?.ActiveDocument;
            }

            Word.Application app = GetApplication();
            if (app == null)
            {
                return null;
            }

            Word.Documents documents = null;
            Word.Document matchedDocument = null;
            try
            {
                documents = app.Documents;
                for (int i = 1; i <= documents.Count; i++)
                {
                    Word.Document doc = null;
                    try
                    {
                        doc = documents[i];
                        if (string.Equals(NormalizeFilePath(doc?.FullName), sourceDocumentFullName, StringComparison.OrdinalIgnoreCase))
                        {
                            matchedDocument = doc;
                            doc = null;
                            break;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(doc);
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseComObject(documents);
            }

            if (matchedDocument != null)
            {
                return matchedDocument;
            }

            return GetApplication()?.ActiveDocument;
        }

        private bool IsMappedToCurrentSource(string sourceId, string targetId)
        {
            return !string.IsNullOrWhiteSpace(sourceId) &&
                   !string.IsNullOrWhiteSpace(targetId) &&
                   mappingsBySourceId.TryGetValue(sourceId, out RequirementTraceMapping mapping) &&
                   mapping.TargetRequirementIds.Any(item => string.Equals(item, targetId, StringComparison.OrdinalIgnoreCase));
        }

        private bool HasMappedTargets(string sourceId)
        {
            return !string.IsNullOrWhiteSpace(sourceId) &&
                   mappingsBySourceId.TryGetValue(sourceId, out RequirementTraceMapping mapping) &&
                   mapping.TargetRequirementIds.Any(item => !string.IsNullOrWhiteSpace(item));
        }

        private bool IsTargetMappedToAnySource(string targetId)
        {
            return !string.IsNullOrWhiteSpace(targetId) &&
                   mappingsBySourceId.Values.Any(mapping => mapping.TargetRequirementIds.Any(item => string.Equals(item, targetId, StringComparison.OrdinalIgnoreCase)));
        }

        private void UpdateTitlesAndStatus()
        {
            int sourceCount = currentSourceViewItems.Count;
            List<RequirementItem> visibleTargets = GetVisibleTargets().ToList();
            int targetCount = visibleTargets.Count;
            int selectedIndex = selectedSource == null
                ? 0
                : currentSourceViewItems.FindIndex(item => string.Equals(item.Id, selectedSource.Id, StringComparison.OrdinalIgnoreCase)) + 1;
            string selectedText = selectedIndex > 0 ? $"，当前 {selectedIndex}/{sourceCount}" : string.Empty;
            string sourceTitle = GetTemplateSideDisplayName(true);
            string targetTitle = GetTemplateSideDisplayName(false);

            lblSourceTitle.Text = string.IsNullOrWhiteSpace(sourceSnapshot?.DisplayName)
                ? $"当前文档需求：{sourceTitle}（{sourceCount}，已映射 {mappingsBySourceId.Count}{selectedText}）"
                : $"当前文档需求：{sourceSnapshot.DisplayName} - {sourceTitle}（{sourceCount}，已映射 {mappingsBySourceId.Count}{selectedText}）";
            lblTargetTitle.Text = string.IsNullOrWhiteSpace(targetSnapshot?.DisplayName)
                ? $"互追文档需求：{targetTitle}（{targetCount}）"
                : $"互追文档需求：{targetSnapshot.DisplayName} - {targetTitle}（{targetCount}）";
            lblRecommendedTitle.Text = $"候选推荐（当前 {gridTargetRecommended.Rows.Count} 条）";
            lblAllTargetTitle.Text = $"详细需求列表（{gridTargetAll.Rows.Count} 条）";
        }

        private Word.Application GetApplication()
        {
            return applicationAccessor?.Invoke();
        }

        private void ClearAllGrids()
        {
            gridSource.Rows.Clear();
            gridTargetRecommended.Rows.Clear();
            gridTargetAll.Rows.Clear();
            sourceSnapshot = null;
            targetSnapshot = null;
            sourceDocumentFullName = null;
            mappingsBySourceId.Clear();
            selectedSource = null;
            selectedTargetId = null;
            currentSourceViewItems.Clear();
            RenderCompactSourceCards();
            ApplyViewModeVisibility();
        }

        private static FilterMode GetFilterMode(ComboBox comboBox)
        {
            if (comboBox == null)
            {
                return FilterMode.All;
            }

            switch (comboBox.SelectedIndex)
            {
                case 1:
                    return FilterMode.Mapped;
                case 2:
                    return FilterMode.Unmapped;
                default:
                    return FilterMode.All;
            }
        }

        private static HashSet<string> ExtractKeywords(string text)
        {
            HashSet<string> keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string normalized = NormalizeKeywordText(text);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return keywords;
            }

            foreach (string part in normalized.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (part.Length >= 2)
                {
                    keywords.Add(part);
                }

                if (part.Length > 3)
                {
                    for (int i = 0; i + 2 <= part.Length && keywords.Count < 40; i++)
                    {
                        keywords.Add(part.Substring(i, 2));
                    }
                }
            }

            return keywords;
        }

        private static int GetNameScore(HashSet<string> keywords, string targetName)
        {
            string target = NormalizeKeywordText(targetName);
            if (string.IsNullOrWhiteSpace(target))
            {
                return 0;
            }

            int score = 0;
            foreach (string keyword in keywords)
            {
                if (target.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    score += keyword.Length > 2 ? 2 : 1;
                }
            }

            return score;
        }

        private static string NormalizeKeywordText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder(text.Length);
            foreach (char c in text.Trim())
            {
                builder.Append(char.IsLetterOrDigit(c) ? c : ' ');
            }

            return builder.ToString();
        }

        private static string GetDocumentTypeDisplayName(RequirementTrackingDocumentSnapshot snapshot)
        {
            string name = ((snapshot?.DisplayName ?? string.Empty) + " " + (snapshot?.FullName ?? string.Empty)).Trim();
            if (ContainsAny(name, "软件需求规格说明", "需求规格说明", "SRS"))
            {
                return "软件需求规格说明";
            }

            if (ContainsAny(name, "软件设计说明", "软件设计描述", "SDD", "SDS"))
            {
                return "软件设计说明";
            }

            if (ContainsAny(name, "软件研制任务书", "研制任务书"))
            {
                return "软件研制任务书";
            }

            if (ContainsAny(name, "软件测试说明", "软件测试描述", "测试说明", "STD", "STS"))
            {
                return "软件测试说明";
            }

            if (ContainsAny(name, "系统规格说明", "系统/子系统规格说明", "系统子系统规格说明", "SSS"))
            {
                return "系统规格说明";
            }

            return string.IsNullOrWhiteSpace(snapshot?.DisplayName)
                ? "追踪文档"
                : Path.GetFileNameWithoutExtension(snapshot.DisplayName);
        }

        private static bool ContainsAny(string text, params string[] tokens)
        {
            if (string.IsNullOrWhiteSpace(text) || tokens == null)
            {
                return false;
            }

            return tokens.Any(token => !string.IsNullOrWhiteSpace(token) &&
                                       text.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static string NormalizeFilePath(string path)
        {
            try
            {
                return string.IsNullOrWhiteSpace(path) ? string.Empty : Path.GetFullPath(path.Trim());
            }
            catch
            {
                return string.Empty;
            }
        }

        private static int FindRowByRequirementId(DataGridView grid, string id)
        {
            if (grid == null || string.IsNullOrWhiteSpace(id))
            {
                return -1;
            }

            for (int i = 0; i < grid.Rows.Count; i++)
            {
                RequirementItem item = grid.Rows[i].Tag as RequirementItem;
                if (string.Equals(item?.Id, id, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static void BindSourceCard(Label label, string caption, RequirementItem item)
        {
            label.Tag = item;
            label.Text = item == null
                ? $"{caption}：\r\n无"
                : $"{caption}：\r\n{item.DisplayText}";
            label.Enabled = item != null;
        }

        private static Label CreatePaneTitle(string text)
        {
            return new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
                Text = text,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true
            };
        }

        private static Label CreateToolbarLabel(string text)
        {
            return new Label
            {
                AutoSize = true,
                Height = 28,
                Text = text,
                TextAlign = ContentAlignment.MiddleRight,
                Margin = new Padding(0, 4, 4, 0),
                Padding = new Padding(0, 5, 0, 0)
            };
        }

        private static Label CreateSourceCardLabel(bool current = false)
        {
            return new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = false,
                BorderStyle = BorderStyle.FixedSingle,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(12, 9, 12, 9),
                Margin = new Padding(0, 0, 0, 8),
                Cursor = Cursors.Hand,
                BackColor = current ? Color.FromArgb(225, 239, 255) : Color.White,
                ForeColor = current ? Color.FromArgb(24, 82, 168) : Color.FromArgb(43, 57, 76),
                Font = new Font("Microsoft YaHei UI", current ? 10F : 9F, current ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Point)
            };
        }

        private static Button CreatePrimaryButton(string text)
        {
            Button button = new Button
            {
                Dock = DockStyle.Fill,
                Text = text,
                Height = 36,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point),
                Margin = new Padding(4, 6, 0, 6),
                UseVisualStyleBackColor = false,
                BackColor = Color.FromArgb(42, 122, 226),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(58, 136, 236);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(33, 105, 199);
            return button;
        }

        private static Button CreateSmallButton(string text)
        {
            Button button = new Button
            {
                AutoSize = false,
                Width = text.Length > 4 ? 136 : 78,
                Height = 30,
                Text = text,
                Margin = new Padding(6, 4, 0, 0),
                UseVisualStyleBackColor = false,
                BackColor = Color.FromArgb(235, 243, 255),
                ForeColor = Color.FromArgb(36, 89, 171),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(176, 203, 240);
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(223, 236, 255);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(209, 227, 253);
            return button;
        }

        private static ComboBox CreateComboBox(params string[] items)
        {
            ComboBox comboBox = new ComboBox
            {
                Width = 88,
                Height = 28,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Margin = new Padding(0, 4, 6, 0)
            };
            comboBox.Items.AddRange(items.Cast<object>().ToArray());
            comboBox.SelectedIndex = 0;
            return comboBox;
        }

        private static Control WrapPane(Control content)
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 0, 8, 0)
            };
            panel.Controls.Add(content);
            return panel;
        }

        private static DataGridView CreateRequirementGrid(bool includeCheckColumn)
        {
            DataGridView grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                ColumnHeadersHeight = 32,
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(229, 234, 242),
                MultiSelect = false,
                ReadOnly = !includeCheckColumn,
                RowHeadersVisible = false,
                RowTemplate = { Height = 30 },
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 247, 250);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(43, 57, 76);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            grid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 239, 255);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            EnableDoubleBuffering(grid);

            if (includeCheckColumn)
            {
                grid.Columns.Add(new DataGridViewCheckBoxColumn
                {
                    Name = "Mapped",
                    HeaderText = "追踪",
                    Width = 54,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    ReadOnly = false
                });
            }

            grid.Columns.Add(CreateRequirementTextColumn("Id", "需求标识", 30f));
            grid.Columns.Add(CreateRequirementTextColumn("Name", "需求名称", 50f));
            grid.Columns.Add(CreateRequirementTextColumn("SectionNumber", "章节号", 20f));
            return grid;
        }

        private void RestoreTargetSelection(DataGridView grid)
        {
            if (grid == null || grid.Rows.Count == 0 || string.IsNullOrWhiteSpace(selectedTargetId))
            {
                return;
            }

            int rowIndex = FindRowByRequirementId(grid, selectedTargetId);
            if (rowIndex < 0)
            {
                return;
            }

            try
            {
                grid.ClearSelection();
                grid.Rows[rowIndex].Selected = true;
                if (grid.Rows[rowIndex].Cells.Count > 0)
                {
                    grid.CurrentCell = grid.Rows[rowIndex].Cells[0];
                }
            }
            catch
            {
            }
        }

        private static DataGridViewTextBoxColumn CreateRequirementTextColumn(string name, string headerText, float fillWeight)
        {
            return new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = headerText,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                FillWeight = fillWeight,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = DataGridViewContentAlignment.MiddleCenter },
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
        }

        private static Panel CreateTargetSectionPanel(Control title, Control content)
        {
            Panel panel = new Panel
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 26f));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.Controls.Add(title, 0, 0);
            layout.Controls.Add(content, 0, 1);
            panel.Controls.Add(layout);
            return panel;
        }

        private RequirementTraceTemplate GetCurrentTraceTemplate()
        {
            if (cmbTraceTemplate == null)
            {
                return RequirementTraceTemplate.SrsToSds;
            }

            switch (cmbTraceTemplate.SelectedIndex)
            {
                case 1:
                    return RequirementTraceTemplate.SdsToSrs;
                case 2:
                    return RequirementTraceTemplate.SdsToSdd;
                case 3:
                    return RequirementTraceTemplate.SddToSds;
                case 4:
                    return RequirementTraceTemplate.Custom;
                default:
                    return RequirementTraceTemplate.SrsToSds;
            }
        }

        private void OnTraceTemplateChanged()
        {
            bool custom = GetCurrentTraceTemplate() == RequirementTraceTemplate.Custom;
            lblCustomSourceTitle.Visible = custom;
            txtCustomSourceTitle.Visible = custom;
            lblCustomTargetTitle.Visible = custom;
            txtCustomTargetTitle.Visible = custom;
            LoadTraceMappingsFromSourceDocument();
            RefreshViewPreservingSelection();
        }

        private bool MatchesSourceTemplate(RequirementItem item)
        {
            return MatchesTemplateRequirement(item?.Id, GetSourceTemplatePrefix());
        }

        private bool MatchesTargetTemplate(RequirementItem item)
        {
            return MatchesTemplateRequirement(item?.Id, GetTargetTemplatePrefix());
        }

        private static bool MatchesTemplateRequirement(string id, string prefix)
        {
            return RequirementItem.ContainsRequirementPrefix(id, prefix);
        }

        private string GetSourceTemplatePrefix()
        {
            switch (GetCurrentTraceTemplate())
            {
                case RequirementTraceTemplate.SdsToSrs:
                case RequirementTraceTemplate.SdsToSdd:
                    return "SDS";
                case RequirementTraceTemplate.SddToSds:
                    return "SDD";
                default:
                    return "SRS";
            }
        }

        private string GetTargetTemplatePrefix()
        {
            switch (GetCurrentTraceTemplate())
            {
                case RequirementTraceTemplate.SrsToSds:
                case RequirementTraceTemplate.SddToSds:
                    return "SDS";
                case RequirementTraceTemplate.SdsToSrs:
                    return "SRS";
                default:
                    return "SDD";
            }
        }

        private string GetTemplateSideDisplayName(bool sourceSide)
        {
            switch (GetCurrentTraceTemplate())
            {
                case RequirementTraceTemplate.Custom:
                    return sourceSide
                        ? (txtCustomSourceTitle.Text ?? string.Empty).Trim()
                        : (txtCustomTargetTitle.Text ?? string.Empty).Trim();
                case RequirementTraceTemplate.SrsToSds:
                    return sourceSide ? "软件需求规格说明" : "软件概要设计说明";
                case RequirementTraceTemplate.SdsToSrs:
                    return sourceSide ? "软件概要设计说明" : "软件需求规格说明";
                case RequirementTraceTemplate.SdsToSdd:
                    return sourceSide ? "软件概要设计说明" : "软件详细设计说明";
                case RequirementTraceTemplate.SddToSds:
                    return sourceSide ? "软件详细设计说明" : "软件概要设计说明";
                default:
                    return sourceSide ? "当前文档" : "互追文档";
            }
        }

        private void WireTargetGrid(DataGridView grid)
        {
            grid.CurrentCellDirtyStateChanged += GridTarget_CurrentCellDirtyStateChanged;
            grid.CellValueChanged += GridTarget_CellValueChanged;
            grid.SelectionChanged += GridTarget_SelectionChanged;
        }

        private static void EnableDoubleBuffering(DataGridView grid)
        {
            try
            {
                typeof(DataGridView)
                    .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    ?.SetValue(grid, true, null);
            }
            catch
            {
            }
        }

    }
}
