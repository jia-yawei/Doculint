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
        private readonly Button btnAddSelection;
        private readonly Button btnBatchExtraction;
        private readonly Button btnClearCurrent;
        private readonly Button btnSave;
        private readonly Button btnExportTraceTable;
        private readonly Label lblStatus;
        private Word.Document currentDocument;
        private readonly List<RequirementItem> requirements = new List<RequirementItem>();
        private bool batchExtractionEnabled;

        internal event Action<int> RequirementActivated;
        internal event Action<bool> BatchExtractionModeChanged;

        internal RequirementExtractionPaneControl(Func<Word.Application> applicationAccessor)
        {
            this.applicationAccessor = applicationAccessor;
            Dock = DockStyle.Fill;
            BackColor = Color.White;

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            FlowLayoutPanel toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                WrapContents = false
            };

            btnRefresh = CreateButton("提取标识");
            btnRefresh.Click += (_, __) => LoadCurrentDocumentRequirements();
            btnAddSelection = CreateButton("添加");
            btnAddSelection.Click += (_, __) => AddSelectionAsRequirement();
            btnBatchExtraction = CreateButton("批量提取");
            btnBatchExtraction.Click += (_, __) => ToggleBatchExtractionMode();
            btnClearCurrent = CreateButton("清除");
            btnClearCurrent.Click += (_, __) => ClearCurrentRequirementName();
            btnSave = CreateButton("保存");
            btnSave.Click += (_, __) => SaveToCurrentDocument();
            btnExportTraceTable = CreateButton("导出追踪表");
            btnExportTraceTable.Click += (_, __) => ExportTraceTable();
            toolbar.Controls.Add(btnRefresh);
            toolbar.Controls.Add(btnAddSelection);
            toolbar.Controls.Add(btnBatchExtraction);
            toolbar.Controls.Add(btnClearCurrent);
            toolbar.Controls.Add(btnSave);
            toolbar.Controls.Add(btnExportTraceTable);

            grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = true,
                AllowUserToResizeColumns = false,
                AllowUserToResizeRows = false,
                AutoGenerateColumns = false,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
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
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 247, 250);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(32, 45, 64);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 9F);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 238, 255);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            EnableDoubleBuffering(grid);
            grid.Columns.Add(CreateTextColumn("Id", "需求标识", true));
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
                btnRefresh.Text = "重新提取";
                ApplyBatchExtractionState(batchExtractionEnabled);
                SetStatus($"已提取 {requirements.Count} 个需求标识，请选中需求名称后点击“添加”");
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
            TryAddSelectionAsRequirement(false);
        }

        internal bool TryAutoAddSelection()
        {
            return TryAddSelectionAsRequirement(true);
        }

        internal bool BatchExtractionEnabled => batchExtractionEnabled;

        internal void SetBatchExtractionMode(bool enabled)
        {
            if (batchExtractionEnabled == enabled)
            {
                return;
            }

            batchExtractionEnabled = enabled;
            ApplyBatchExtractionState(enabled);
            BatchExtractionModeChanged?.Invoke(enabled);
        }

        private void ToggleBatchExtractionMode()
        {
            SetBatchExtractionMode(!batchExtractionEnabled);
            SetStatus(batchExtractionEnabled
                ? "批量提取已开启：鼠标选中文本后会自动填入当前行"
                : "批量提取已关闭：恢复手动添加");
        }

        private void ApplyBatchExtractionState(bool enabled)
        {
            btnAddSelection.Enabled = !enabled;
            btnBatchExtraction.Text = enabled ? "批量提取：开" : "批量提取";
            btnBatchExtraction.BackColor = enabled ? Color.FromArgb(35, 138, 79) : Color.FromArgb(43, 108, 176);
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
            RenderRows();
            SelectRow(rowIndex);
            SetStatus($"已清除 {item.Id} 的需求名称和章节号");
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
                Padding = new Padding(10, 4, 10, 4)
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private static string GetCurrentDocumentDisplayName(Word.Document doc)
        {
            try
            {
                string name = string.IsNullOrWhiteSpace(doc?.Name) ? "当前文档" : doc.Name;
                return Path.GetFileNameWithoutExtension(name);
            }
            catch
            {
                return "当前文档";
            }
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

        private static void EnableDoubleBuffering(DataGridView targetGrid)
        {
            try
            {
                typeof(DataGridView)
                    .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    ?.SetValue(targetGrid, true, null);
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
