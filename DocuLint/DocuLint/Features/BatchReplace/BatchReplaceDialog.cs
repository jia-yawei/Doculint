using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    public enum BatchFindType
    {
        PlainText,
        WordWildcards
    }

    public enum BatchFindScope
    {
        None,
        Body,
        HeaderFooter,
        Special,
        All
    }

    public class BatchReplaceRule
    {
        public bool Enabled { get; set; }
        public string FindText { get; set; }
        public string ReplaceText { get; set; }
        public BatchFindType FindType { get; set; }
        public BatchFindScope Scope { get; set; }
        public bool ApplyToFileName { get; set; }
        public bool HighlightOnly { get; set; }
        public bool MatchCase { get; set; }
        public bool MatchWholeWord { get; set; }
        public bool MatchSoundsLike { get; set; }
        public bool MatchAllWordForms { get; set; }
    }

    public class BatchReplaceExecutionRequest
    {
        public List<string> FilePaths { get; set; }
        public List<BatchReplaceRule> Rules { get; set; }
        public bool MatchCase { get; set; }
        public bool MatchWholeWord { get; set; }
        public bool FindOnly { get; set; }
    }

    public class BatchReplaceDialog : Form
    {
        private readonly DataGridView dgvRules;
        private readonly TreeView tvFiles;

        public BatchReplaceExecutionRequest Request { get; private set; }

        public BatchReplaceDialog()
        {
            Text = "批量查找与替换";
            Width = 1360;
            Height = 900;
            MinimumSize = new Size(1220, 820);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Regular, GraphicsUnit.Point);

            TabControl tabs = new TabControl { Dock = DockStyle.Fill };
            TabPage pageConfig = new TabPage("查找和替换");
            tabs.TabPages.Add(pageConfig);
            Controls.Add(tabs);

            SplitContainer split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical
            };
            pageConfig.Controls.Add(split);
            Shown += (s, e) => BeginInvoke((Action)(() => ApplySafeSplitLayout(split, 860, 700, 450)));
            SizeChanged += (s, e) => ApplySafeSplitLayout(split, 860, 700, 450);

            TableLayoutPanel leftLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            split.Panel1.Controls.Add(leftLayout);

            FlowLayoutPanel toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                WrapContents = false,
                Padding = new Padding(6, 10, 6, 6)
            };
            leftLayout.Controls.Add(toolbar, 0, 0);

            Button btnAddRule = new Button { Text = "添加行", Width = 110, Height = 34 };
            Button btnDeleteRule = new Button { Text = "删除行", Width = 110, Height = 34 };
            toolbar.Controls.Add(btnAddRule);
            toolbar.Controls.Add(btnDeleteRule);

            dgvRules = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoGenerateColumns = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                ColumnHeadersHeight = 38,
                RowTemplate = { Height = 34 }
            };
            leftLayout.Controls.Add(dgvRules, 0, 1);

            DataGridViewTextBoxColumn colRowNo = new DataGridViewTextBoxColumn
            {
                Name = "colRowNo",
                HeaderText = "列表",
                FillWeight = 40,
                ReadOnly = true,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    BackColor = Color.FromArgb(236, 246, 255),
                    SelectionBackColor = Color.FromArgb(210, 232, 252),
                    SelectionForeColor = Color.Black
                }
            };
            DataGridViewTextBoxColumn colFindText = new DataGridViewTextBoxColumn
            {
                Name = "colFindText",
                HeaderText = "查找",
                FillWeight = 150
            };
            DataGridViewTextBoxColumn colReplaceText = new DataGridViewTextBoxColumn
            {
                Name = "colReplaceText",
                HeaderText = "替换为",
                FillWeight = 150
            };
            dgvRules.Columns.AddRange(
                colRowNo,
                colFindText,
                colReplaceText);

            TableLayoutPanel rightLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(8)
            };
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            split.Panel2.Controls.Add(rightLayout);

            TableLayoutPanel patternLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5
            };
            patternLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 96));
            patternLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            patternLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            patternLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            patternLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            rightLayout.Controls.Add(patternLayout, 0, 0);

            Label lblPattern = new Label { Text = "文件来源:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            patternLayout.Controls.Add(lblPattern, 0, 0);

            Label lblSourceHint = new Label
            {
                Dock = DockStyle.Fill,
                Text = "可一键选活动组，或浏览文件夹/多文件",
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(3, 10, 3, 8)
            };
            patternLayout.Controls.Add(lblSourceHint, 1, 0);

            Button btnAddActiveGroup = new Button { Text = "活动组", Dock = DockStyle.Fill, Margin = new Padding(3, 8, 3, 8) };
            Button btnAddFiles = new Button { Text = "多文件", Dock = DockStyle.Fill, Margin = new Padding(3, 8, 3, 8) };
            Button btnAddFolder = new Button { Text = "文件夹", Dock = DockStyle.Fill, Margin = new Padding(3, 8, 3, 8) };
            Button btnRemoveNode = new Button { Text = "移除", Width = 88, Height = 30, Margin = new Padding(0, 4, 0, 4) };
            patternLayout.Controls.Add(btnAddActiveGroup, 2, 0);
            patternLayout.Controls.Add(btnAddFiles, 3, 0);
            patternLayout.Controls.Add(btnAddFolder, 4, 0);

            TableLayoutPanel fileHeaderLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2
            };
            fileHeaderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            fileHeaderLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            Label lblTargetFiles = new Label
            {
                Text = "要替换的文件:",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true,
                Margin = new Padding(3, 8, 3, 3)
            };
            fileHeaderLayout.Controls.Add(lblTargetFiles, 0, 0);
            fileHeaderLayout.Controls.Add(btnRemoveNode, 1, 0);
            rightLayout.Controls.Add(fileHeaderLayout, 0, 1);

            tvFiles = new TreeView
            {
                CheckBoxes = true,
                Dock = DockStyle.Fill
            };
            rightLayout.Controls.Add(tvFiles, 0, 2);

            FlowLayoutPanel actionLayout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 8, 0, 0)
            };
            rightLayout.Controls.Add(actionLayout, 0, 3);

            Button btnReplace = new Button { Text = "开始替换", Width = 92, Height = 34 };
            actionLayout.Controls.Add(btnReplace);

            btnAddRule.Click += (s, e) => AddDefaultRule();
            btnDeleteRule.Click += (s, e) => DeleteSelectedRules();
            btnAddActiveGroup.Click += (s, e) => AddActiveGroupFiles();
            btnAddFiles.Click += (s, e) => AddFileNodes();
            btnAddFolder.Click += (s, e) => AddFolderNode();
            btnRemoveNode.Click += (s, e) => RemoveSelectedNode();
            tvFiles.NodeMouseClick += TvFiles_NodeMouseClick;

            btnReplace.Click += (s, e) => Submit();
        }

        private void ApplySafeSplitLayout(SplitContainer split, int preferredDistance, int desiredPanel1Min, int desiredPanel2Min)
        {
            if (split.IsDisposed)
            {
                return;
            }

            int availableWidth = split.ClientSize.Width - split.SplitterWidth;
            if (availableWidth <= 0)
            {
                return;
            }

            int panel1Min = desiredPanel1Min;
            int panel2Min = desiredPanel2Min;

            if (panel1Min + panel2Min > availableWidth)
            {
                double ratio = (double)desiredPanel1Min / (desiredPanel1Min + desiredPanel2Min);
                panel1Min = Math.Max(220, (int)Math.Round(availableWidth * ratio));
                panel2Min = Math.Max(220, availableWidth - panel1Min);

                if (panel1Min + panel2Min > availableWidth)
                {
                    panel2Min = Math.Max(120, availableWidth - panel1Min);
                    panel1Min = Math.Max(120, availableWidth - panel2Min);
                }
            }

            split.Panel1MinSize = 0;
            split.Panel2MinSize = 0;

            int initialDistance = preferredDistance;
            if (initialDistance < 0)
            {
                initialDistance = 0;
            }
            if (initialDistance > availableWidth)
            {
                initialDistance = availableWidth;
            }

            if (split.SplitterDistance != initialDistance)
            {
                split.SplitterDistance = initialDistance;
            }

            split.Panel1MinSize = panel1Min;
            split.Panel2MinSize = panel2Min;

            int minDistance = split.Panel1MinSize;
            int maxDistance = availableWidth - split.Panel2MinSize;
            if (maxDistance < minDistance)
            {
                return;
            }

            int finalDistance = preferredDistance;
            if (finalDistance < minDistance)
            {
                finalDistance = minDistance;
            }
            if (finalDistance > maxDistance)
            {
                finalDistance = maxDistance;
            }

            if (split.SplitterDistance != finalDistance)
            {
                split.SplitterDistance = finalDistance;
            }
        }

        private void AddDefaultRule()
        {
            int row = dgvRules.Rows.Add();
            dgvRules.Rows[row].Cells["colRowNo"].Value = row + 1;
            dgvRules.Rows[row].Cells["colFindText"].Value = "";
            dgvRules.Rows[row].Cells["colReplaceText"].Value = "";
            RefreshRuleRowNumbers();
        }

        private void DeleteSelectedRules()
        {
            foreach (DataGridViewRow row in dgvRules.SelectedRows)
            {
                if (!row.IsNewRow)
                {
                    dgvRules.Rows.Remove(row);
                }
            }
            RefreshRuleRowNumbers();
        }

        private void RefreshRuleRowNumbers()
        {
            for (int i = 0; i < dgvRules.Rows.Count; i++)
            {
                dgvRules.Rows[i].Cells["colRowNo"].Value = i + 1;
            }
        }

        private void AddFolderNode()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "请选择文件夹";
                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                string folder = dialog.SelectedPath;
                SearchOption option = SearchOption.AllDirectories;

                string[] files;
                try
                {
                    files = Directory.GetFiles(folder, "*.*", option)
                        .Where(IsWordDocumentPath)
                        .ToArray();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("读取文件失败: " + ex.Message, "批量查找与替换");
                    return;
                }

                TreeNode root = new TreeNode(folder) { Checked = true, Tag = null };
                foreach (string file in files)
                {
                    TreeNode fileNode = new TreeNode(Path.GetFileName(file))
                    {
                        Checked = true,
                        Tag = file
                    };
                    root.Nodes.Add(fileNode);
                }

                root.Expand();
                tvFiles.Nodes.Add(root);
            }
        }

        private void AddFileNodes()
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "请选择文件";
                dialog.Multiselect = true;
                dialog.Filter = "Word 文档 (*.doc;*.docx;*.docm)|*.doc;*.docx;*.docm|所有文件 (*.*)|*.*";

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                foreach (string file in dialog.FileNames)
                {
                    AddFilePathNode(file);
                }
            }
        }

        private void AddActiveGroupFiles()
        {
            DocumentGroupStore store = new DocumentGroupStore();
            DocumentGroupCatalog catalog = store.Load();
            DocumentGroupItem activeGroup = store.EnsureActiveGroup(catalog, TryGetCurrentDocumentPathOrEmpty());
            if (activeGroup == null)
            {
                MessageBox.Show("当前没有活动文档组。", "批量查找与替换");
                return;
            }

            int addedCount = 0;
            foreach (DocumentGroupDocumentItem item in activeGroup.Documents ?? new List<DocumentGroupDocumentItem>())
            {
                if (item == null || string.IsNullOrWhiteSpace(item.FilePath))
                {
                    continue;
                }

                if (!IsWordDocumentPath(item.FilePath))
                {
                    continue;
                }

                int before = CountAllFileNodes();
                AddFilePathNode(item.FilePath);
                if (CountAllFileNodes() > before)
                {
                    addedCount++;
                }
            }

            if (addedCount == 0)
            {
                MessageBox.Show("活动组中没有可用的 Word 文件。", "批量查找与替换");
            }
        }

        private void AddFilePathNode(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return;
            }

            string folder = Path.GetDirectoryName(filePath) ?? "(未知目录)";
            TreeNode root = tvFiles.Nodes.Cast<TreeNode>()
                .FirstOrDefault(n => string.Equals(n.Text, folder, StringComparison.OrdinalIgnoreCase));

            if (root == null)
            {
                root = new TreeNode(folder) { Checked = true, Tag = null };
                tvFiles.Nodes.Add(root);
            }

            bool exists = root.Nodes.Cast<TreeNode>()
                .Any(n => n.Tag is string p && string.Equals(p, filePath, StringComparison.OrdinalIgnoreCase));
            if (exists)
            {
                return;
            }

            root.Nodes.Add(new TreeNode(Path.GetFileName(filePath))
            {
                Checked = true,
                Tag = filePath
            });
            root.Expand();
        }

        private void RemoveSelectedNode()
        {
            if (tvFiles.SelectedNode == null || tvFiles.SelectedNode.Parent == null)
            {
                MessageBox.Show("请先在文件列表中选中一个文件。", "批量查找与替换");
                return;
            }

            TreeNode selectedFileNode = tvFiles.SelectedNode;
            TreeNode parentNode = selectedFileNode.Parent;
            if (parentNode == null)
            {
                MessageBox.Show("请先在文件列表中选中一个文件。", "批量查找与替换");
                return;
            }

            parentNode.Nodes.Remove(selectedFileNode);
            if (parentNode.Nodes.Count == 0)
            {
                tvFiles.Nodes.Remove(parentNode);
            }
        }

        private void TvFiles_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            tvFiles.SelectedNode = e.Node;
        }

        private void Submit()
        {
            List<BatchReplaceRule> rules = BuildRules();
            if (rules.Count == 0)
            {
                MessageBox.Show("请至少配置一条有效规则。", "批量查找与替换");
                return;
            }

            List<string> files = GetCheckedFiles();
            if (files.Count == 0)
            {
                MessageBox.Show("请至少选择一个文件。", "批量查找与替换");
                return;
            }

            Request = new BatchReplaceExecutionRequest
            {
                FilePaths = files,
                Rules = rules,
                MatchCase = true,
                MatchWholeWord = true,
                FindOnly = false
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        private List<string> GetCheckedFiles()
        {
            List<string> files = new List<string>();
            foreach (TreeNode root in tvFiles.Nodes)
            {
                foreach (TreeNode child in root.Nodes)
                {
                    if (root.Checked && child.Checked && child.Tag is string path && File.Exists(path))
                    {
                        files.Add(path);
                    }
                }
            }

            return files.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private int CountAllFileNodes()
        {
            int count = 0;
            foreach (TreeNode root in tvFiles.Nodes)
            {
                count += root.Nodes.Count;
            }

            return count;
        }

        private static bool IsWordDocumentPath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return false;
            }

            string extension = Path.GetExtension(filePath);
            return string.Equals(extension, ".doc", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".docx", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".docm", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".dot", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".dotx", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(extension, ".dotm", StringComparison.OrdinalIgnoreCase);
        }

        private static string TryGetCurrentDocumentPathOrEmpty()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app?.ActiveDocument == null)
                {
                    return string.Empty;
                }

                return app.ActiveDocument.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private List<BatchReplaceRule> BuildRules()
        {
            List<BatchReplaceRule> rules = new List<BatchReplaceRule>();

            foreach (DataGridViewRow row in dgvRules.Rows)
            {
                string findText = Convert.ToString(row.Cells["colFindText"].Value) ?? string.Empty;
                string replaceText = Convert.ToString(row.Cells["colReplaceText"].Value) ?? string.Empty;

                if (string.IsNullOrWhiteSpace(findText))
                {
                    continue;
                }

                bool matchCase = true;
                bool matchWholeWord = true;
                bool matchSoundsLike = false;
                bool matchAllWordForms = false;

                BatchFindType findType = BatchFindType.PlainText;

                BatchFindScope scope = BatchFindScope.All;
                bool hasFileName = false;

                rules.Add(new BatchReplaceRule
                {
                    Enabled = true,
                    FindText = findText,
                    ReplaceText = replaceText,
                    FindType = findType,
                    Scope = scope,
                    ApplyToFileName = hasFileName,
                    HighlightOnly = false,
                    MatchCase = matchCase,
                    MatchWholeWord = matchWholeWord,
                    MatchSoundsLike = matchSoundsLike,
                    MatchAllWordForms = matchAllWordForms
                });
            }

            return rules;
        }
    }
}
