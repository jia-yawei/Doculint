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
            BackColor = Color.FromArgb(246, 248, 252);

            TabControl tabs = new TabControl
            {
                Dock = DockStyle.Fill,
                Padding = new Point(16, 6)
            };
            TabPage pageConfig = new TabPage("查找和替换")
            {
                BackColor = Color.FromArgb(246, 248, 252),
                Padding = new Padding(8)
            };
            tabs.TabPages.Add(pageConfig);
            Controls.Add(tabs);

            SplitContainer split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                BackColor = Color.FromArgb(218, 226, 238),
                SplitterWidth = 8
            };
            pageConfig.Controls.Add(split);
            Shown += (s, e) => BeginInvoke((Action)(() => ApplySafeSplitLayout(split, 500, 450, 700)));
            SizeChanged += (s, e) => ApplySafeSplitLayout(split, 500, 450, 700);

            TableLayoutPanel leftLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = Color.White,
                Padding = new Padding(12)
            };
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            leftLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            split.Panel2.Controls.Add(leftLayout);

            FlowLayoutPanel toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                WrapContents = false,
                Padding = new Padding(0, 8, 0, 6),
                BackColor = Color.White
            };
            leftLayout.Controls.Add(toolbar, 0, 0);

            Button btnAddRule = new Button { Text = "添加行", Width = 110, Height = 34 };
            Button btnDeleteRule = new Button { Text = "删除行", Width = 110, Height = 34 };
            ApplySecondaryButtonStyle(btnAddRule);
            ApplySecondaryButtonStyle(btnDeleteRule);
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
                RowTemplate = { Height = 36 },
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                EnableHeadersVisualStyles = false,
                GridColor = Color.FromArgb(229, 234, 242)
            };
            ApplyGridStyle(dgvRules);
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
                Padding = new Padding(12),
                BackColor = Color.White
            };
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            rightLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 52));
            split.Panel1.Controls.Add(rightLayout);

            TableLayoutPanel patternLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3
            };
            patternLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 112));
            patternLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            patternLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 116));
            rightLayout.Controls.Add(patternLayout, 0, 0);

            Label lblPattern = new Label { Text = "文件来源:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Font = new Font(Font, FontStyle.Bold) };
            patternLayout.Controls.Add(lblPattern, 0, 0);

            Label lblSourceHint = new Label
            {
                Dock = DockStyle.Fill,
                Text = "选择单个文件，或选择文件夹导入其中全部 Word 文档",
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(3, 10, 3, 8)
            };
            patternLayout.Controls.Add(lblSourceHint, 1, 0);

            Button btnBrowseSource = new Button { Text = "浏览...", Dock = DockStyle.Fill, Margin = new Padding(3, 8, 3, 8) };
            Button btnRemoveNode = new Button { Text = "移除", Width = 88, Height = 30, Margin = new Padding(0, 4, 0, 4) };
            ApplySecondaryButtonStyle(btnBrowseSource);
            ApplySecondaryButtonStyle(btnRemoveNode);
            patternLayout.Controls.Add(btnBrowseSource, 2, 0);

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
                Margin = new Padding(3, 8, 3, 3),
                Font = new Font(Font, FontStyle.Bold)
            };
            fileHeaderLayout.Controls.Add(lblTargetFiles, 0, 0);
            fileHeaderLayout.Controls.Add(btnRemoveNode, 1, 0);
            rightLayout.Controls.Add(fileHeaderLayout, 0, 1);

            tvFiles = new TreeView
            {
                CheckBoxes = true,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                BackColor = Color.White,
                Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
                ItemHeight = 26
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

            Button btnReplace = new Button { Text = "开始替换", Width = 112, Height = 36 };
            ApplyPrimaryButtonStyle(btnReplace);
            actionLayout.Controls.Add(btnReplace);

            btnAddRule.Click += (s, e) => AddDefaultRule();
            btnDeleteRule.Click += (s, e) => DeleteSelectedRules();
            btnBrowseSource.Click += (s, e) => ShowSourceBrowseMenu(btnBrowseSource);
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

        private void ShowSourceBrowseMenu(Control anchor)
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem fileItem = new ToolStripMenuItem("选择文件...");
            fileItem.ToolTipText = "导入一个或多个指定的 Word 文档";
            fileItem.Click += (_, __) => AddFileNodes();
            ToolStripMenuItem folderItem = new ToolStripMenuItem("选择文件夹...");
            folderItem.ToolTipText = "导入所选文件夹及其子文件夹中的全部 Word 文档";
            folderItem.Click += (_, __) => AddFolderNode();
            menu.Items.Add(fileItem);
            menu.Items.Add(folderItem);
            menu.Closed += (_, __) => menu.Dispose();
            menu.Show(anchor, new Point(0, anchor.Height));
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

        private static void ApplyGridStyle(DataGridView grid)
        {
            if (grid == null)
            {
                return;
            }

            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 247, 250);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(43, 57, 76);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(225, 239, 255);
            grid.DefaultCellStyle.SelectionForeColor = Color.Black;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(250, 252, 255);
            EnableDoubleBuffering(grid);
        }

        private static void ApplyPrimaryButtonStyle(Button button)
        {
            if (button == null)
            {
                return;
            }

            button.UseVisualStyleBackColor = false;
            button.BackColor = Color.FromArgb(42, 122, 226);
            button.ForeColor = Color.White;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(58, 136, 236);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(33, 105, 199);
            button.Cursor = Cursors.Hand;
        }

        private static void ApplySecondaryButtonStyle(Button button)
        {
            if (button == null)
            {
                return;
            }

            button.UseVisualStyleBackColor = false;
            button.BackColor = Color.FromArgb(235, 243, 255);
            button.ForeColor = Color.FromArgb(36, 89, 171);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = Color.FromArgb(176, 203, 240);
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(223, 236, 255);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(209, 227, 253);
            button.Cursor = Cursors.Hand;
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
