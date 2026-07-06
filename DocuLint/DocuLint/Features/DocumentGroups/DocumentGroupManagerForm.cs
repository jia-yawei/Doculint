using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class DocumentGroupManagerForm : Form
    {
        private readonly DocumentGroupStore store;
        private readonly DocumentGroupCatalog catalog;
        private readonly string currentDocumentPath;
        private readonly Font activeGroupNodeFont;
        private readonly Font selectedDocumentNodeFont;
        private TreeView groupTreeView;
        private TextBox pathTextBox;
        private ImageList treeImageList;
        private ContextMenuStrip groupContextMenu;
        private ContextMenuStrip documentContextMenu;
        private Button createGroupButton;
        private Button addOtherDocumentsButton;
        private Button addCurrentDocumentButton;
        private Button moveButton;
        private Button setActiveButton;
        private Button renameButton;
        private Button deleteGroupButton;
        private Button closeButton;

        public bool DataChanged { get; private set; }

        public DocumentGroupManagerForm(DocumentGroupStore store, DocumentGroupCatalog catalog, string currentDocumentPath)
        {
            this.store = store;
            this.catalog = catalog;
            this.currentDocumentPath = currentDocumentPath;
            this.store.EnsureActiveGroup(this.catalog, this.currentDocumentPath);

            Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            activeGroupNodeFont = new Font(Font, FontStyle.Bold);
            selectedDocumentNodeFont = new Font("Microsoft YaHei", 10.5F, FontStyle.Bold, GraphicsUnit.Point, 134);
            Text = "文档组管理";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(860, 560);
            Size = new Size(940, 640);

            groupTreeView = new NoToolTipTreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false,
                FullRowSelect = true,
                BorderStyle = BorderStyle.None,
                Indent = 22,
                ShowLines = false,
                ShowPlusMinus = true,
                ShowRootLines = false,
                ItemHeight = 26,
                BackColor = Color.FromArgb(250, 252, 255),
                DrawMode = TreeViewDrawMode.OwnerDrawText,
                ShowNodeToolTips = false
            };
            groupTreeView.AfterSelect += (sender, e) => RefreshSelectionDetail();
            groupTreeView.NodeMouseDoubleClick += GroupTreeView_NodeMouseDoubleClick;
            groupTreeView.NodeMouseClick += GroupTreeView_NodeMouseClick;
            groupTreeView.DrawNode += GroupTreeView_DrawNode;

            InitializeTreeImages();
            InitializeContextMenus();

            TableLayoutPanel rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(16, 14, 16, 14)
            };
            rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            rootLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 168F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 88F));

            Panel treePanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 0, 12, 12)
            };
            treePanel.Controls.Add(groupTreeView);

            Panel rightPanel = BuildRightPanel();

            pathTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(250, 252, 255)
            };

            rootLayout.Controls.Add(treePanel, 0, 0);
            rootLayout.Controls.Add(rightPanel, 1, 0);
            rootLayout.Controls.Add(pathTextBox, 0, 1);
            rootLayout.SetColumnSpan(pathTextBox, 1);

            Panel closeHost = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 0, 0, 0)
            };
            closeButton = CreateActionButton("关闭", (sender, e) => Close());
            ApplyButtonTheme(closeButton, Color.FromArgb(236, 241, 250), Color.FromArgb(64, 76, 96));
            closeButton.Dock = DockStyle.Bottom;
            closeHost.Controls.Add(closeButton);
            rootLayout.Controls.Add(closeHost, 1, 1);

            Controls.Add(rootLayout);
            Load += (sender, e) => RefreshTree();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                activeGroupNodeFont?.Dispose();
                selectedDocumentNodeFont?.Dispose();
            }

            base.Dispose(disposing);
        }

        private Panel BuildRightPanel()
        {
            Panel rightPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 2, 0, 0)
            };

            FlowLayoutPanel actionPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };

            createGroupButton = CreateActionButton("新建组", (sender, e) => CreateGroup());
            addOtherDocumentsButton = CreateActionButton("加入其他文档", (sender, e) => AddFiles());
            addCurrentDocumentButton = CreateActionButton("加入当前文档", (sender, e) => AddCurrentDocument());
            setActiveButton = CreateActionButton("设为活动", (sender, e) => SetActiveGroup());
            moveButton = CreateActionButton("移动", (sender, e) => MoveSelectedDocument());
            renameButton = CreateActionButton("改名", (sender, e) => RenameGroup());
            deleteGroupButton = CreateActionButton("删除", (sender, e) => DeleteSelection());
            ApplyActionButtonStyles();

            actionPanel.Controls.Add(createGroupButton);
            actionPanel.Controls.Add(addOtherDocumentsButton);
            actionPanel.Controls.Add(addCurrentDocumentButton);
            actionPanel.Controls.Add(setActiveButton);
            actionPanel.Controls.Add(moveButton);
            actionPanel.Controls.Add(renameButton);
            actionPanel.Controls.Add(deleteGroupButton);

            rightPanel.Controls.Add(actionPanel);
            return rightPanel;
        }

        private Button CreateActionButton(string text, EventHandler onClick)
        {
            Button button = new Button
            {
                Text = text,
                Width = 168,
                Height = 38,
                Margin = new Padding(0, 0, 0, 12),
                Font = new Font("Microsoft YaHei", 10.5F, FontStyle.Bold, GraphicsUnit.Point, 134),
                FlatStyle = FlatStyle.Flat,
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            button.Click += onClick;
            return button;
        }

        private void ApplyActionButtonStyles()
        {
            Color primary = Color.FromArgb(221, 235, 255);
            Color primaryText = Color.FromArgb(44, 91, 173);
            Color danger = Color.FromArgb(255, 232, 232);
            Color dangerText = Color.FromArgb(171, 58, 58);

            ApplyButtonTheme(createGroupButton, primary, primaryText);
            ApplyButtonTheme(addOtherDocumentsButton, primary, primaryText);
            ApplyButtonTheme(addCurrentDocumentButton, primary, primaryText);
            ApplyButtonTheme(setActiveButton, primary, primaryText);
            ApplyButtonTheme(moveButton, primary, primaryText);
            ApplyButtonTheme(renameButton, primary, primaryText);
            ApplyButtonTheme(deleteGroupButton, danger, dangerText);
        }

        private static void ApplyButtonTheme(Button button, Color backColor, Color foreColor)
        {
            if (button == null)
            {
                return;
            }

            button.BackColor = backColor;
            button.ForeColor = foreColor;
            button.FlatAppearance.MouseOverBackColor = ShiftBrightness(backColor, 12);
            button.FlatAppearance.MouseDownBackColor = ShiftBrightness(backColor, -18);
        }

        private static Color ShiftBrightness(Color color, int shift)
        {
            int r = Math.Max(0, Math.Min(255, color.R + shift));
            int g = Math.Max(0, Math.Min(255, color.G + shift));
            int b = Math.Max(0, Math.Min(255, color.B + shift));
            return Color.FromArgb(r, g, b);
        }

        private DocumentGroupItem SelectedGroup
        {
            get
            {
                TreeNode node = groupTreeView.SelectedNode;
                if (node?.Tag is DocumentGroupItem selectedGroup)
                {
                    return selectedGroup;
                }

                if (node?.Tag is DocumentTreeNodeTag tag)
                {
                    return tag.Group;
                }

                return null;
            }
        }

        private DocumentGroupDocumentItem SelectedDocument
        {
            get
            {
                if (groupTreeView.SelectedNode?.Tag is DocumentTreeNodeTag tag)
                {
                    return tag.Document;
                }

                return null;
            }
        }

        private void RefreshTree()
        {
            store.EnsureActiveGroup(catalog, currentDocumentPath);
            string selectedGroupId = SelectedGroup?.Id;
            string selectedDocumentPath = SelectedDocument?.FilePath;

            groupTreeView.BeginUpdate();
            groupTreeView.Nodes.Clear();

            foreach (DocumentGroupItem group in catalog.GetOrderedGroups())
            {
                bool isActive = string.Equals(group.Id, catalog.ActiveGroupId, StringComparison.OrdinalIgnoreCase);
                TreeNode groupNode = new TreeNode(isActive ? $"{group.Name}  [活动]" : group.Name)
                {
                    Tag = group,
                    ImageKey = isActive ? "group-active" : "group",
                    SelectedImageKey = isActive ? "group-active" : "group",
                    NodeFont = isActive ? activeGroupNodeFont : null,
                    ForeColor = isActive ? Color.FromArgb(24, 92, 188) : Color.FromArgb(56, 56, 56)
                };

                foreach (DocumentGroupDocumentItem document in group.Documents.OrderByDescending(item => item.LastKnownWriteTime ?? DateTime.MinValue))
                {
                    string displayName = document.DisplayName ?? Path.GetFileName(document.FilePath);
                    string lastSavedText = document.LastKnownWriteTime.HasValue
                        ? document.LastKnownWriteTime.Value.ToString("yyyy-MM-dd HH:mm:ss")
                        : "未知";
                    bool exists = !string.IsNullOrWhiteSpace(document.FilePath) && File.Exists(document.FilePath);

                    TreeNode documentNode = new TreeNode($"{displayName} [{lastSavedText}]")
                    {
                        Tag = new DocumentTreeNodeTag(group, document),
                        ImageKey = exists ? "doc" : "doc-missing",
                        SelectedImageKey = exists ? "doc" : "doc-missing",
                        ForeColor = exists ? Color.FromArgb(80, 80, 80) : Color.FromArgb(176, 62, 48)
                    };
                    groupNode.Nodes.Add(documentNode);
                }

                groupTreeView.Nodes.Add(groupNode);
                groupNode.Expand();
            }

            groupTreeView.EndUpdate();
            RestoreSelection(selectedGroupId, selectedDocumentPath);
            RefreshSelectionDetail();
        }

        private void RestoreSelection(string selectedGroupId, string selectedDocumentPath)
        {
            if (groupTreeView.Nodes.Count == 0)
            {
                return;
            }

            foreach (TreeNode groupNode in groupTreeView.Nodes)
            {
                DocumentGroupItem group = groupNode.Tag as DocumentGroupItem;
                if (!string.Equals(group?.Id, selectedGroupId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(selectedDocumentPath))
                {
                    foreach (TreeNode documentNode in groupNode.Nodes)
                    {
                        DocumentTreeNodeTag tag = documentNode.Tag as DocumentTreeNodeTag;
                        if (string.Equals(tag?.Document?.FilePath, selectedDocumentPath, StringComparison.OrdinalIgnoreCase))
                        {
                            groupTreeView.SelectedNode = documentNode;
                            return;
                        }
                    }
                }

                groupTreeView.SelectedNode = groupNode;
                return;
            }

            groupTreeView.SelectedNode = groupTreeView.Nodes[0];
        }

        private void RefreshSelectionDetail()
        {
            bool isDocumentSelected = SelectedDocument != null;
            if (isDocumentSelected)
            {
                pathTextBox.Text = SelectedDocument.FilePath ?? string.Empty;
            }
            else if (SelectedGroup != null)
            {
                bool isActiveGroup = string.Equals(SelectedGroup.Id, catalog.ActiveGroupId, StringComparison.OrdinalIgnoreCase);
                int documentCount = SelectedGroup.Documents?.Count ?? 0;
                pathTextBox.Text = $"文档组：{SelectedGroup.Name}\r\n状态：{(isActiveGroup ? "活动文档组" : "普通文档组")}\r\n文档数：{documentCount}";
            }
            else
            {
                pathTextBox.Text = string.Empty;
            }

            moveButton.Enabled = isDocumentSelected;
            addOtherDocumentsButton.Enabled = SelectedGroup != null;
            addCurrentDocumentButton.Enabled = SelectedGroup != null;
            setActiveButton.Enabled = SelectedGroup != null &&
                !string.Equals(SelectedGroup.Id, catalog.ActiveGroupId, StringComparison.OrdinalIgnoreCase);
            renameButton.Enabled = SelectedGroup != null;
            deleteGroupButton.Enabled = SelectedGroup != null || isDocumentSelected;
        }

        private void SetActiveGroup()
        {
            DocumentGroupItem group = SelectedGroup;
            if (group == null)
            {
                MessageBox.Show(this, "请先在左侧树中选择一个文档组。", "文档组管理");
                return;
            }

            try
            {
                store.SetActiveGroup(catalog, group.Id);
                DataChanged = true;
                RefreshTree();
                SelectGroup(group.Id);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "文档组管理");
            }
        }

        private void CreateGroup()
        {
            using (TextPromptForm prompt = new TextPromptForm("新建文档组", "请输入新的文档组名称："))
            {
                if (prompt.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    DocumentGroupItem created = store.CreateGroup(catalog, prompt.ResultText);
                    store.EnsureActiveGroup(catalog, currentDocumentPath);
                    DataChanged = true;
                    RefreshTree();
                    SelectGroup(created.Id);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "文档组管理");
                }
            }
        }

        private void RenameGroup()
        {
            DocumentGroupItem group = SelectedGroup;
            if (group == null)
            {
                MessageBox.Show(this, "请先在左侧树中选择一个文档组。", "文档组管理");
                return;
            }

            using (TextPromptForm prompt = new TextPromptForm("重命名文档组", "请输入新的文档组名称：", group.Name))
            {
                if (prompt.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    store.RenameGroup(catalog, group.Id, prompt.ResultText);
                    DataChanged = true;
                    RefreshTree();
                    SelectGroup(group.Id);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "文档组管理");
                }
            }
        }

        private void DeleteGroup()
        {
            DocumentGroupItem group = SelectedGroup;
            if (group == null)
            {
                MessageBox.Show(this, "请先在左侧树中选择要删除的文档组。", "文档组管理");
                return;
            }

            DialogResult result = MessageBox.Show(
                this,
                $"确定删除文档组“{group.Name}”吗？\r\n组内文档记录也会一起移除。",
                "文档组管理",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
            {
                return;
            }

            store.DeleteGroup(catalog, group.Id);
            store.EnsureActiveGroup(catalog, currentDocumentPath);
            DataChanged = true;
            RefreshTree();
        }

        private void AddCurrentDocument()
        {
            DocumentGroupItem group = SelectedGroup;
            if (group == null)
            {
                MessageBox.Show(this, "请先在左侧树中选择目标文档组。", "文档组管理");
                return;
            }

            if (string.IsNullOrWhiteSpace(currentDocumentPath))
            {
                MessageBox.Show(this, "当前文档还没有保存，无法加入文档组。", "文档组管理");
                return;
            }

            try
            {
                if (!AddDocumentToGroupWithPrompt(group, currentDocumentPath))
                {
                    return;
                }

                store.RefreshDocumentMetadata(catalog);
                DataChanged = true;
                RefreshTree();
                SelectDocument(group.Id, currentDocumentPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "文档组管理");
            }
        }

        private void AddFiles()
        {
            DocumentGroupItem group = SelectedGroup;
            if (group == null)
            {
                MessageBox.Show(this, "请先在左侧树中选择目标文档组。", "文档组管理");
                return;
            }

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择要加入文档组的文档";
                dialog.Multiselect = true;
                dialog.Filter = "Word 文档|*.doc;*.docx;*.docm;*.dot;*.dotx;*.dotm|所有文件|*.*";
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    foreach (string fileName in dialog.FileNames)
                    {
                        AddDocumentToGroupWithPrompt(group, fileName);
                    }

                    store.RefreshDocumentMetadata(catalog);
                    DataChanged = true;
                    RefreshTree();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "文档组管理");
                }
            }
        }

        private void RemoveSelectedDocument()
        {
            DocumentGroupItem group = SelectedGroup;
            DocumentGroupDocumentItem document = SelectedDocument;
            if (group == null || document == null)
            {
                MessageBox.Show(this, "请先在左侧树中选择一个文档。", "文档组管理");
                return;
            }

            store.RemoveDocumentFromGroup(catalog, group.Id, document.FilePath);
            DataChanged = true;
            RefreshTree();
            SelectGroup(group.Id);
        }

        private void MoveSelectedDocument()
        {
            DocumentGroupItem sourceGroup = SelectedGroup;
            DocumentGroupDocumentItem document = SelectedDocument;
            if (sourceGroup == null || document == null)
            {
                MessageBox.Show(this, "请先在左侧树中选择一个文档。", "文档组管理");
                return;
            }

            DocumentGroupItem[] targetGroups = catalog.GetOrderedGroups()
                .Where(item => !string.Equals(item.Id, sourceGroup.Id, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (targetGroups.Length == 0)
            {
                MessageBox.Show(this, "当前没有其他文档组可供移动。", "文档组管理");
                return;
            }

            using (DocumentGroupPickerForm picker = new DocumentGroupPickerForm(
                targetGroups,
                document.DisplayName ?? Path.GetFileName(document.FilePath),
                catalog.ActiveGroupId,
                "移动文档",
                "选择要移动到的目标文档组：",
                "移动到选中组"))
            {
                if (picker.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    store.RemoveDocumentFromGroup(catalog, sourceGroup.Id, document.FilePath);
                    DocumentGroupItem targetGroup = catalog.Groups.FirstOrDefault(item =>
                        string.Equals(item.Id, picker.SelectedGroupId, StringComparison.OrdinalIgnoreCase));
                    if (targetGroup == null)
                    {
                        throw new InvalidOperationException("未找到目标文档组。");
                    }

                    if (!AddDocumentToGroupWithPrompt(targetGroup, document.FilePath))
                    {
                        store.AddDocumentToGroup(catalog, sourceGroup.Id, document.FilePath);
                    }

                    store.RefreshDocumentMetadata(catalog);
                    DataChanged = true;
                    RefreshTree();
                    SelectDocument(picker.SelectedGroupId, document.FilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, "文档组管理");
                }
            }
        }

        private void DeleteSelection()
        {
            if (SelectedDocument != null)
            {
                RemoveSelectedDocument();
                return;
            }

            DeleteGroup();
        }

        private void OpenSelectedDocument()
        {
            DocumentGroupDocumentItem document = SelectedDocument;
            if (document == null)
            {
                MessageBox.Show(this, "请先在左侧树中选择要打开的文档。", "文档组管理");
                return;
            }

            if (!File.Exists(document.FilePath))
            {
                MessageBox.Show(this, "该文档文件不存在，可能已被移动或删除。", "文档组管理");
                return;
            }

            Process.Start(document.FilePath);
        }

        private bool AddDocumentToGroupWithPrompt(DocumentGroupItem group, string filePath)
        {
            if (group == null)
            {
                return false;
            }

            string normalizedPath = Path.GetFullPath((filePath ?? string.Empty).Trim());
            string displayName = Path.GetFileName(normalizedPath);
            DocumentGroupDocumentItem existingByName = (group.Documents ?? new System.Collections.Generic.List<DocumentGroupDocumentItem>())
                .FirstOrDefault(item => string.Equals(item?.DisplayName ?? Path.GetFileName(item?.FilePath ?? string.Empty), displayName, StringComparison.CurrentCultureIgnoreCase));

            if (existingByName != null)
            {
                DialogResult replaceResult = MessageBox.Show(
                    this,
                    $"文档“{displayName}”已经存在，是否替换为新文档？",
                    "文档组管理",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (replaceResult != DialogResult.Yes)
                {
                    return false;
                }

                store.RemoveDocumentFromGroup(catalog, group.Id, existingByName.FilePath);
            }

            store.AddDocumentToGroup(catalog, group.Id, normalizedPath);
            return true;
        }

        private void GroupTreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node?.Tag is DocumentTreeNodeTag)
            {
                OpenSelectedDocument();
            }
        }

        private void GroupTreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button != MouseButtons.Right || e.Node == null)
            {
                return;
            }

            groupTreeView.SelectedNode = e.Node;

            if (e.Node.Tag is DocumentTreeNodeTag)
            {
                documentContextMenu.Show(groupTreeView, e.Location);
                return;
            }

            if (e.Node.Tag is DocumentGroupItem)
            {
                groupContextMenu.Show(groupTreeView, e.Location);
            }
        }

        private void GroupTreeView_DrawNode(object sender, DrawTreeNodeEventArgs e)
        {
            if (e.Node == null)
            {
                return;
            }

            bool isDocumentNode = e.Node.Tag is DocumentTreeNodeTag;
            if (!isDocumentNode)
            {
                e.DrawDefault = true;
                return;
            }

            bool isSelected = (e.State & TreeNodeStates.Selected) == TreeNodeStates.Selected;

            Color foreColor = e.Node.ForeColor.IsEmpty ? groupTreeView.ForeColor : e.Node.ForeColor;
            Font textFont = e.Node.NodeFont ?? groupTreeView.Font;
            if (isSelected)
            {
                foreColor = Color.FromArgb(24, 92, 188);
                textFont = selectedDocumentNodeFont;
            }

            using (SolidBrush backgroundBrush = new SolidBrush(groupTreeView.BackColor))
            {
                Rectangle fullRowBounds = new Rectangle(
                    0,
                    e.Bounds.Top,
                    groupTreeView.ClientSize.Width,
                    e.Bounds.Height);
                e.Graphics.FillRectangle(backgroundBrush, fullRowBounds);
                TextRenderer.DrawText(
                    e.Graphics,
                    e.Node.Text ?? string.Empty,
                    textFont,
                    e.Bounds,
                    foreColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
        }

        private void SelectGroup(string groupId)
        {
            foreach (TreeNode groupNode in groupTreeView.Nodes)
            {
                DocumentGroupItem group = groupNode.Tag as DocumentGroupItem;
                if (string.Equals(group?.Id, groupId, StringComparison.OrdinalIgnoreCase))
                {
                    groupTreeView.SelectedNode = groupNode;
                    groupNode.EnsureVisible();
                    return;
                }
            }
        }

        private void SelectDocument(string groupId, string filePath)
        {
            foreach (TreeNode groupNode in groupTreeView.Nodes)
            {
                DocumentGroupItem group = groupNode.Tag as DocumentGroupItem;
                if (!string.Equals(group?.Id, groupId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                foreach (TreeNode documentNode in groupNode.Nodes)
                {
                    DocumentTreeNodeTag tag = documentNode.Tag as DocumentTreeNodeTag;
                    if (string.Equals(tag?.Document?.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        groupTreeView.SelectedNode = documentNode;
                        documentNode.EnsureVisible();
                        return;
                    }
                }
            }
        }

        private sealed class DocumentTreeNodeTag
        {
            public DocumentTreeNodeTag(DocumentGroupItem group, DocumentGroupDocumentItem document)
            {
                Group = group;
                Document = document;
            }

            public DocumentGroupItem Group { get; }

            public DocumentGroupDocumentItem Document { get; }
        }

        private sealed class NoToolTipTreeView : TreeView
        {
            private const int TvsNotooltips = 0x0080;

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams cp = base.CreateParams;
                    cp.Style |= TvsNotooltips;
                    return cp;
                }
            }
        }

        private void InitializeTreeImages()
        {
            treeImageList = new ImageList
            {
                ColorDepth = ColorDepth.Depth32Bit,
                ImageSize = new Size(16, 16)
            };
            treeImageList.Images.Add("group", CreateGroupIcon(Color.FromArgb(80, 130, 215), Color.FromArgb(226, 238, 255)));
            treeImageList.Images.Add("group-active", CreateGroupIcon(Color.FromArgb(31, 160, 96), Color.FromArgb(220, 247, 233)));
            treeImageList.Images.Add("doc", CreateDocumentIcon(Color.FromArgb(66, 118, 220), Color.White));
            treeImageList.Images.Add("doc-missing", CreateDocumentIcon(Color.FromArgb(196, 68, 68), Color.FromArgb(255, 243, 243)));
            groupTreeView.ImageList = treeImageList;
        }

        private static Bitmap CreateGroupIcon(Color accent, Color fill)
        {
            Bitmap bitmap = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (SolidBrush fillBrush = new SolidBrush(fill))
                using (Pen borderPen = new Pen(accent, 1.2f))
                {
                    Rectangle tabRect = new Rectangle(2, 3, 6, 4);
                    Rectangle bodyRect = new Rectangle(1, 6, 14, 8);
                    g.FillRectangle(fillBrush, tabRect);
                    g.FillRectangle(fillBrush, bodyRect);
                    g.DrawRectangle(borderPen, tabRect);
                    g.DrawRectangle(borderPen, bodyRect);
                }
            }

            return bitmap;
        }

        private static Bitmap CreateDocumentIcon(Color accent, Color fill)
        {
            Bitmap bitmap = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (SolidBrush fillBrush = new SolidBrush(fill))
                using (Pen borderPen = new Pen(accent, 1.2f))
                using (SolidBrush cornerBrush = new SolidBrush(ShiftBrightness(fill, -16)))
                {
                    Rectangle body = new Rectangle(3, 1, 10, 13);
                    g.FillRectangle(fillBrush, body);
                    g.DrawRectangle(borderPen, body);

                    Point[] foldedCorner =
                    {
                        new Point(10, 1),
                        new Point(13, 4),
                        new Point(10, 4)
                    };
                    g.FillPolygon(cornerBrush, foldedCorner);
                    g.DrawLine(borderPen, 10, 1, 13, 4);

                    using (Pen linePen = new Pen(ShiftBrightness(accent, 20), 1f))
                    {
                        g.DrawLine(linePen, 5, 7, 11, 7);
                        g.DrawLine(linePen, 5, 9, 11, 9);
                        g.DrawLine(linePen, 5, 11, 9, 11);
                    }
                }
            }

            return bitmap;
        }

        private void InitializeContextMenus()
        {
            groupContextMenu = new ContextMenuStrip
            {
                Font = Font
            };
            groupContextMenu.Items.Add("新建文档组", null, (sender, e) => CreateGroup());
            groupContextMenu.Items.Add("设为活动文档组", null, (sender, e) => SetActiveGroup());
            groupContextMenu.Items.Add("重命名当前组", null, (sender, e) => RenameGroup());
            groupContextMenu.Items.Add("删除当前组", null, (sender, e) => DeleteGroup());
            groupContextMenu.Items.Add(new ToolStripSeparator());
            groupContextMenu.Items.Add("加入其他文档", null, (sender, e) => AddFiles());
            groupContextMenu.Items.Add("加入当前文档", null, (sender, e) => AddCurrentDocument());

            documentContextMenu = new ContextMenuStrip
            {
                Font = Font
            };
            documentContextMenu.Items.Add("打开当前文档", null, (sender, e) => OpenSelectedDocument());
            documentContextMenu.Items.Add("移动到其他组", null, (sender, e) => MoveSelectedDocument());
            documentContextMenu.Items.Add("移出当前文档", null, (sender, e) => RemoveSelectedDocument());
        }

        private static class ShellIconHelper
        {
            public static Bitmap GetSmallFolderBitmap()
            {
                return GetShellBitmap(@"C:\", NativeMethods.FileAttributeDirectory);
            }

            public static Bitmap GetSmallFileBitmap(string extension)
            {
                return GetShellBitmap("sample" + extension, NativeMethods.FileAttributeNormal);
            }

            private static Bitmap GetShellBitmap(string path, uint attributes)
            {
                NativeMethods.SHFILEINFO info = new NativeMethods.SHFILEINFO();
                IntPtr result = NativeMethods.SHGetFileInfo(
                    path,
                    attributes,
                    ref info,
                    (uint)Marshal.SizeOf(info),
                    NativeMethods.ShgfiIcon | NativeMethods.ShgfiSmallIcon | NativeMethods.ShgfiUseFileAttributes);

                if (result == IntPtr.Zero || info.hIcon == IntPtr.Zero)
                {
                    return SystemIcons.Application.ToBitmap();
                }

                try
                {
                    using (Icon icon = Icon.FromHandle(info.hIcon))
                    {
                        return icon.ToBitmap();
                    }
                }
                finally
                {
                    NativeMethods.DestroyIcon(info.hIcon);
                }
            }
        }

        private static class NativeMethods
        {
            public const uint ShgfiIcon = 0x000000100;
            public const uint ShgfiSmallIcon = 0x000000001;
            public const uint ShgfiUseFileAttributes = 0x000000010;
            public const uint FileAttributeDirectory = 0x00000010;
            public const uint FileAttributeNormal = 0x00000080;

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
            public struct SHFILEINFO
            {
                public IntPtr hIcon;
                public int iIcon;
                public uint dwAttributes;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public string szDisplayName;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
                public string szTypeName;
            }

            [DllImport("shell32.dll", CharSet = CharSet.Auto)]
            public static extern IntPtr SHGetFileInfo(
                string pszPath,
                uint dwFileAttributes,
                ref SHFILEINFO psfi,
                uint cbFileInfo,
                uint uFlags);

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool DestroyIcon(IntPtr hIcon);
        }
    }
}
