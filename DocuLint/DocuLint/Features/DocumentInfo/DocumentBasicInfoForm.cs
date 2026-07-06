using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class DocumentBasicInfoForm : Form
    {
        private readonly DocumentBasicInfo workingInfo;

        private readonly ListView listView;
        private readonly Label lblEmptyState;
        private readonly Button btnAdd;
        private readonly Button btnEdit;
        private readonly Button btnDelete;
        private readonly Button btnSaveDocument;
        private readonly Button btnClose;

        internal DocumentBasicInfoForm(DocumentBasicInfo info)
        {
            workingInfo = CloneInfo(info);

            Text = "基本信息";
            Width = 1460;
            Height = 980;
            MinimumSize = new Size(1320, 860);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.FromArgb(244, 247, 250);
            Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(20)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 68f));

            Panel headerPanel = CreateSurfacePanel(new Padding(20, 12, 20, 8));
            Label lblTitle = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 14F, FontStyle.Bold, GraphicsUnit.Point),
                ForeColor = Color.FromArgb(30, 43, 60),
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "文档基本信息"
            };
            headerPanel.Controls.Add(lblTitle);

            TableLayoutPanel bodyLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2
            };
            bodyLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            bodyLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220f));

            Panel leftPanel = CreateSurfacePanel(new Padding(18));
            lblEmptyState = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei UI", 14F, FontStyle.Regular, GraphicsUnit.Point),
                ForeColor = Color.FromArgb(130, 138, 148),
                TextAlign = ContentAlignment.MiddleCenter,
                Text = "当前没有基本信息"
            };

            listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                GridLines = false,
                MultiSelect = false,
                BorderStyle = BorderStyle.None,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            listView.Columns.Add("基本信息类型", 260, HorizontalAlignment.Left);
            listView.Columns.Add("内容", 760, HorizontalAlignment.Left);
            listView.SelectedIndexChanged += (_, __) => RefreshSelectionState();
            listView.Resize += (_, __) => ResizeColumns();

            leftPanel.Controls.Add(listView);
            leftPanel.Controls.Add(lblEmptyState);

            Panel rightPanel = CreateSurfacePanel(new Padding(18, 18, 18, 18));
            FlowLayoutPanel actionPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };

            btnAdd = CreateActionButton("增加");
            btnEdit = CreateActionButton("修改");
            btnDelete = CreateDangerActionButton("删除");
            actionPanel.Controls.Add(btnAdd);
            actionPanel.Controls.Add(btnEdit);
            actionPanel.Controls.Add(btnDelete);
            rightPanel.Controls.Add(actionPanel);

            bodyLayout.Controls.Add(leftPanel, 0, 0);
            bodyLayout.Controls.Add(rightPanel, 1, 0);

            FlowLayoutPanel bottomButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 12, 0, 0)
            };
            btnSaveDocument = CreatePrimaryActionButton("保存");
            btnClose = CreateActionButton("关闭");
            bottomButtons.Controls.Add(btnSaveDocument);
            bottomButtons.Controls.Add(btnClose);

            root.Controls.Add(headerPanel, 0, 0);
            root.Controls.Add(bodyLayout, 0, 1);
            root.Controls.Add(bottomButtons, 0, 2);
            Controls.Add(root);

            btnAdd.Click += (_, __) => AddField();
            btnEdit.Click += (_, __) => EditField();
            btnDelete.Click += (_, __) => DeleteField();
            btnSaveDocument.Click += (_, __) => { DialogResult = DialogResult.OK; Close(); };
            btnClose.Click += (_, __) => Close();

            AcceptButton = btnSaveDocument;
            CancelButton = btnClose;

            RefreshList();
            RefreshSelectionState();
        }

        internal DocumentBasicInfo BuildInfo()
        {
            return CloneInfo(workingInfo);
        }

        private void AddField()
        {
            using (DocumentBasicInfoItemEditForm form = new DocumentBasicInfoItemEditForm("增加基本信息", null))
            {
                if (form.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                if (!form.TryBuildField(out DocumentBasicInfoField field))
                {
                    return;
                }

                workingInfo.Fields.Add(field);
                RefreshList();
                SelectIndex(workingInfo.Fields.Count - 1);
            }
        }

        private void EditField()
        {
            int index = GetSelectedIndex();
            if (index < 0 || index >= workingInfo.Fields.Count)
            {
                return;
            }

            using (DocumentBasicInfoItemEditForm form = new DocumentBasicInfoItemEditForm("修改基本信息", workingInfo.Fields[index]))
            {
                if (form.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                if (!form.TryBuildField(out DocumentBasicInfoField field))
                {
                    return;
                }

                workingInfo.Fields[index] = field;
                RefreshList();
                SelectIndex(index);
            }
        }

        private void DeleteField()
        {
            int index = GetSelectedIndex();
            if (index < 0 || index >= workingInfo.Fields.Count)
            {
                return;
            }

            workingInfo.Fields.RemoveAt(index);
            RefreshList();
            if (workingInfo.Fields.Count > 0)
            {
                SelectIndex(Math.Min(index, workingInfo.Fields.Count - 1));
            }
        }

        private void RefreshList()
        {
            listView.BeginUpdate();
            try
            {
                listView.Items.Clear();
                foreach (DocumentBasicInfoField field in workingInfo.Fields)
                {
                    string fieldName = string.IsNullOrWhiteSpace(field?.Name) ? "未命名类型" : field.Name.Trim();
                    string fieldValue = string.IsNullOrWhiteSpace(field?.Value) ? "未填写" : NormalizeSingleLine(field.Value);

                    ListViewItem item = new ListViewItem(fieldName);
                    item.SubItems.Add(fieldValue);
                    item.Tag = field;
                    listView.Items.Add(item);
                }
            }
            finally
            {
                listView.EndUpdate();
            }

            lblEmptyState.Visible = listView.Items.Count == 0;
            listView.Visible = listView.Items.Count > 0;
            ResizeColumns();
        }

        private void RefreshSelectionState()
        {
            bool hasSelection = GetSelectedIndex() >= 0;
            btnEdit.Enabled = hasSelection;
            btnDelete.Enabled = hasSelection;
        }

        private int GetSelectedIndex()
        {
            return listView.SelectedIndices.Count > 0 ? listView.SelectedIndices[0] : -1;
        }

        private void SelectIndex(int index)
        {
            if (index < 0 || index >= listView.Items.Count)
            {
                return;
            }

            listView.Items[index].Selected = true;
            listView.Items[index].Focused = true;
            listView.EnsureVisible(index);
        }

        private void ResizeColumns()
        {
            if (listView.Columns.Count < 2)
            {
                return;
            }

            int clientWidth = Math.Max(200, listView.ClientSize.Width);
            int firstColumnWidth = Math.Max(260, (int)(clientWidth * 0.34));
            int secondColumnWidth = Math.Max(300, clientWidth - firstColumnWidth - 4);

            listView.Columns[0].Width = firstColumnWidth;
            listView.Columns[1].Width = secondColumnWidth;
        }

        private static string NormalizeSingleLine(string value)
        {
            return (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
        }

        private static Panel CreateSurfacePanel(Padding padding)
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = padding
            };
        }

        private static Button CreateActionButton(string text)
        {
            Button button = new Button
            {
                Text = text,
                Width = 132,
                Height = 40,
                Margin = new Padding(0, 0, 0, 12),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(248, 250, 253),
                ForeColor = Color.FromArgb(43, 55, 71),
                Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Bold, GraphicsUnit.Point)
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(220, 226, 234);
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(236, 243, 255);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(223, 234, 255);
            return button;
        }

        private static Button CreatePrimaryActionButton(string text)
        {
            Button button = CreateActionButton(text);
            button.BackColor = Color.FromArgb(39, 110, 241);
            button.ForeColor = Color.White;
            button.Width = 104;
            button.FlatAppearance.BorderColor = Color.FromArgb(39, 110, 241);
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(31, 96, 214);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(26, 84, 191);
            return button;
        }

        private static Button CreateDangerActionButton(string text)
        {
            Button button = CreateActionButton(text);
            button.BackColor = Color.FromArgb(255, 244, 244);
            button.ForeColor = Color.FromArgb(183, 52, 52);
            button.FlatAppearance.BorderColor = Color.FromArgb(244, 210, 210);
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 233, 233);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(255, 219, 219);
            return button;
        }

        private static DocumentBasicInfo CloneInfo(DocumentBasicInfo info)
        {
            DocumentBasicInfo safeInfo = info ?? new DocumentBasicInfo();
            return new DocumentBasicInfo
            {
                Fields = (safeInfo.Fields ?? new List<DocumentBasicInfoField>())
                    .Select(field => new DocumentBasicInfoField
                    {
                        Name = field?.Name ?? string.Empty,
                        Value = field?.Value ?? string.Empty
                    })
                    .ToList()
            };
        }
    }
}
