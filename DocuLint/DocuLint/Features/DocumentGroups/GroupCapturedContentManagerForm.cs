using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class GroupCapturedContentManagerForm : Form
    {
        private readonly List<DocumentGroupCapturedContentItem> workingItems;
        private readonly ListView listView;
        private readonly Label emptyLabel;
        private readonly WebBrowser previewBrowser;
        private readonly Label previewPlaceholder;
        private readonly Button renameButton;
        private readonly Button deleteButton;

        public GroupCapturedContentManagerForm(string groupName, IEnumerable<DocumentGroupCapturedContentItem> items)
        {
            workingItems = CloneItems(items);

            Text = $"抓取管理 - {groupName}";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            MaximizeBox = true;
            ShowInTaskbar = false;
            Width = 1080;
            Height = 760;
            Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(14)
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 44f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 56f));

            SplitContainer contentSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 560
            };

            Panel listPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10)
            };

            emptyLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = "当前文档组没有抓取内容",
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(120, 126, 136),
                Font = new Font("Microsoft YaHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point)
            };

            listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                MultiSelect = false
            };
            listView.Columns.Add("标题", 520, HorizontalAlignment.Left);
            listView.SelectedIndexChanged += (_, __) =>
            {
                RefreshSelectionState();
                RefreshPreview();
            };
            listView.Resize += (_, __) => ResizeColumns();

            listPanel.Controls.Add(listView);
            listPanel.Controls.Add(emptyLabel);

            Panel previewPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(10)
            };
            previewBrowser = new WebBrowser
            {
                Dock = DockStyle.Fill,
                ScriptErrorsSuppressed = true,
                AllowWebBrowserDrop = false,
                IsWebBrowserContextMenuEnabled = false,
                WebBrowserShortcutsEnabled = true
            };
            previewPlaceholder = new Label
            {
                Dock = DockStyle.Fill,
                Text = "请选择抓取项以预览原格式内容",
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.FromArgb(120, 126, 136),
                Font = new Font("Microsoft YaHei UI", 11F, FontStyle.Regular, GraphicsUnit.Point)
            };
            previewPanel.Controls.Add(previewBrowser);
            previewPanel.Controls.Add(previewPlaceholder);

            contentSplit.Panel1.Controls.Add(previewPanel);
            contentSplit.Panel2.Controls.Add(listPanel);

            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = false,
                Padding = new Padding(0, 2, 0, 0)
            };

            renameButton = CreateActionButton("修改标题");
            deleteButton = CreateActionButton("删除");
            renameButton.Click += (_, __) => RenameSelected();
            deleteButton.Click += (_, __) => DeleteSelected();
            actions.Controls.Add(renameButton);
            actions.Controls.Add(deleteButton);

            FlowLayoutPanel bottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };
            Button okButton = CreateActionButton("确定");
            Button cancelButton = CreateActionButton("取消");
            okButton.Click += (_, __) => { DialogResult = DialogResult.OK; Close(); };
            cancelButton.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            bottom.Controls.Add(okButton);
            bottom.Controls.Add(cancelButton);

            root.Controls.Add(actions, 0, 0);
            root.Controls.Add(contentSplit, 0, 1);
            root.Controls.Add(bottom, 0, 2);
            Controls.Add(root);

            RefreshList();
            RefreshSelectionState();
            RefreshPreview();
        }

        public List<DocumentGroupCapturedContentItem> BuildItems()
        {
            return CloneItems(workingItems);
        }

        private void RenameSelected()
        {
            int index = GetSelectedIndex();
            if (index < 0 || index >= workingItems.Count)
            {
                return;
            }

            DocumentGroupCapturedContentItem item = workingItems[index];
            using (TextPromptForm prompt = new TextPromptForm("修改抓取标题", "请输入新的标题：", item.Title))
            {
                if (prompt.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                string newTitle = (prompt.ResultText ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(newTitle))
                {
                    MessageBox.Show(this, "标题不能为空。", "抓取管理");
                    return;
                }

                item.Title = newTitle;
                item.UpdatedAt = DateTime.Now;
                RefreshList();
                SelectIndex(index);
                RefreshPreview();
            }
        }

        private void DeleteSelected()
        {
            int index = GetSelectedIndex();
            if (index < 0 || index >= workingItems.Count)
            {
                return;
            }

            workingItems.RemoveAt(index);
            RefreshList();
            if (workingItems.Count > 0)
            {
                SelectIndex(Math.Min(index, workingItems.Count - 1));
            }

            RefreshPreview();
        }

        private void RefreshList()
        {
            listView.BeginUpdate();
            try
            {
                listView.Items.Clear();
                foreach (DocumentGroupCapturedContentItem item in workingItems)
                {
                    ListViewItem row = new ListViewItem(string.IsNullOrWhiteSpace(item?.Title) ? "未命名抓取" : item.Title.Trim());
                    listView.Items.Add(row);
                }
            }
            finally
            {
                listView.EndUpdate();
            }

            listView.Visible = workingItems.Count > 0;
            emptyLabel.Visible = workingItems.Count == 0;
            ResizeColumns();
        }

        private void ResizeColumns()
        {
            if (listView.Columns.Count < 1)
            {
                return;
            }

            int width = Math.Max(360, listView.ClientSize.Width);
            listView.Columns[0].Width = Math.Max(220, width - 6);
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

        private void RefreshSelectionState()
        {
            bool hasSelection = GetSelectedIndex() >= 0;
            renameButton.Enabled = hasSelection;
            deleteButton.Enabled = hasSelection;
        }

        private void RefreshPreview()
        {
            DocumentGroupCapturedContentItem item = GetSelectedItem();
            if (item == null)
            {
                ShowPreviewPlaceholder("请选择抓取项以预览原格式内容");
                return;
            }

            if (CapturedContentPreviewRenderer.TryRenderHtml(item.ContentWordOpenXml, out string htmlPath) && !string.IsNullOrWhiteSpace(htmlPath))
            {
                previewPlaceholder.Visible = false;
                previewBrowser.Visible = true;
                previewBrowser.Navigate(htmlPath);
                return;
            }

            previewPlaceholder.Visible = false;
            previewBrowser.Visible = true;
            previewBrowser.DocumentText = BuildFallbackHtml(item);
        }

        private DocumentGroupCapturedContentItem GetSelectedItem()
        {
            int index = GetSelectedIndex();
            if (index < 0 || index >= workingItems.Count)
            {
                return null;
            }

            return workingItems[index];
        }

        private void ShowPreviewPlaceholder(string text)
        {
            previewBrowser.Visible = false;
            previewPlaceholder.Visible = true;
            previewPlaceholder.Text = text;
        }

        private static string BuildFallbackHtml(DocumentGroupCapturedContentItem item)
        {
            string title = EscapeHtml(item?.Title ?? string.Empty);
            string text = EscapeHtml(item?.PreviewText ?? string.Empty);
            return "<html><head><meta charset='utf-8'></head><body style='font-family:\"Microsoft YaHei UI\";padding:12px;'>" +
                "<div style='font-size:14px;font-weight:700;margin-bottom:10px;'>" + title + "</div>" +
                "<div style='font-size:13px;color:#444;white-space:pre-wrap;line-height:1.6;'>" + text + "</div>" +
                "</body></html>";
        }

        private static string EscapeHtml(string text)
        {
            return (text ?? string.Empty)
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");
        }

        private static Button CreateActionButton(string text)
        {
            return new Button
            {
                Text = text,
                Width = 120,
                Height = 34,
                Margin = new Padding(0, 0, 10, 0)
            };
        }

        private static List<DocumentGroupCapturedContentItem> CloneItems(IEnumerable<DocumentGroupCapturedContentItem> items)
        {
            return (items ?? Enumerable.Empty<DocumentGroupCapturedContentItem>())
                .Where(item => item != null)
                .Select(item => new DocumentGroupCapturedContentItem
                {
                    Id = string.IsNullOrWhiteSpace(item.Id) ? Guid.NewGuid().ToString("N") : item.Id,
                    Title = item.Title ?? string.Empty,
                    PreviewText = item.PreviewText ?? string.Empty,
                    ContentWordOpenXml = item.ContentWordOpenXml ?? string.Empty,
                    CreatedAt = item.CreatedAt,
                    UpdatedAt = item.UpdatedAt
                })
                .ToList();
        }
    }
}
