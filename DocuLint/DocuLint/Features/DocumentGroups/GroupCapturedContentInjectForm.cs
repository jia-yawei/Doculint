using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class GroupCapturedContentInjectForm : Form
    {
        private readonly List<DocumentGroupCapturedContentItem> items;
        private readonly ListView listView;
        private readonly WebBrowser previewBrowser;
        private readonly Label previewPlaceholder;
        public event Action<DocumentGroupCapturedContentItem> InjectRequested;

        public DocumentGroupCapturedContentItem SelectedItem { get; private set; }

        public GroupCapturedContentInjectForm(string groupName, IEnumerable<DocumentGroupCapturedContentItem> sourceItems)
        {
            items = (sourceItems ?? Enumerable.Empty<DocumentGroupCapturedContentItem>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.ContentWordOpenXml))
                .ToList();

            Text = $"内容注入 - {groupName}";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1080;
            Height = 760;
            MinimizeBox = false;
            MaximizeBox = true;
            ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point);

            SplitContainer split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 380
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
            listView.SelectedIndexChanged += (_, __) => RefreshPreview();
            listView.DoubleClick += (_, __) => ConfirmSelection();
            listView.Resize += (_, __) =>
            {
                if (listView.Columns.Count > 0)
                {
                    listView.Columns[0].Width = Math.Max(220, listView.ClientSize.Width - 6);
                }
            };

            Panel previewPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(8)
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

            FlowLayoutPanel bottomButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                Height = 48,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };
            Button injectButton = new Button { Text = "注入", Width = 90, Height = 30, Margin = new Padding(6, 8, 6, 8) };
            Button cancelButton = new Button { Text = "取消", Width = 90, Height = 30, Margin = new Padding(6, 8, 6, 8) };
            injectButton.Click += (_, __) => ConfirmSelection();
            cancelButton.Click += (_, __) => { DialogResult = DialogResult.Cancel; Close(); };
            bottomButtons.Controls.Add(injectButton);
            bottomButtons.Controls.Add(cancelButton);

            split.Panel1.Controls.Add(previewPanel);
            split.Panel2.Controls.Add(listView);

            Controls.Add(split);
            Controls.Add(bottomButtons);

            LoadItems();
            RefreshPreview();
        }

        private void LoadItems()
        {
            listView.BeginUpdate();
            try
            {
                listView.Items.Clear();
                foreach (DocumentGroupCapturedContentItem item in items)
                {
                    ListViewItem row = new ListViewItem(string.IsNullOrWhiteSpace(item.Title) ? "未命名抓取" : item.Title.Trim());
                    row.Tag = item;
                    listView.Items.Add(row);
                }
            }
            finally
            {
                listView.EndUpdate();
            }

            if (listView.Items.Count > 0)
            {
                listView.Items[0].Selected = true;
            }

            if (listView.Columns.Count > 0)
            {
                listView.Columns[0].Width = Math.Max(220, listView.ClientSize.Width - 6);
            }
        }

        private void RefreshPreview()
        {
            DocumentGroupCapturedContentItem item = GetSelected();
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

        private void ConfirmSelection()
        {
            DocumentGroupCapturedContentItem item = GetSelected();
            if (item == null)
            {
                return;
            }

            SelectedItem = item;
            InjectRequested?.Invoke(item);
        }

        private DocumentGroupCapturedContentItem GetSelected()
        {
            if (listView.SelectedItems.Count <= 0)
            {
                return null;
            }

            return listView.SelectedItems[0].Tag as DocumentGroupCapturedContentItem;
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
    }
}
