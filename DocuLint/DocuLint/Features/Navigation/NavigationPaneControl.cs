using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal enum NavigationPaneTab
    {
        Bookmarks,
        Captions,
        Markers
    }

    internal sealed class NavigationPaneEntry
    {
        public int Start { get; set; }
        public string Text { get; set; }
    }

    internal sealed class NavigationPaneControl : UserControl
    {
        private readonly TabControl tabControl;
        private readonly Label lblBookmarkHeader;
        private readonly Label lblCaptionHeader;
        private readonly Label lblMarkerHeader;
        private readonly ListBox lstBookmarks;
        private readonly ListBox lstCaptions;
        private readonly ListBox lstMarkers;
        private readonly List<NavigationPaneEntry> bookmarkEntries = new List<NavigationPaneEntry>();
        private readonly List<NavigationPaneEntry> captionEntries = new List<NavigationPaneEntry>();
        private readonly List<NavigationPaneEntry> markerEntries = new List<NavigationPaneEntry>();

        internal event Action<int> BookmarkActivated;
        internal event Action<int> CaptionActivated;
        internal event Action<int> MarkerActivated;
        internal event Action<NavigationPaneTab> SelectedTabChanged;

        internal NavigationPaneControl()
        {
            Dock = DockStyle.Fill;

            tabControl = new TabControl
            {
                Dock = DockStyle.Fill
            };
            tabControl.SelectedIndexChanged += (_, __) => RaiseSelectedTabChanged();

            lblBookmarkHeader = CreateHeaderLabel("书签");
            lblCaptionHeader = CreateHeaderLabel("题注");
            lblMarkerHeader = CreateHeaderLabel("标识");
            lstBookmarks = CreateListBox();
            lstCaptions = CreateListBox();
            lstMarkers = CreateListBox();

            lstBookmarks.DoubleClick += (_, __) => ActivateBookmark();
            lstCaptions.DoubleClick += (_, __) => ActivateCaption();
            lstMarkers.DoubleClick += (_, __) => ActivateMarker();
            lstBookmarks.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    ActivateBookmark();
                }
            };
            lstCaptions.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    ActivateCaption();
                }
            };
            lstMarkers.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    ActivateMarker();
                }
            };

            tabControl.TabPages.Add(CreatePage("书签", lblBookmarkHeader, lstBookmarks));
            tabControl.TabPages.Add(CreatePage("题注", lblCaptionHeader, lstCaptions));
            tabControl.TabPages.Add(CreateMarkerPage());
            Controls.Add(tabControl);
        }

        internal void SetBookmarkEntries(IList<NavigationPaneEntry> entries, string docName)
        {
            SetEntries(bookmarkEntries, lstBookmarks, lblBookmarkHeader, entries, docName, "书签");
        }

        internal void SetCaptionEntries(IList<NavigationPaneEntry> entries, string docName)
        {
            SetEntries(captionEntries, lstCaptions, lblCaptionHeader, entries, docName, "题注");
        }

        internal void SetMarkerEntries(IList<NavigationPaneEntry> entries, string docName, string docTypeName)
        {
            SetEntries(markerEntries, lstMarkers, lblMarkerHeader, entries, docName, $"{docTypeName} 标识");
        }

        internal void SelectTab(NavigationPaneTab tab)
        {
            switch (tab)
            {
                case NavigationPaneTab.Bookmarks:
                    tabControl.SelectedIndex = 0;
                    break;
                case NavigationPaneTab.Captions:
                    tabControl.SelectedIndex = 1;
                    break;
                case NavigationPaneTab.Markers:
                    tabControl.SelectedIndex = 2;
                    break;
            }
        }

        internal NavigationPaneTab GetSelectedTab()
        {
            switch (tabControl.SelectedIndex)
            {
                case 0:
                    return NavigationPaneTab.Bookmarks;
                case 1:
                    return NavigationPaneTab.Captions;
                default:
                    return NavigationPaneTab.Markers;
            }
        }

        private void ActivateBookmark()
        {
            int index = lstBookmarks.SelectedIndex;
            if (index < 0 || index >= bookmarkEntries.Count)
            {
                return;
            }

            BookmarkActivated?.Invoke(bookmarkEntries[index].Start);
        }

        private void ActivateCaption()
        {
            int index = lstCaptions.SelectedIndex;
            if (index < 0 || index >= captionEntries.Count)
            {
                return;
            }

            CaptionActivated?.Invoke(captionEntries[index].Start);
        }

        private void ActivateMarker()
        {
            int index = lstMarkers.SelectedIndex;
            if (index < 0 || index >= markerEntries.Count)
            {
                return;
            }

            int start = markerEntries[index].Start;
            if (start < 0)
            {
                return;
            }

            MarkerActivated?.Invoke(start);
        }

        private void RaiseSelectedTabChanged()
        {
            SelectedTabChanged?.Invoke(GetSelectedTab());
        }

        private static void SetEntries(
            List<NavigationPaneEntry> targetEntries,
            ListBox listBox,
            Label header,
            IList<NavigationPaneEntry> newEntries,
            string docName,
            string categoryName)
        {
            targetEntries.Clear();
            if (newEntries != null)
            {
                targetEntries.AddRange(newEntries.Where(item => item != null));
            }

            header.Text = $"{docName} - {categoryName} {targetEntries.Count} 项";

            listBox.BeginUpdate();
            try
            {
                listBox.Items.Clear();
                foreach (NavigationPaneEntry entry in targetEntries)
                {
                    listBox.Items.Add(entry.Text ?? string.Empty);
                }
            }
            finally
            {
                listBox.EndUpdate();
            }

            if (listBox.Items.Count > 0)
            {
                listBox.SelectedIndex = 0;
            }
        }

        private static TabPage CreatePage(string title, Label header, ListBox listBox)
        {
            TabPage page = new TabPage(title);
            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(8)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.Controls.Add(header, 0, 0);
            layout.Controls.Add(listBox, 0, 1);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage CreateMarkerPage()
        {
            TabPage page = CreatePage("标识", lblMarkerHeader, lstMarkers);
            return page;
        }

        private static Label CreateHeaderLabel(string text)
        {
            return new Label
            {
                Dock = DockStyle.Fill,
                Text = text,
                ForeColor = Color.FromArgb(80, 80, 80),
                TextAlign = ContentAlignment.MiddleLeft
            };
        }

        private static ListBox CreateListBox()
        {
            return new ListBox
            {
                Dock = DockStyle.Fill,
                HorizontalScrollbar = true,
                IntegralHeight = false
            };
        }
    }
}
