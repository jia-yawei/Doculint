using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class CaptionListEntry
    {
        public int Start { get; set; }
        public string Text { get; set; }
    }

    internal sealed class CaptionListPaneControl : UserControl
    {
        private static readonly Font NavigationLikeFont = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);

        private readonly Label lblHeader;
        private readonly ListBox lstCaptions;
        private readonly List<CaptionListEntry> entries = new List<CaptionListEntry>();

        internal event Action<int> CaptionActivated;

        internal CaptionListPaneControl()
        {
            Dock = DockStyle.Fill;

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(8)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28f));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            lblHeader = new Label
            {
                Dock = DockStyle.Fill,
                Font = NavigationLikeFont,
                ForeColor = Color.FromArgb(80, 80, 80),
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "题注列表"
            };

            lstCaptions = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = NavigationLikeFont,
                DrawMode = DrawMode.OwnerDrawFixed,
                HorizontalScrollbar = true,
                IntegralHeight = false,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(34, 34, 34)
            };
            lstCaptions.ItemHeight = CalculateItemHeight(lstCaptions.Font);
            lstCaptions.DoubleClick += (_, __) => ActivateSelected();
            lstCaptions.KeyDown += LstCaptions_KeyDown;
            lstCaptions.DrawItem += LstCaptions_DrawItem;

            layout.Controls.Add(lblHeader, 0, 0);
            layout.Controls.Add(lstCaptions, 0, 1);

            Controls.Add(layout);
        }

        internal void SetEntries(IList<CaptionListEntry> newEntries, string docName)
        {
            entries.Clear();
            if (newEntries != null)
            {
                entries.AddRange(newEntries.Where(e => e != null));
            }

            lblHeader.Text = $"{docName} - 题注 {entries.Count} 项";

            lstCaptions.BeginUpdate();
            try
            {
                lstCaptions.Items.Clear();
                foreach (CaptionListEntry entry in entries)
                {
                    lstCaptions.Items.Add(entry.Text ?? string.Empty);
                }
            }
            finally
            {
                lstCaptions.EndUpdate();
            }

            if (lstCaptions.Items.Count > 0)
            {
                lstCaptions.SelectedIndex = 0;
            }
        }

        private void LstCaptions_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
            {
                return;
            }

            e.Handled = true;
            e.SuppressKeyPress = true;
            ActivateSelected();
        }

        private static int CalculateItemHeight(Font font)
        {
            if (font == null)
            {
                return 20;
            }

            return Math.Max(18, (int)Math.Ceiling(font.GetHeight() + 6f));
        }

        private void LstCaptions_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();

            if (e.Index < 0 || e.Index >= lstCaptions.Items.Count)
            {
                return;
            }

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            string text = Convert.ToString(lstCaptions.Items[e.Index]) ?? string.Empty;
            Rectangle textBounds = new Rectangle(e.Bounds.Left + 8, e.Bounds.Top, Math.Max(0, e.Bounds.Width - 12), e.Bounds.Height);
            Color foreColor = selected ? SystemColors.HighlightText : lstCaptions.ForeColor;

            TextRenderer.DrawText(
                e.Graphics,
                text,
                lstCaptions.Font,
                textBounds,
                foreColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);

            e.DrawFocusRectangle();
        }

        private void ActivateSelected()
        {
            int index = lstCaptions.SelectedIndex;
            if (index < 0 || index >= entries.Count)
            {
                return;
            }

            CaptionListEntry entry = entries[index];
            CaptionActivated?.Invoke(entry.Start);
        }
    }
}

