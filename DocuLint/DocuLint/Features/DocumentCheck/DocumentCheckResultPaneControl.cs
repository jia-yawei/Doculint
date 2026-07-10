using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class DocumentCheckResultPaneControl : UserControl
    {
        private static readonly Font PaneFont = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Regular, GraphicsUnit.Point);
        private static readonly Font HeaderFont = new Font("Microsoft YaHei UI", 10f, FontStyle.Bold, GraphicsUnit.Point);

        private readonly Label headerLabel;
        private readonly ListBox resultListBox;
        private readonly List<NavigationPaneEntry> entries = new List<NavigationPaneEntry>();

        internal event Action<int> IssueActivated;

        internal DocumentCheckResultPaneControl()
        {
            Dock = DockStyle.Fill;
            Font = PaneFont;
            BackColor = Color.White;

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10),
                BackColor = Color.White
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40f));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));

            headerLabel = new Label
            {
                Dock = DockStyle.Fill,
                AutoEllipsis = true,
                Font = HeaderFont,
                ForeColor = Color.FromArgb(48, 48, 48),
                TextAlign = ContentAlignment.MiddleLeft
            };

            resultListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawVariable,
                Font = PaneFont,
                IntegralHeight = false,
                ItemHeight = 38
            };
            resultListBox.MeasureItem += resultListBox_MeasureItem;
            resultListBox.DrawItem += resultListBox_DrawItem;
            resultListBox.DoubleClick += (_, __) => ActivateSelected();
            resultListBox.KeyDown += (sender, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    ActivateSelected();
                }
            };

            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                BackColor = Color.White,
                Padding = new Padding(0, 6, 0, 0)
            };
            buttons.Controls.Add(CreateButton("下一个", (_, __) => MoveSelection(1)));
            buttons.Controls.Add(CreateButton("上一个", (_, __) => MoveSelection(-1)));

            layout.Controls.Add(headerLabel, 0, 0);
            layout.Controls.Add(resultListBox, 0, 1);
            layout.Controls.Add(buttons, 0, 2);
            Controls.Add(layout);
        }

        private void resultListBox_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= resultListBox.Items.Count)
            {
                e.ItemHeight = 38;
                return;
            }

            string text = Convert.ToString(resultListBox.Items[e.Index]) ?? string.Empty;
            int width = Math.Max(120, resultListBox.ClientSize.Width - 34);
            Size size = TextRenderer.MeasureText(text, PaneFont, new Size(width, int.MaxValue), TextFormatFlags.WordBreak);
            e.ItemHeight = Math.Max(38, size.Height + 16);
        }

        private void resultListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= resultListBox.Items.Count)
            {
                return;
            }

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color backColor = selected ? Color.FromArgb(229, 241, 251) : Color.White;
            Color foreColor = Color.FromArgb(32, 32, 32);
            using (SolidBrush brush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(brush, e.Bounds);
            }

            Rectangle marker = new Rectangle(e.Bounds.Left + 8, e.Bounds.Top + 13, 6, 6);
            using (SolidBrush brush = new SolidBrush(selected ? Color.FromArgb(0, 120, 215) : Color.FromArgb(160, 160, 160)))
            {
                e.Graphics.FillEllipse(brush, marker);
            }

            Rectangle textBounds = new Rectangle(e.Bounds.Left + 24, e.Bounds.Top + 8, e.Bounds.Width - 30, e.Bounds.Height - 12);
            TextRenderer.DrawText(
                e.Graphics,
                Convert.ToString(resultListBox.Items[e.Index]) ?? string.Empty,
                PaneFont,
                textBounds,
                foreColor,
                TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl);

            if (selected)
            {
                e.DrawFocusRectangle();
            }
        }

        internal void SetEntries(IList<NavigationPaneEntry> newEntries, string docName)
        {
            entries.Clear();
            if (newEntries != null)
            {
                entries.AddRange(newEntries.Where(item => item != null));
            }

            headerLabel.Text = entries.Count > 0
                ? $"发现 {entries.Count} 项问题"
                : "未发现问题";

            resultListBox.BeginUpdate();
            try
            {
                resultListBox.Items.Clear();
                foreach (NavigationPaneEntry entry in entries)
                {
                    resultListBox.Items.Add(entry.Text ?? string.Empty);
                }
            }
            finally
            {
                resultListBox.EndUpdate();
            }

            if (resultListBox.Items.Count > 0)
            {
                resultListBox.SelectedIndex = 0;
            }
        }

        private static Button CreateButton(string text, EventHandler onClick)
        {
            Button button = new Button
            {
                AutoSize = false,
                Width = 78,
                Height = 28,
                FlatStyle = FlatStyle.System,
                Text = text
            };
            button.Click += onClick;
            return button;
        }

        private void MoveSelection(int offset)
        {
            if (resultListBox.Items.Count == 0)
            {
                return;
            }

            int next = resultListBox.SelectedIndex < 0 ? 0 : resultListBox.SelectedIndex + offset;
            resultListBox.SelectedIndex = Math.Max(0, Math.Min(resultListBox.Items.Count - 1, next));
            ActivateSelected();
        }

        private void ActivateSelected()
        {
            int index = resultListBox.SelectedIndex;
            if (index < 0 || index >= entries.Count)
            {
                return;
            }

            IssueActivated?.Invoke(entries[index].Start);
        }
    }
}
