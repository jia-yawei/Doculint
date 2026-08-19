using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal sealed class CommonPhrasesPaneControl : UserControl
    {
        private readonly Func<Word.Application> applicationAccessor;
        private readonly Label statusLabel;
        private readonly ListBox phraseList;
        private readonly Button insertButton;
        private List<string> phrases = new List<string>();
        private int measuredListWidth = -1;

        internal CommonPhrasesPaneControl(Func<Word.Application> applicationAccessor)
        {
            this.applicationAccessor = applicationAccessor;
            Dock = DockStyle.Fill;
            BackColor = Color.White;
            Font = new Font("Microsoft YaHei UI", 9F);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(10)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            statusLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8),
                ForeColor = Color.FromArgb(82, 91, 104)
            };
            phraseList = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                BorderStyle = BorderStyle.FixedSingle,
                HorizontalScrollbar = false,
                DrawMode = DrawMode.OwnerDrawVariable,
                ItemHeight = 40
            };
            phraseList.MeasureItem += PhraseList_MeasureItem;
            phraseList.DrawItem += PhraseList_DrawItem;
            phraseList.Resize += (_, __) =>
            {
                RefreshItemMeasurements();
            };
            phraseList.DoubleClick += (_, __) => InsertSelectedPhrase();
            insertButton = new Button
            {
                Text = "插入",
                Dock = DockStyle.Right,
                Width = 88,
                Height = 32,
                Enabled = false,
                Margin = new Padding(0, 10, 0, 0)
            };
            insertButton.Click += (_, __) => InsertSelectedPhrase();

            layout.Controls.Add(statusLabel, 0, 0);
            layout.Controls.Add(phraseList, 0, 1);
            layout.Controls.Add(insertButton, 0, 2);
            Controls.Add(layout);
            ReloadPhrases();
        }

        internal void ReloadPhrases()
        {
            phrases = CommonPhraseLibrary.LoadConfiguredPhrases().ToList();
            phraseList.BeginUpdate();
            try
            {
                PopulatePhraseItems();
            }
            finally
            {
                phraseList.EndUpdate();
            }

            bool configured = !string.IsNullOrWhiteSpace(CommonPhraseLibrary.ConfiguredPath);
            statusLabel.Text = phrases.Count > 0
                ? "已加载 " + phrases.Count + " 条常用语"
                : configured ? "常用语库为空或无法读取" : "尚未加载常用语库";
            insertButton.Enabled = phrases.Count > 0;
        }

        private void RefreshItemMeasurements()
        {
            int currentWidth = phraseList.ClientSize.Width;
            if (currentWidth <= 0 || currentWidth == measuredListWidth || phraseList.Items.Count == 0)
            {
                return;
            }

            measuredListWidth = currentWidth;
            int selectedIndex = phraseList.SelectedIndex;
            phraseList.BeginUpdate();
            try
            {
                // ListBox caches variable item heights. Re-adding items makes it measure
                // them again after the task pane has received its actual width.
                PopulatePhraseItems();
                if (selectedIndex >= 0 && selectedIndex < phraseList.Items.Count)
                {
                    phraseList.SelectedIndex = selectedIndex;
                }
            }
            finally
            {
                phraseList.EndUpdate();
            }
        }

        private void PopulatePhraseItems()
        {
            phraseList.Items.Clear();
            foreach (string phrase in phrases)
            {
                phraseList.Items.Add(phrase);
            }
        }

        private void InsertSelectedPhrase()
        {
            string phrase = phraseList.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(phrase))
            {
                return;
            }

            try
            {
                Word.Selection selection = applicationAccessor?.Invoke()?.Selection;
                if (selection?.Range == null)
                {
                    MessageBox.Show(this, "请先将光标放在文档中的插入位置。", "常用语");
                    return;
                }

                selection.Range.Text = phrase;
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "插入常用语失败：\r\n" + ex.Message, "常用语", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void PhraseList_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= phraseList.Items.Count)
            {
                e.ItemHeight = 36;
                return;
            }

            string text = Convert.ToString(phraseList.Items[e.Index]) ?? string.Empty;
            int textWidth = Math.Max(80, phraseList.ClientSize.Width - 24);
            Size measured = TextRenderer.MeasureText(
                text,
                phraseList.Font,
                new Size(textWidth, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
            int lineHeight = TextRenderer.MeasureText("A", phraseList.Font, Size.Empty, TextFormatFlags.NoPadding).Height;
            // Keep exactly one blank text line between adjacent phrases.
            e.ItemHeight = Math.Max(lineHeight * 2 + 4, measured.Height + lineHeight + 4);
        }

        private void PhraseList_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= phraseList.Items.Count)
            {
                return;
            }

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color backColor = selected
                ? SystemColors.Highlight
                : phraseList.BackColor;
            Color foreColor = selected
                ? SystemColors.HighlightText
                : phraseList.ForeColor;
            using (SolidBrush backBrush = new SolidBrush(backColor))
            using (SolidBrush foreBrush = new SolidBrush(foreColor))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
                Rectangle textBounds = new Rectangle(
                    e.Bounds.Left + 8,
                    e.Bounds.Top + 6,
                    Math.Max(20, e.Bounds.Width - 16),
                    Math.Max(20, e.Bounds.Height - 12));
                e.Graphics.DrawString(
                    Convert.ToString(phraseList.Items[e.Index]) ?? string.Empty,
                    phraseList.Font,
                    foreBrush,
                    textBounds,
                    new StringFormat
                    {
                        Alignment = StringAlignment.Near,
                        LineAlignment = StringAlignment.Near,
                        Trimming = StringTrimming.Word,
                        FormatFlags = StringFormatFlags.LineLimit
                    });
            }

            e.DrawFocusRectangle();
        }
    }
}
