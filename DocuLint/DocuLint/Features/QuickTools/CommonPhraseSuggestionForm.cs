using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    // 类似代码编辑器补全列表的轻量候选窗体，不使用模态对话框打断文档编辑。
    internal sealed class CommonPhraseSuggestionForm : Form
    {
        private readonly Func<Word.Application> applicationAccessor;
        private readonly string documentKey;
        private readonly int replacementStart;
        private readonly int replacementEnd;
        private readonly ListBox suggestionList;
        private readonly ToolTip phraseToolTip;
        private bool closing;

        internal CommonPhraseSuggestionForm(
            Func<Word.Application> applicationAccessor,
            string documentKey,
            int replacementStart,
            int replacementEnd,
            IReadOnlyList<CommonPhraseLibrary.Suggestion> suggestions)
        {
            this.applicationAccessor = applicationAccessor;
            this.documentKey = documentKey ?? string.Empty;
            this.replacementStart = replacementStart;
            this.replacementEnd = replacementEnd;

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Microsoft YaHei UI", 9F);
            BackColor = Color.FromArgb(210, 214, 220);
            Padding = new Padding(1);
            TopMost = true;

            suggestionList = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = Math.Max(26, TextRenderer.MeasureText("常用语", Font).Height + 10),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(35, 39, 44),
                HorizontalScrollbar = false,
                SelectionMode = SelectionMode.One
            };
            suggestionList.DrawItem += DrawSuggestion;
            suggestionList.SelectedIndexChanged += (_, __) => UpdateToolTip();
            suggestionList.DoubleClick += (_, __) => AcceptSelected();
            suggestionList.KeyDown += SuggestionList_KeyDown;
            foreach (CommonPhraseLibrary.Suggestion suggestion in suggestions ?? new List<CommonPhraseLibrary.Suggestion>())
            {
                suggestionList.Items.Add(suggestion);
            }

            phraseToolTip = new ToolTip
            {
                InitialDelay = 350,
                AutoPopDelay = 8000,
                ReshowDelay = 100
            };

            Controls.Add(suggestionList);
            Shown += (_, __) =>
            {
                suggestionList.SelectedIndex = suggestionList.Items.Count > 0 ? 0 : -1;
                suggestionList.Focus();
                UpdateToolTip();
            };
            Deactivate += (_, __) => HideSuggestion();
        }

        internal void ShowAt(Word.Application application)
        {
            Point location = Cursor.Position;
            try
            {
                Word.Window window = application?.ActiveWindow;
                Word.Selection selection = application?.Selection;
                if (window != null && selection?.Range != null)
                {
                    window.GetPoint(
                        out int left,
                        out int top,
                        out int width,
                        out int height,
                        selection.Range);
                    location = new Point(left, top + Math.Max(height, 18) + 2);
                }
            }
            catch
            {
                // Word may not expose a screen point while repaginating; use the mouse as fallback.
            }

            int desiredWidth = 260;
            using (Graphics graphics = CreateGraphics())
            {
                foreach (CommonPhraseLibrary.Suggestion suggestion in suggestionList.Items.Cast<CommonPhraseLibrary.Suggestion>())
                {
                    desiredWidth = Math.Max(
                        desiredWidth,
                        Math.Min(520, (int)Math.Ceiling(graphics.MeasureString(suggestion.Phrase ?? string.Empty, Font).Width) + 28));
                }
            }

            Width = desiredWidth;
            Height = Math.Min(190, Math.Max(30, suggestionList.ItemHeight * suggestionList.Items.Count + 2));
            Rectangle workingArea = Screen.FromPoint(location).WorkingArea;
            if (location.X + Width > workingArea.Right)
            {
                location.X = Math.Max(workingArea.Left, workingArea.Right - Width);
            }

            if (location.Y + Height > workingArea.Bottom)
            {
                location.Y = Math.Max(workingArea.Top, location.Y - Height - 24);
            }

            Location = location;
            Show();
            Activate();
        }

        internal void HideSuggestion()
        {
            if (closing || IsDisposed)
            {
                return;
            }

            closing = true;
            try
            {
                Hide();
            }
            finally
            {
                closing = false;
            }
        }

        private void SuggestionList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                AcceptSelected();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                HideSuggestion();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void AcceptSelected()
        {
            CommonPhraseLibrary.Suggestion suggestion = suggestionList.SelectedItem as CommonPhraseLibrary.Suggestion;
            if (suggestion == null)
            {
                HideSuggestion();
                return;
            }

            try
            {
                Word.Application application = applicationAccessor?.Invoke();
                Word.Document document = application?.ActiveDocument;
                if (document == null || !string.Equals(GetDocumentKey(document), documentKey, StringComparison.OrdinalIgnoreCase))
                {
                    HideSuggestion();
                    return;
                }

                Word.Range replacement = document.Range(replacementStart, replacementEnd);
                replacement.Text = suggestion.Phrase;
                replacement.Select();
                document.Activate();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "补全常用语失败：\r\n" + ex.Message, "常用语", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                HideSuggestion();
            }
        }

        private void DrawSuggestion(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= suggestionList.Items.Count)
            {
                return;
            }

            CommonPhraseLibrary.Suggestion suggestion = suggestionList.Items[e.Index] as CommonPhraseLibrary.Suggestion;
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            using (Brush background = new SolidBrush(selected ? Color.FromArgb(222, 235, 252) : Color.White))
            using (Brush foreground = new SolidBrush(selected ? Color.FromArgb(20, 55, 95) : suggestionList.ForeColor))
            {
                e.Graphics.FillRectangle(background, e.Bounds);
                e.Graphics.DrawString(suggestion?.Phrase ?? string.Empty, Font, foreground, e.Bounds.Left + 10, e.Bounds.Top + 5);
            }

            if ((e.State & DrawItemState.Focus) == DrawItemState.Focus)
            {
                using (Pen pen = new Pen(Color.FromArgb(145, 185, 230)))
                {
                    Rectangle focusBounds = e.Bounds;
                    focusBounds.Width--;
                    focusBounds.Height--;
                    e.Graphics.DrawRectangle(pen, focusBounds);
                }
            }
        }

        private void UpdateToolTip()
        {
            CommonPhraseLibrary.Suggestion suggestion = suggestionList.SelectedItem as CommonPhraseLibrary.Suggestion;
            phraseToolTip.SetToolTip(suggestionList, suggestion?.Phrase ?? string.Empty);
        }

        private static string GetDocumentKey(Word.Document document)
        {
            try
            {
                return !string.IsNullOrWhiteSpace(document?.FullName)
                    ? document.FullName
                    : document?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                phraseToolTip?.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
