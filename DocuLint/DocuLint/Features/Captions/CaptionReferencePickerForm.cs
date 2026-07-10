using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class CaptionReferencePickerForm : Form
    {
        private readonly ListBox captionListBox;
        private readonly ComboBox referenceKindComboBox;

        public int SelectedIndex => captionListBox.SelectedIndex;

        public Ribbon1.CaptionReferenceKind SelectedReferenceKind
        {
            get
            {
                switch (referenceKindComboBox.SelectedIndex)
                {
                    case 1:
                        return Ribbon1.CaptionReferenceKind.FullCaption;
                    case 2:
                        return Ribbon1.CaptionReferenceKind.PageNumber;
                    default:
                        return Ribbon1.CaptionReferenceKind.Number;
                }
            }
        }

        public CaptionReferencePickerForm(IList<string> captions, Ribbon1.CaptionReferenceKind defaultKind)
        {
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            Text = "引用题注";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(640, 430);

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 4,
                Padding = new Padding(14)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));

            Label typeLabel = new Label
            {
                Dock = DockStyle.Fill,
                Text = "引用类型",
                TextAlign = ContentAlignment.MiddleLeft
            };

            referenceKindComboBox = new ComboBox
            {
                Dock = DockStyle.Top,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            referenceKindComboBox.Items.AddRange(new object[] { "标签和编号", "完整题注", "页码" });
            referenceKindComboBox.SelectedIndex = defaultKind == Ribbon1.CaptionReferenceKind.FullCaption
                ? 1
                : (defaultKind == Ribbon1.CaptionReferenceKind.PageNumber ? 2 : 0);

            captionListBox = new ListBox
            {
                Dock = DockStyle.Fill,
                HorizontalScrollbar = true,
                IntegralHeight = false
            };
            if (captions != null)
            {
                foreach (string caption in captions)
                {
                    captionListBox.Items.Add(caption ?? string.Empty);
                }
            }

            if (captionListBox.Items.Count > 0)
            {
                captionListBox.SelectedIndex = 0;
            }

            captionListBox.DoubleClick += (_, __) => ConfirmSelection();
            captionListBox.KeyDown += CaptionListBox_KeyDown;

            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false
            };

            Button insertButton = new Button
            {
                Text = "插入",
                Width = 84,
                Height = 30
            };
            insertButton.Click += (_, __) => ConfirmSelection();

            Button cancelButton = new Button
            {
                Text = "取消",
                Width = 84,
                Height = 30
            };
            cancelButton.Click += (_, __) =>
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };

            buttonPanel.Controls.Add(insertButton);
            buttonPanel.Controls.Add(cancelButton);

            layout.Controls.Add(typeLabel, 0, 0);
            layout.Controls.Add(referenceKindComboBox, 0, 1);
            layout.Controls.Add(captionListBox, 0, 2);
            layout.Controls.Add(buttonPanel, 0, 3);
            Controls.Add(layout);
        }

        private void CaptionListBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
            {
                return;
            }

            e.Handled = true;
            e.SuppressKeyPress = true;
            ConfirmSelection();
        }

        private void ConfirmSelection()
        {
            if (captionListBox.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择一个题注。", "文档不加班");
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
