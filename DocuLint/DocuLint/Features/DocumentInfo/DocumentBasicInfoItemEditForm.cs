using System;
using System.Drawing;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class DocumentBasicInfoItemEditForm : Form
    {
        private readonly TextBox txtName;
        private readonly TextBox txtValue;

        internal DocumentBasicInfoItemEditForm(string title, DocumentBasicInfoField field)
        {
            DocumentBasicInfoField safeField = field ?? new DocumentBasicInfoField();

            Text = title;
            Font = new Font("Microsoft YaHei UI", 10F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(620, 360);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(18)
            };
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 24f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 42f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 24f));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 52f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 1f));

            Label lblName = new Label
            {
                Dock = DockStyle.Fill,
                Text = "基本信息类型",
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(72, 80, 92)
            };
            txtName = new TextBox
            {
                Dock = DockStyle.Fill,
                Text = safeField.Name ?? string.Empty
            };

            Label lblValue = new Label
            {
                Dock = DockStyle.Fill,
                Text = "基本信息内容",
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(72, 80, 92)
            };
            txtValue = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = safeField.Value ?? string.Empty
            };

            FlowLayoutPanel buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                Padding = new Padding(0, 10, 0, 0)
            };

            Button btnOk = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Width = 96,
                Height = 34
            };
            Button btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Width = 96,
                Height = 34
            };

            buttonPanel.Controls.Add(btnOk);
            buttonPanel.Controls.Add(btnCancel);

            root.Controls.Add(lblName, 0, 0);
            root.Controls.Add(txtName, 0, 1);
            root.Controls.Add(lblValue, 0, 2);
            root.Controls.Add(txtValue, 0, 3);
            root.Controls.Add(buttonPanel, 0, 4);
            Controls.Add(root);

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        internal bool TryBuildField(out DocumentBasicInfoField field)
        {
            string name = (txtName.Text ?? string.Empty).Trim();
            string value = (txtValue.Text ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(name))
            {
                MessageBox.Show(this, "请先填写基本信息类型。", "文档不加班");
                txtName.Focus();
                field = null;
                return false;
            }

            field = new DocumentBasicInfoField
            {
                Name = name,
                Value = value
            };
            return true;
        }
    }
}
