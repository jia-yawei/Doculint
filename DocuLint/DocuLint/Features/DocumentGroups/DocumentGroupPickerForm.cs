using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DocuLint
{
    internal sealed class DocumentGroupPickerForm : Form
    {
        private readonly TreeView groupTreeView;
        private readonly ImageList imageList;

        public string SelectedGroupId => (groupTreeView.SelectedNode?.Tag as DocumentGroupItem)?.Id;

        public DocumentGroupPickerForm(
            IEnumerable<DocumentGroupItem> groups,
            string documentName,
            string activeGroupId = null,
            string title = "选择文档组",
            string instructionPrefix = "选择一个文档组，将当前文档加入该组：",
            string confirmButtonText = "加入选中组")
        {
            Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point, 134);
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(500, 360);

            Label tipLabel = new Label
            {
                Left = 16,
                Top = 16,
                Width = 468,
                Height = 44,
                Text = instructionPrefix + "\r\n" + (documentName ?? "当前文档")
            };

            imageList = new ImageList
            {
                ColorDepth = ColorDepth.Depth32Bit,
                ImageSize = new Size(16, 16)
            };
            imageList.Images.Add("folder", ShellIconHelper.GetSmallFolderBitmap());

            groupTreeView = new TreeView
            {
                Left = 16,
                Top = 72,
                Width = 468,
                Height = 220,
                HideSelection = false,
                FullRowSelect = true,
                ImageList = imageList
            };

            TreeNode activeNode = null;
            foreach (DocumentGroupItem group in groups ?? Enumerable.Empty<DocumentGroupItem>())
            {
                bool isActive = !string.IsNullOrWhiteSpace(activeGroupId) &&
                    string.Equals(group?.Id, activeGroupId, StringComparison.OrdinalIgnoreCase);
                TreeNode node = new TreeNode(isActive ? $"{group.Name}  [活动]" : group.Name)
                {
                    Tag = group,
                    ImageKey = "folder",
                    SelectedImageKey = "folder"
                };
                groupTreeView.Nodes.Add(node);
                if (isActive)
                {
                    activeNode = node;
                }
            }

            if (groupTreeView.Nodes.Count > 0)
            {
                groupTreeView.SelectedNode = activeNode ?? groupTreeView.Nodes[0];
            }

            Button okButton = new Button
            {
                Text = confirmButtonText,
                DialogResult = DialogResult.OK,
                Left = 296,
                Top = 310,
                Width = 90,
                Height = 32
            };

            Button cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Left = 394,
                Top = 310,
                Width = 90,
                Height = 32
            };

            Controls.Add(tipLabel);
            Controls.Add(groupTreeView);
            Controls.Add(okButton);
            Controls.Add(cancelButton);

            AcceptButton = okButton;
            CancelButton = cancelButton;
        }

        private static class ShellIconHelper
        {
            public static Bitmap GetSmallFolderBitmap()
            {
                NativeMethods.SHFILEINFO info = new NativeMethods.SHFILEINFO();
                IntPtr result = NativeMethods.SHGetFileInfo(
                    @"C:\",
                    NativeMethods.FileAttributeDirectory,
                    ref info,
                    (uint)Marshal.SizeOf(info),
                    NativeMethods.ShgfiIcon | NativeMethods.ShgfiSmallIcon | NativeMethods.ShgfiUseFileAttributes);

                if (result == IntPtr.Zero || info.hIcon == IntPtr.Zero)
                {
                    return SystemIcons.Application.ToBitmap();
                }

                try
                {
                    using (Icon icon = Icon.FromHandle(info.hIcon))
                    {
                        return icon.ToBitmap();
                    }
                }
                finally
                {
                    NativeMethods.DestroyIcon(info.hIcon);
                }
            }
        }

        private static class NativeMethods
        {
            public const uint ShgfiIcon = 0x000000100;
            public const uint ShgfiSmallIcon = 0x000000001;
            public const uint ShgfiUseFileAttributes = 0x000000010;
            public const uint FileAttributeDirectory = 0x00000010;

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
            public struct SHFILEINFO
            {
                public IntPtr hIcon;
                public int iIcon;
                public uint dwAttributes;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
                public string szDisplayName;
                [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
                public string szTypeName;
            }

            [DllImport("shell32.dll", CharSet = CharSet.Auto)]
            public static extern IntPtr SHGetFileInfo(
                string pszPath,
                uint dwFileAttributes,
                ref SHFILEINFO psfi,
                uint cbFileInfo,
                uint uFlags);

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool DestroyIcon(IntPtr hIcon);
        }
    }
}
