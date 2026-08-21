using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal sealed class DocumentVersionEntry
    {
        internal int Number { get; set; }
        internal DateTime CreatedAt { get; set; }
        internal string SnapshotPath { get; set; }
        internal string ArchiveFilePath { get; set; }
        internal string ArchiveEntryName { get; set; }
        internal string Note { get; set; }

        internal string DisplayName => $"V{Number:000}  {CreatedAt:yyyy-MM-dd HH:mm:ss}";
    }

    internal static class DocumentVersionArchive
    {
        private const string ManifestFileName = "manifest.xml";

        internal static string GetArchiveDirectory(Word.Document document)
        {
            string fullName = document?.FullName;
            if (string.IsNullOrWhiteSpace(fullName) || !File.Exists(fullName))
            {
                return string.Empty;
            }

            string directory = Path.GetDirectoryName(fullName);
            return Path.Combine(directory, ".doculint-versions");
        }

        private static string GetArchiveFilePath(Word.Document document)
        {
            string fullName = document?.FullName;
            string archiveDirectory = GetArchiveDirectory(document);
            if (string.IsNullOrWhiteSpace(fullName) || string.IsNullOrWhiteSpace(archiveDirectory))
            {
                return string.Empty;
            }

            // Keep one archive file per source document inside the single hidden folder.
            return Path.Combine(archiveDirectory, Path.GetFileName(fullName) + ".versions.zip");
        }

        private static string GetLegacyArchiveDirectory(Word.Document document)
        {
            string fullName = document?.FullName;
            if (string.IsNullOrWhiteSpace(fullName) || !File.Exists(fullName))
            {
                return string.Empty;
            }

            return Path.Combine(
                Path.GetDirectoryName(fullName),
                Path.GetFileNameWithoutExtension(fullName) + ".doculint-versions");
        }

        internal static List<DocumentVersionEntry> Load(Word.Document document)
        {
            string archiveFilePath = GetArchiveFilePath(document);
            if (File.Exists(archiveFilePath))
            {
                return LoadZipArchive(archiveFilePath);
            }

            // Read archives created by the earlier per-document-folder format.
            string legacyDirectory = GetLegacyArchiveDirectory(document);
            string manifestPath = string.IsNullOrWhiteSpace(legacyDirectory)
                ? string.Empty
                : Path.Combine(legacyDirectory, ManifestFileName);
            if (!File.Exists(manifestPath))
            {
                return new List<DocumentVersionEntry>();
            }

            try
            {
                XDocument xml = XDocument.Load(manifestPath);
                return xml.Root?.Elements("version")
                    .Select(element => new DocumentVersionEntry
                    {
                        Number = ParseInt((string)element.Attribute("number")),
                        CreatedAt = ParseDate((string)element.Attribute("createdAt")),
                        SnapshotPath = Path.Combine(legacyDirectory, (string)element.Attribute("file") ?? string.Empty),
                        ArchiveEntryName = (string)element.Attribute("file") ?? string.Empty,
                        Note = (string)element.Attribute("note") ?? string.Empty
                    })
                    .Where(entry => entry.Number > 0 && File.Exists(entry.SnapshotPath))
                    .OrderByDescending(entry => entry.Number)
                    .ToList() ?? new List<DocumentVersionEntry>();
            }
            catch
            {
                return new List<DocumentVersionEntry>();
            }
        }

        internal static DocumentVersionEntry Archive(Word.Document document, string note)
        {
            string fullName = document?.FullName;
            if (string.IsNullOrWhiteSpace(fullName) || !File.Exists(fullName))
            {
                throw new InvalidOperationException("当前文档尚未保存，请先保存文档后再存档。");
            }

            string archiveDirectory = GetArchiveDirectory(document);
            Directory.CreateDirectory(archiveDirectory);
            SetHidden(archiveDirectory);
            string archiveFilePath = GetArchiveFilePath(document);
            List<DocumentVersionEntry> entries = Load(document);

            // On first use, Word has not saved the user's current edits yet. Preserve the
            // on-disk document as the baseline before saving the edited current document.
            if (entries.Count == 0 && !document.Saved)
            {
                DateTime baselineTime = File.GetLastWriteTime(fullName);
                DocumentVersionEntry baseline = CreateSnapshotEntry(
                    fullName,
                    archiveFilePath,
                    1,
                    baselineTime,
                    "首次存档前版本（自动创建）");
                entries.Add(baseline);
                SaveArchive(archiveFilePath, entries);
            }

            document.Save();
            int nextNumber = entries.Count == 0 ? 1 : entries.Max(item => item.Number) + 1;
            DateTime createdAt = DateTime.Now;
            DocumentVersionEntry entry = CreateSnapshotEntry(
                fullName,
                archiveFilePath,
                nextNumber,
                createdAt,
                (note ?? string.Empty).Trim());
            entries.Add(entry);
            SaveArchive(archiveFilePath, entries);
            return entry;
        }

        internal static void Delete(Word.Document document, DocumentVersionEntry entry)
        {
            if (entry == null)
            {
                return;
            }

            string archiveFilePath = GetArchiveFilePath(document);
            List<DocumentVersionEntry> remaining = Load(document)
                .Where(item => item.Number != entry.Number)
                .ToList();
            if (remaining.Count == 0)
            {
                if (File.Exists(archiveFilePath))
                {
                    File.Delete(archiveFilePath);
                }
                return;
            }

            SaveArchive(archiveFilePath, remaining);
        }

        internal static Word.Document OpenSnapshot(Word.Application application, DocumentVersionEntry entry)
        {
            if (application == null || entry == null || !File.Exists(entry.SnapshotPath))
            {
                return null;
            }

            return application.Documents.Open(
                FileName: entry.SnapshotPath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: true);
        }

        internal static Word.Document CompareWithCurrent(
            Word.Application application,
            Word.Document currentDocument,
            DocumentVersionEntry entry)
        {
            if (application == null || currentDocument == null || entry == null || !File.Exists(entry.SnapshotPath))
            {
                return null;
            }

            Word.Document archivedDocument = null;
            try
            {
                archivedDocument = application.Documents.Open(
                    FileName: entry.SnapshotPath,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false);
                return application.CompareDocuments(
                    OriginalDocument: archivedDocument,
                    RevisedDocument: currentDocument,
                    Destination: Word.WdCompareDestination.wdCompareDestinationNew,
                    CompareFormatting: true,
                    CompareCaseChanges: true,
                    CompareWhitespace: true,
                    CompareTables: true,
                    CompareHeaders: true,
                    CompareFootnotes: true,
                    CompareTextboxes: true,
                    CompareFields: true,
                    CompareComments: true,
                    CompareMoves: true,
                    RevisedAuthor: "搞快点");
            }
            finally
            {
                try
                {
                    archivedDocument?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch
                {
                }
            }
        }

        internal static Word.Document CompareSnapshots(
            Word.Application application,
            DocumentVersionEntry originalEntry,
            DocumentVersionEntry revisedEntry)
        {
            if (application == null || originalEntry == null || revisedEntry == null ||
                !File.Exists(originalEntry.SnapshotPath) || !File.Exists(revisedEntry.SnapshotPath))
            {
                return null;
            }

            Word.Document originalDocument = null;
            Word.Document revisedDocument = null;
            try
            {
                originalDocument = application.Documents.Open(
                    FileName: originalEntry.SnapshotPath,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false);
                revisedDocument = application.Documents.Open(
                    FileName: revisedEntry.SnapshotPath,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false);
                return application.CompareDocuments(
                    OriginalDocument: originalDocument,
                    RevisedDocument: revisedDocument,
                    Destination: Word.WdCompareDestination.wdCompareDestinationNew,
                    CompareFormatting: true,
                    CompareCaseChanges: true,
                    CompareWhitespace: true,
                    CompareTables: true,
                    CompareHeaders: true,
                    CompareFootnotes: true,
                    CompareTextboxes: true,
                    CompareFields: true,
                    CompareComments: true,
                    CompareMoves: true,
                    RevisedAuthor: "搞快点");
            }
            finally
            {
                try
                {
                    originalDocument?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    revisedDocument?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch
                {
                }
            }
        }

        internal static Word.Document RestoreSnapshot(
            Word.Application application,
            Word.Document currentDocument,
            DocumentVersionEntry entry)
        {
            if (application == null || currentDocument == null || entry == null ||
                string.IsNullOrWhiteSpace(currentDocument.FullName) || !File.Exists(entry.SnapshotPath))
            {
                return null;
            }

            string currentPath = currentDocument.FullName;
            currentDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            File.Copy(entry.SnapshotPath, currentPath, true);
            return application.Documents.Open(
                FileName: currentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);
        }

        private static List<DocumentVersionEntry> LoadZipArchive(string archiveFilePath)
        {
            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(archiveFilePath))
                {
                    ZipArchiveEntry manifestEntry = archive.GetEntry(ManifestFileName);
                    if (manifestEntry == null)
                    {
                        return new List<DocumentVersionEntry>();
                    }

                    XDocument xml;
                    using (Stream stream = manifestEntry.Open())
                    {
                        xml = XDocument.Load(stream);
                    }

                    return xml.Root?.Elements("version")
                        .Select(element =>
                        {
                            string entryName = (string)element.Attribute("file") ?? string.Empty;
                            ZipArchiveEntry snapshot = archive.GetEntry(entryName);
                            if (snapshot == null)
                            {
                                return null;
                            }

                            return new DocumentVersionEntry
                            {
                                Number = ParseInt((string)element.Attribute("number")),
                                CreatedAt = ParseDate((string)element.Attribute("createdAt")),
                                SnapshotPath = ExtractSnapshot(archiveFilePath, snapshot),
                                ArchiveFilePath = archiveFilePath,
                                ArchiveEntryName = entryName,
                                Note = (string)element.Attribute("note") ?? string.Empty
                            };
                        })
                        .Where(entry => entry != null && entry.Number > 0 && File.Exists(entry.SnapshotPath))
                        .OrderByDescending(entry => entry.Number)
                        .ToList() ?? new List<DocumentVersionEntry>();
                }
            }
            catch
            {
                return new List<DocumentVersionEntry>();
            }
        }

        private static DocumentVersionEntry CreateSnapshotEntry(
            string sourcePath,
            string archiveFilePath,
            int number,
            DateTime createdAt,
            string note)
        {
            string extension = Path.GetExtension(sourcePath);
            if (string.IsNullOrWhiteSpace(extension))
            {
                extension = ".docx";
            }

            string fileName = $"V{number:000}_{createdAt:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}{extension}";
            string snapshotPath = Path.Combine(Path.GetTempPath(), "DocuLint-VersionSnapshots", fileName);
            Directory.CreateDirectory(Path.GetDirectoryName(snapshotPath));
            File.Copy(sourcePath, snapshotPath, false);
            return new DocumentVersionEntry
            {
                Number = number,
                CreatedAt = createdAt,
                SnapshotPath = snapshotPath,
                ArchiveFilePath = archiveFilePath,
                ArchiveEntryName = $"V{number:000}_{createdAt:yyyyMMdd_HHmmss}{extension}",
                Note = note ?? string.Empty
            };
        }

        private static string ExtractSnapshot(string archiveFilePath, ZipArchiveEntry archiveEntry)
        {
            string extension = Path.GetExtension(archiveEntry.Name);
            string fileName = $"{Path.GetFileNameWithoutExtension(archiveEntry.Name)}_{Guid.NewGuid():N}{extension}";
            string snapshotPath = Path.Combine(Path.GetTempPath(), "DocuLint-VersionSnapshots", fileName);
            Directory.CreateDirectory(Path.GetDirectoryName(snapshotPath));
            using (Stream source = archiveEntry.Open())
            using (FileStream target = File.Create(snapshotPath))
            {
                source.CopyTo(target);
            }
            return snapshotPath;
        }

        private static void SaveArchive(string archiveFilePath, IEnumerable<DocumentVersionEntry> sourceEntries)
        {
            if (string.IsNullOrWhiteSpace(archiveFilePath))
            {
                return;
            }

            List<DocumentVersionEntry> entries = (sourceEntries ?? Enumerable.Empty<DocumentVersionEntry>())
                .OrderBy(entry => entry.Number)
                .ToList();
            string archiveDirectory = Path.GetDirectoryName(archiveFilePath);
            Directory.CreateDirectory(archiveDirectory);
            SetHidden(archiveDirectory);

            foreach (DocumentVersionEntry entry in entries)
            {
                if (string.IsNullOrWhiteSpace(entry.ArchiveEntryName))
                {
                    string extension = Path.GetExtension(entry.SnapshotPath);
                    entry.ArchiveEntryName = $"V{entry.Number:000}_{entry.CreatedAt:yyyyMMdd_HHmmss}{extension}";
                }
                entry.ArchiveFilePath = archiveFilePath;
            }

            string temporaryPath = archiveFilePath + ".tmp";
            if (File.Exists(temporaryPath))
            {
                File.Delete(temporaryPath);
            }

            using (FileStream stream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, false, Encoding.UTF8))
            {
                XElement root = new XElement("versions",
                    new XAttribute("updatedAt", DateTime.Now.ToString("o")),
                    entries.Select(entry => new XElement("version",
                        new XAttribute("number", entry.Number),
                        new XAttribute("createdAt", entry.CreatedAt.ToString("o")),
                        new XAttribute("file", entry.ArchiveEntryName),
                        new XAttribute("note", entry.Note ?? string.Empty))));
                ZipArchiveEntry manifest = archive.CreateEntry(ManifestFileName, CompressionLevel.Fastest);
                using (StreamWriter writer = new StreamWriter(manifest.Open(), new UTF8Encoding(false)))
                {
                    writer.Write(root.ToString(SaveOptions.DisableFormatting));
                }

                foreach (DocumentVersionEntry entry in entries)
                {
                    if (!File.Exists(entry.SnapshotPath))
                    {
                        throw new FileNotFoundException("版本快照文件不存在。", entry.SnapshotPath);
                    }

                    ZipArchiveEntry snapshot = archive.CreateEntry(entry.ArchiveEntryName, CompressionLevel.Fastest);
                    using (Stream source = File.OpenRead(entry.SnapshotPath))
                    using (Stream target = snapshot.Open())
                    {
                        source.CopyTo(target);
                    }
                }
            }

            File.Copy(temporaryPath, archiveFilePath, true);
            File.Delete(temporaryPath);
        }

        private static void SetHidden(string path)
        {
            try
            {
                if (Directory.Exists(path))
                {
                    File.SetAttributes(path, File.GetAttributes(path) | FileAttributes.Hidden);
                }
            }
            catch
            {
                // Hidden attribute is cosmetic and may be unavailable on some file systems.
            }
        }

        private static int ParseInt(string value)
        {
            return int.TryParse(value, out int result) ? result : 0;
        }

        private static DateTime ParseDate(string value)
        {
            return DateTime.TryParse(value, out DateTime result) ? result : DateTime.MinValue;
        }
    }

    internal sealed class DocumentVersionManagementForm : Form
    {
        private readonly Word.Application application;
        private readonly Word.Document document;
        private readonly ListView versionList;
        private readonly TextBox noteTextBox;
        private readonly Label statusLabel;
        private List<DocumentVersionEntry> entries = new List<DocumentVersionEntry>();

        internal DocumentVersionManagementForm(Word.Application application, Word.Document document)
        {
            this.application = application;
            this.document = document;
            Text = "文档版本管理";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            StartPosition = FormStartPosition.CenterScreen;
            ShowInTaskbar = false;
            ClientSize = new Size(860, 520);
            MinimumSize = new Size(720, 440);
            BackColor = Color.White;

            TableLayoutPanel layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(16)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            Panel hintPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 38,
                BackColor = Color.FromArgb(240, 248, 251),
                Margin = new Padding(0, 0, 0, 10)
            };
            Label hint = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 10, 0),
                Text = "当前文档为最新版。历史版本集中保存在文档同目录、已隐藏的 .doculint-versions 文件夹中。"
            };
            hintPanel.Controls.Add(hint);
            TableLayoutPanel noteRow = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                RowCount = 1,
                Height = 34,
                Margin = new Padding(0, 0, 0, 10)
            };
            // Reserve enough width for all four Chinese characters even on high-DPI hosts.
            noteRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 128f));
            noteRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            noteRow.Controls.Add(new Label
            {
                Text = "存档备注",
                AutoSize = false,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(0, 0, 8, 0),
                MinimumSize = new Size(128, 0)
            }, 0, 0);
            noteTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 3, 0, 3)
            };
            noteRow.Controls.Add(noteTextBox, 1, 0);

            versionList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                HideSelection = false,
                MultiSelect = false
            };
            versionList.Columns.Add("版本", 100);
            versionList.Columns.Add("存档时间", 180);
            versionList.Columns.Add("备注", 480);
            versionList.Resize += (_, __) => ResizeVersionColumns();
            versionList.DoubleClick += (_, __) => OpenSelectedSnapshot();

            FlowLayoutPanel buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 10, 0, 0),
                BackColor = Color.White
            };
            Button archiveButton = CreateButton("存档");
            archiveButton.Click += (_, __) => ArchiveCurrentDocument();
            Button compareButton = CreateButton("查看改动");
            compareButton.Click += (_, __) => CompareWithPreviousSnapshot();
            Button compareCurrentButton = CreateButton("与当前比较");
            compareCurrentButton.Click += (_, __) => CompareSelectedSnapshot();
            Button openButton = CreateButton("打开版本");
            openButton.Click += (_, __) => OpenSelectedSnapshot();
            Button restoreButton = CreateButton("恢复到此版本");
            restoreButton.Click += (_, __) => RestoreSelectedSnapshot();
            Button deleteButton = CreateButton("删除版本");
            deleteButton.Click += (_, __) => DeleteSelectedSnapshot();
            Button refreshButton = CreateButton("刷新");
            refreshButton.Click += (_, __) => ReloadEntries();
            buttons.Controls.Add(archiveButton);
            buttons.Controls.Add(compareButton);
            buttons.Controls.Add(compareCurrentButton);
            buttons.Controls.Add(openButton);
            buttons.Controls.Add(restoreButton);
            buttons.Controls.Add(deleteButton);
            buttons.Controls.Add(refreshButton);

            statusLabel = new Label
            {
                AutoSize = true,
                ForeColor = Color.FromArgb(82, 91, 104),
                Margin = new Padding(0, 8, 0, 0)
            };

            layout.Controls.Add(hintPanel, 0, 0);
            layout.Controls.Add(noteRow, 0, 1);
            layout.Controls.Add(versionList, 0, 2);
            layout.Controls.Add(buttons, 0, 3);
            layout.Controls.Add(statusLabel, 0, 4);
            Controls.Add(layout);
            ReloadEntries();
        }

        private void ReloadEntries()
        {
            entries = DocumentVersionArchive.Load(document);
            versionList.BeginUpdate();
            try
            {
                versionList.Items.Clear();
                foreach (DocumentVersionEntry entry in entries)
                {
                    ListViewItem item = new ListViewItem($"V{entry.Number:000}") { Tag = entry };
                    item.SubItems.Add(entry.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"));
                    item.SubItems.Add(entry.Note ?? string.Empty);
                    versionList.Items.Add(item);
                }
            }
            finally
            {
                versionList.EndUpdate();
            }

            ResizeVersionColumns();

            statusLabel.Text = entries.Count == 0
                ? "尚无存档版本。"
                : $"共 {entries.Count} 个版本，当前文档为最新版。";
        }

        private void ArchiveCurrentDocument()
        {
            try
            {
                DocumentVersionEntry entry = DocumentVersionArchive.Archive(document, noteTextBox.Text);
                noteTextBox.Clear();
                ReloadEntries();
                SelectEntry(entry);
                statusLabel.Text = $"已存档 {entry.DisplayName}。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "存档失败：\r\n" + ex.Message, "文档版本管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CompareSelectedSnapshot()
        {
            DocumentVersionEntry entry = GetSelectedEntry();
            if (entry == null)
            {
                MessageBox.Show(this, "请先选择一个版本。", "文档版本管理");
                return;
            }

            try
            {
                Word.Document comparison = DocumentVersionArchive.CompareWithCurrent(application, document, entry);
                if (comparison == null)
                {
                    MessageBox.Show(this, "无法创建比较文档。", "文档版本管理");
                    return;
                }

                ActivateDocumentAndClose(comparison);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "比较版本失败：\r\n" + ex.Message, "文档版本管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CompareWithPreviousSnapshot()
        {
            DocumentVersionEntry entry = GetSelectedEntry();
            if (entry == null)
            {
                MessageBox.Show(this, "请先选择一个版本。", "文档版本管理");
                return;
            }

            DocumentVersionEntry previousEntry = entries
                .Where(item => item.Number < entry.Number)
                .OrderByDescending(item => item.Number)
                .FirstOrDefault();
            if (previousEntry == null)
            {
                MessageBox.Show(this, "该版本是第一个存档版本，没有更早的版本可供比较。", "文档版本管理");
                return;
            }

            try
            {
                Word.Document comparison = DocumentVersionArchive.CompareSnapshots(application, previousEntry, entry);
                if (comparison == null)
                {
                    MessageBox.Show(this, "无法创建比较文档。", "文档版本管理");
                    return;
                }

                ActivateDocumentAndClose(comparison);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "比较版本失败：\r\n" + ex.Message, "文档版本管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OpenSelectedSnapshot()
        {
            DocumentVersionEntry entry = GetSelectedEntry();
            if (entry == null)
            {
                MessageBox.Show(this, "请先选择一个版本。", "文档版本管理");
                return;
            }

            try
            {
                Word.Document opened = DocumentVersionArchive.OpenSnapshot(application, entry);
                if (opened == null)
                {
                    MessageBox.Show(this, "无法打开版本快照。", "文档版本管理");
                    return;
                }

                ActivateDocumentAndClose(opened);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "打开版本失败：\r\n" + ex.Message, "文档版本管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void RestoreSelectedSnapshot()
        {
            DocumentVersionEntry entry = GetSelectedEntry();
            if (entry == null)
            {
                MessageBox.Show(this, "请先选择一个版本。", "文档版本管理");
                return;
            }

            DialogResult result = MessageBox.Show(
                this,
                $"确认将当前文档恢复为 {entry.DisplayName} 吗？\r\n恢复前会自动存档当前最新版，恢复后文档会重新打开。",
                "恢复文档版本",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
            {
                return;
            }

            try
            {
                DocumentVersionArchive.Archive(document, $"恢复 {entry.DisplayName} 前自动存档");
                Word.Document restored = DocumentVersionArchive.RestoreSnapshot(application, document, entry);
                if (restored == null)
                {
                    MessageBox.Show(this, "恢复版本失败，当前文档没有被重新打开。", "文档版本管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ActivateDocumentAndClose(restored);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "恢复版本失败：\r\n" + ex.Message, "文档版本管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void DeleteSelectedSnapshot()
        {
            DocumentVersionEntry entry = GetSelectedEntry();
            if (entry == null)
            {
                MessageBox.Show(this, "请先选择一个版本。", "文档版本管理");
                return;
            }

            if (MessageBox.Show(this, $"确认删除 {entry.DisplayName} 吗？\r\n当前最新版不会受到影响。", "删除版本", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            {
                return;
            }

            try
            {
                DocumentVersionArchive.Delete(document, entry);
                ReloadEntries();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "删除版本失败：\r\n" + ex.Message, "文档版本管理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private DocumentVersionEntry GetSelectedEntry()
        {
            return versionList.SelectedItems.Count == 0
                ? null
                : versionList.SelectedItems[0].Tag as DocumentVersionEntry;
        }

        private void SelectEntry(DocumentVersionEntry entry)
        {
            foreach (ListViewItem item in versionList.Items)
            {
                if (ReferenceEquals(item.Tag, entry))
                {
                    item.Selected = true;
                    item.Focused = true;
                    item.EnsureVisible();
                    break;
                }
            }
        }

        private void ResizeVersionColumns()
        {
            if (versionList.Columns.Count < 3 || versionList.ClientSize.Width <= 0)
            {
                return;
            }

            versionList.Columns[0].Width = 100;
            versionList.Columns[1].Width = 180;
            versionList.Columns[2].Width = Math.Max(220, versionList.ClientSize.Width - 100 - 180 - 6);
        }

        private void ActivateDocumentAndClose(Word.Document targetDocument)
        {
            try
            {
                Hide();
                application.Visible = true;
                targetDocument?.Activate();
            }
            finally
            {
                Close();
            }
        }

        private static Button CreateButton(string text)
        {
            return new Button
            {
                Text = text,
                AutoSize = true,
                MinimumSize = new Size(88, 30),
                Margin = new Padding(0, 0, 8, 0)
            };
        }
    }
}
