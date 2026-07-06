using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace DocuLint
{
    internal sealed class DocumentGroupStore
    {
        private static readonly StringComparer PathComparer = StringComparer.OrdinalIgnoreCase;
        private readonly string storagePath;
        private readonly string backupStoragePath;

        public DocumentGroupStore()
        {
            string root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DocuLint");
            storagePath = Path.Combine(root, "document-groups.xml");
            backupStoragePath = Path.Combine(root, "document-groups.backup.xml");
        }

        public DocumentGroupCatalog Load()
        {
            try
            {
                DocumentGroupCatalog catalog = TryLoadFromFile(storagePath);
                if (catalog != null)
                {
                    return catalog;
                }

                DocumentGroupCatalog backupCatalog = TryLoadFromFile(backupStoragePath);
                if (backupCatalog != null)
                {
                    return backupCatalog;
                }
            }
            catch
            {
            }

            return new DocumentGroupCatalog();
        }

        public void Save(DocumentGroupCatalog catalog)
        {
            DocumentGroupCatalog safeCatalog = catalog ?? new DocumentGroupCatalog();
            EnsureCatalogCollections(safeCatalog);
            string directory = Path.GetDirectoryName(storagePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            XmlSerializer serializer = new XmlSerializer(typeof(DocumentGroupCatalog));
            string tempPath = storagePath + ".tmp";
            try
            {
                using (FileStream stream = File.Create(tempPath))
                {
                    serializer.Serialize(stream, safeCatalog);
                }

                if (File.Exists(storagePath))
                {
                    try
                    {
                        File.Copy(storagePath, backupStoragePath, true);
                    }
                    catch
                    {
                    }

                    try
                    {
                        File.Replace(tempPath, storagePath, null, true);
                    }
                    catch
                    {
                        File.Copy(tempPath, storagePath, true);
                        File.Delete(tempPath);
                    }
                }
                else
                {
                    File.Move(tempPath, storagePath);
                }
            }
            finally
            {
                if (File.Exists(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch
                    {
                    }
                }
            }
        }

        public DocumentGroupItem EnsureActiveGroup(DocumentGroupCatalog catalog, string currentDocumentPath)
        {
            DocumentGroupCatalog safeCatalog = catalog ?? new DocumentGroupCatalog();
            DocumentGroupItem activeGroup = safeCatalog.GetActiveGroup();
            if (activeGroup != null)
            {
                return activeGroup;
            }

            string normalizedCurrentPath = TryNormalizeFilePath(currentDocumentPath);
            if (!string.IsNullOrWhiteSpace(normalizedCurrentPath))
            {
                DocumentGroupItem matchedGroup = safeCatalog.GetOrderedGroups()
                    .FirstOrDefault(group => (group.Documents ?? new List<DocumentGroupDocumentItem>())
                        .Any(item => PathComparer.Equals(item?.FilePath, normalizedCurrentPath)));
                if (matchedGroup != null)
                {
                    safeCatalog.ActiveGroupId = matchedGroup.Id;
                    Save(safeCatalog);
                    return matchedGroup;
                }
            }

            DocumentGroupItem fallbackGroup = safeCatalog.GetOrderedGroups().FirstOrDefault();
            if (fallbackGroup != null)
            {
                safeCatalog.ActiveGroupId = fallbackGroup.Id;
                Save(safeCatalog);
            }

            return fallbackGroup;
        }

        public DocumentGroupItem CreateGroup(DocumentGroupCatalog catalog, string groupName)
        {
            string safeName = NormalizeGroupName(groupName);
            EnsureUniqueGroupName(catalog, safeName, null);

            DocumentGroupItem group = new DocumentGroupItem
            {
                Name = safeName
            };

            catalog.Groups.Add(group);
            if (catalog.GetActiveGroup() == null)
            {
                catalog.ActiveGroupId = group.Id;
            }

            Save(catalog);
            return group;
        }

        public void RenameGroup(DocumentGroupCatalog catalog, string groupId, string newName)
        {
            DocumentGroupItem group = FindGroup(catalog, groupId);
            if (group == null)
            {
                throw new InvalidOperationException("未找到要重命名的文档组。");
            }

            string safeName = NormalizeGroupName(newName);
            EnsureUniqueGroupName(catalog, safeName, group.Id);
            group.Name = safeName;
            Save(catalog);
        }

        public void DeleteGroup(DocumentGroupCatalog catalog, string groupId)
        {
            DocumentGroupItem group = FindGroup(catalog, groupId);
            if (group == null)
            {
                return;
            }

            catalog.Groups.Remove(group);
            if (string.Equals(catalog.ActiveGroupId, groupId, StringComparison.OrdinalIgnoreCase))
            {
                catalog.ActiveGroupId = catalog.GetOrderedGroups().FirstOrDefault()?.Id ?? string.Empty;
            }

            Save(catalog);
        }

        public void SetActiveGroup(DocumentGroupCatalog catalog, string groupId)
        {
            DocumentGroupItem group = FindGroup(catalog, groupId);
            if (group == null)
            {
                throw new InvalidOperationException("未找到要设置为活动状态的文档组。");
            }

            catalog.ActiveGroupId = group.Id;
            Save(catalog);
        }

        public void AddDocumentToGroup(DocumentGroupCatalog catalog, string groupId, string filePath)
        {
            DocumentGroupItem group = FindGroup(catalog, groupId);
            if (group == null)
            {
                throw new InvalidOperationException("未找到目标文档组。");
            }

            string fullPath = NormalizeFilePath(filePath);
            if (group.Documents.Any(item => PathComparer.Equals(item.FilePath, fullPath)))
            {
                return;
            }

            FileInfo fileInfo = new FileInfo(fullPath);
            group.Documents.Add(new DocumentGroupDocumentItem
            {
                FilePath = fullPath,
                DisplayName = Path.GetFileName(fullPath),
                AddedAt = DateTime.Now,
                LastKnownWriteTime = fileInfo.Exists ? (DateTime?)fileInfo.LastWriteTime : null
            });

            group.Documents = group.Documents
                .OrderByDescending(item => item.LastKnownWriteTime ?? DateTime.MinValue)
                .ThenBy(item => item.DisplayName ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            Save(catalog);
        }

        public void RemoveDocumentFromGroup(DocumentGroupCatalog catalog, string groupId, string filePath)
        {
            DocumentGroupItem group = FindGroup(catalog, groupId);
            if (group == null)
            {
                return;
            }

            string fullPath = NormalizeFilePath(filePath);
            DocumentGroupDocumentItem existing = group.Documents
                .FirstOrDefault(item => PathComparer.Equals(item.FilePath, fullPath));
            if (existing == null)
            {
                return;
            }

            group.Documents.Remove(existing);
            Save(catalog);
        }

        public void RefreshDocumentMetadata(DocumentGroupCatalog catalog)
        {
            foreach (DocumentGroupItem group in catalog.Groups ?? new List<DocumentGroupItem>())
            {
                foreach (DocumentGroupDocumentItem item in group.Documents ?? new List<DocumentGroupDocumentItem>())
                {
                    if (string.IsNullOrWhiteSpace(item.FilePath))
                    {
                        continue;
                    }

                    item.DisplayName = Path.GetFileName(item.FilePath);
                    if (File.Exists(item.FilePath))
                    {
                        item.LastKnownWriteTime = File.GetLastWriteTime(item.FilePath);
                    }
                }
            }

            Save(catalog);
        }

        public int CleanupInvalidDocuments(DocumentGroupCatalog catalog)
        {
            int removedCount = 0;

            foreach (DocumentGroupItem group in catalog.Groups ?? new List<DocumentGroupItem>())
            {
                if (group.Documents == null)
                {
                    continue;
                }

                List<DocumentGroupDocumentItem> invalidItems = group.Documents
                    .Where(item => item == null ||
                        string.IsNullOrWhiteSpace(item.FilePath) ||
                        !File.Exists(item.FilePath))
                    .ToList();

                foreach (DocumentGroupDocumentItem invalidItem in invalidItems)
                {
                    group.Documents.Remove(invalidItem);
                }

                removedCount += invalidItems.Count;
            }

            if (removedCount > 0)
            {
                Save(catalog);
            }

            return removedCount;
        }

        private static DocumentGroupItem FindGroup(DocumentGroupCatalog catalog, string groupId)
        {
            return catalog?.Groups?.FirstOrDefault(item =>
                string.Equals(item.Id, groupId, StringComparison.OrdinalIgnoreCase));
        }

        private static string NormalizeGroupName(string groupName)
        {
            string safeName = (groupName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(safeName))
            {
                throw new InvalidOperationException("文档组名称不能为空。");
            }

            return safeName;
        }

        private static void EnsureUniqueGroupName(DocumentGroupCatalog catalog, string groupName, string excludeGroupId)
        {
            bool exists = catalog.Groups.Any(item =>
                !string.Equals(item.Id, excludeGroupId, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(item.Name, groupName, StringComparison.CurrentCultureIgnoreCase));

            if (exists)
            {
                throw new InvalidOperationException("已存在同名文档组。");
            }
        }

        private static string NormalizeFilePath(string filePath)
        {
            string safePath = (filePath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(safePath))
            {
                throw new InvalidOperationException("文档路径不能为空。");
            }

            return Path.GetFullPath(safePath);
        }

        private static string TryNormalizeFilePath(string filePath)
        {
            try
            {
                return NormalizeFilePath(filePath);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static void EnsureCatalogCollections(DocumentGroupCatalog catalog)
        {
            if (catalog == null)
            {
                return;
            }

            catalog.Groups = catalog.Groups ?? new List<DocumentGroupItem>();
            foreach (DocumentGroupItem group in catalog.Groups)
            {
                if (group == null)
                {
                    continue;
                }

                group.Documents = group.Documents ?? new List<DocumentGroupDocumentItem>();
                group.CapturedContents = group.CapturedContents ?? new List<DocumentGroupCapturedContentItem>();
            }
        }

        private static DocumentGroupCatalog TryLoadFromFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return null;
            }

            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(DocumentGroupCatalog));
                using (FileStream stream = File.OpenRead(path))
                {
                    DocumentGroupCatalog catalog = serializer.Deserialize(stream) as DocumentGroupCatalog;
                    DocumentGroupCatalog safeCatalog = catalog ?? new DocumentGroupCatalog();
                    EnsureCatalogCollections(safeCatalog);
                    return safeCatalog;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
