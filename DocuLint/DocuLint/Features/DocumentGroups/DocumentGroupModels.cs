using System;
using System.Collections.Generic;
using System.Linq;

namespace DocuLint
{
    [Serializable]
    public sealed class DocumentGroupCatalog
    {
        public List<DocumentGroupItem> Groups { get; set; } = new List<DocumentGroupItem>();

        public string ActiveGroupId { get; set; } = string.Empty;

        public IEnumerable<DocumentGroupItem> GetOrderedGroups()
        {
            return (Groups ?? new List<DocumentGroupItem>())
                .OrderBy(item => item.Name ?? string.Empty, StringComparer.CurrentCultureIgnoreCase);
        }

        public DocumentGroupItem GetActiveGroup()
        {
            return (Groups ?? new List<DocumentGroupItem>())
                .FirstOrDefault(item => string.Equals(item.Id, ActiveGroupId, StringComparison.OrdinalIgnoreCase));
        }
    }

    [Serializable]
    public sealed class DocumentGroupItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public string Name { get; set; } = string.Empty;

        public List<DocumentGroupDocumentItem> Documents { get; set; } = new List<DocumentGroupDocumentItem>();

        public List<DocumentGroupCapturedContentItem> CapturedContents { get; set; } = new List<DocumentGroupCapturedContentItem>();
    }

    [Serializable]
    public sealed class DocumentGroupDocumentItem
    {
        public string FilePath { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public DateTime AddedAt { get; set; } = DateTime.Now;

        public DateTime? LastKnownWriteTime { get; set; }
    }

    [Serializable]
    public sealed class DocumentGroupCapturedContentItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        public string Title { get; set; } = string.Empty;

        public string PreviewText { get; set; } = string.Empty;

        public string ContentWordOpenXml { get; set; } = string.Empty;

        public DateTime CreatedAt { get; set; } = DateTime.Now;

        public DateTime UpdatedAt { get; set; } = DateTime.Now;
    }
}
