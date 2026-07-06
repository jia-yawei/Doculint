using System.Collections.Generic;
using System.Linq;

namespace DocuLint
{
    internal sealed class DocumentBasicInfo
    {
        public List<DocumentBasicInfoField> Fields { get; set; } = new List<DocumentBasicInfoField>();

        public bool HasAnyValue()
        {
            return (Fields ?? new List<DocumentBasicInfoField>())
                .Any(field => field != null &&
                    (!string.IsNullOrWhiteSpace(field.Name) || !string.IsNullOrWhiteSpace(field.Value)));
        }
    }

    internal sealed class DocumentBasicInfoField
    {
        public string Name { get; set; } = string.Empty;

        public string Value { get; set; } = string.Empty;
    }
}
