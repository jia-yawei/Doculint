using System.Collections.Generic;
using System.Linq;
using DocuLint.Core.Text;
using DocuLint.Host.Abstractions.Models;

namespace DocuLint.Core.Comments
{
    public static class CommentPaneEntryBuilder
    {
        public static IReadOnlyList<ListPaneEntry> Build(IEnumerable<CommentRecord> records, int maxLength = 140)
        {
            List<ListPaneEntry> entries = new List<ListPaneEntry>();
            if (records == null)
            {
                return entries;
            }

            foreach (CommentRecord record in records.Where(item => item != null))
            {
                string scopeText = PaneTextFormatter.Normalize(record.ScopeText);
                string bodyText = PaneTextFormatter.Normalize(record.BodyText);
                string author = PaneTextFormatter.Normalize(record.Author);

                string anchor = string.IsNullOrWhiteSpace(scopeText) ? "(No anchor text)" : scopeText;
                string suggestion = string.IsNullOrWhiteSpace(bodyText) ? string.Empty : " | " + bodyText;
                string authorPart = string.IsNullOrWhiteSpace(author) ? string.Empty : "[" + author + "] ";

                entries.Add(new ListPaneEntry
                {
                    Start = record.Start,
                    Text = PaneTextFormatter.Truncate(authorPart + anchor + suggestion, maxLength)
                });
            }

            return entries.OrderBy(entry => entry.Start).ToList();
        }
    }
}
