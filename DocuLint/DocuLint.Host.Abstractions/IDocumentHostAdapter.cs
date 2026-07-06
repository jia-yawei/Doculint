using System.Collections.Generic;
using DocuLint.Host.Abstractions.Models;

namespace DocuLint.Host.Abstractions
{
    public interface IDocumentHostAdapter
    {
        string GetActiveDocumentName();

        IReadOnlyList<CommentRecord> GetCommentRecords();

        void NavigateTo(int start);
    }
}
