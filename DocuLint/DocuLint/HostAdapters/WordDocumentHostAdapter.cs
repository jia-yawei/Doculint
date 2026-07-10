using System;
using System.Collections.Generic;
using System.Linq;
using DocuLint.Host.Abstractions;
using DocuLint.Host.Abstractions.Models;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal sealed class WordDocumentHostAdapter : IDocumentHostAdapter
    {
        private readonly Func<Word.Application> appProvider;

        internal WordDocumentHostAdapter(Func<Word.Application> appProvider)
        {
            this.appProvider = appProvider ?? throw new ArgumentNullException(nameof(appProvider));
        }

        public string GetActiveDocumentName()
        {
            Word.Document doc = GetActiveDocument();
            return doc == null ? "当前文档" : doc.Name;
        }

        public IReadOnlyList<CommentRecord> GetCommentRecords()
        {
            List<CommentRecord> records = new List<CommentRecord>();
            Word.Document doc = GetActiveDocument();
            if (doc == null)
            {
                return records;
            }

            int count = 0;
            try
            {
                count = doc.Comments == null ? 0 : doc.Comments.Count;
            }
            catch
            {
                count = 0;
            }

            for (int i = 1; i <= count; i++)
            {
                Word.Comment comment = null;
                try
                {
                    comment = doc.Comments[i];
                }
                catch
                {
                    continue;
                }

                if (comment?.Scope == null)
                {
                    continue;
                }

                int start;
                try
                {
                    start = comment.Scope.Start;
                }
                catch
                {
                    continue;
                }

                string bodyText = string.Empty;
                string author = string.Empty;

                try
                {
                    bodyText = comment.Range?.Text ?? string.Empty;
                }
                catch
                {
                }

                try
                {
                    author = comment.Author ?? string.Empty;
                }
                catch
                {
                }

                records.Add(new CommentRecord
                {
                    Start = start,
                    ScopeText = comment.Scope.Text ?? string.Empty,
                    BodyText = bodyText,
                    Author = author
                });
            }

            return records.OrderBy(item => item.Start).ToList();
        }

        public void NavigateTo(int start)
        {
            Word.Application app = appProvider();
            Word.Document doc = app?.ActiveDocument;
            if (doc == null)
            {
                return;
            }

            int contentEnd = Math.Max(0, doc.Content.End - 1);
            int safeStart = Math.Max(0, Math.Min(start, contentEnd));

            try
            {
                Word.Range target = doc.Range(safeStart, safeStart);
                target.Select();
            }
            catch
            {
            }
        }

        private Word.Document GetActiveDocument()
        {
            Word.Application app = appProvider();
            return app?.ActiveDocument;
        }
    }
}
