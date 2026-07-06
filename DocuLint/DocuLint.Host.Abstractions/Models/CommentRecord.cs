namespace DocuLint.Host.Abstractions.Models
{
    public sealed class CommentRecord
    {
        public int Start { get; set; }

        public string ScopeText { get; set; }

        public string BodyText { get; set; }

        public string Author { get; set; }
    }
}
