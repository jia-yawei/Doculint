namespace DocuLint.Core.Text
{
    public static class PaneTextFormatter
    {
        public static string Normalize(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            return text
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("\a", " ")
                .Replace("\v", " ")
                .Trim();
        }

        public static string Truncate(string text, int maxLength)
        {
            string normalized = Normalize(text);
            if (string.IsNullOrWhiteSpace(normalized) || maxLength <= 0)
            {
                return string.Empty;
            }

            if (normalized.Length <= maxLength)
            {
                return normalized;
            }

            if (maxLength <= 1)
            {
                return normalized.Substring(0, maxLength);
            }

            return normalized.Substring(0, maxLength - 1) + "...";
        }
    }
}
