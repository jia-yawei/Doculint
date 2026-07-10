using System.Collections.Generic;
using System.Linq;

namespace DocuLint
{
    public enum OutlineNumberPattern
    {
        Decimal,
        Parenthesized,
        Dotted
    }

    public class OutlineListRebuildOptions
    {
        public HashSet<int> SelectedLevels { get; set; } = new HashSet<int>();

        public OutlineNumberPattern NumberPattern { get; set; }

        public int Alignment { get; set; }

        public int TrailingCharacter { get; set; }

        public int NumberTextSpacing { get; set; } = 1;

        public bool ClearManualNumbering { get; set; } = true;

        public int MaxSelectedLevel => SelectedLevels.Count == 0 ? 0 : SelectedLevels.Max();
    }

    public class OutlineRebuildResult
    {
        public IReadOnlyList<int> SelectedLevels { get; set; }

        public string ScanScope { get; set; }

        public int TargetParagraphCount { get; set; }

        public int ClearedListCount { get; set; }

        public int ClearedManualNumberCount { get; set; }

        public int AppliedParagraphCount { get; set; }

        public long ScanMilliseconds { get; set; }

        public long CleanupMilliseconds { get; set; }

        public long ApplyMilliseconds { get; set; }

        public long DurationMilliseconds { get; set; }

        public Dictionary<int, string> LinkedStyles { get; set; }
    }
}
