using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace DocuLint
{
    internal enum RequirementTraceTemplate
    {
        SrsToSds = 0,
        SdsToSrs = 1,
        SdsToSdd = 2,
        SddToSds = 3
    }

    internal enum RequirementTrackingDocumentKind
    {
        Unknown = 0,
        SystemSpecification = 1,
        RequirementSpecification = 2,
        SoftwareDesignDescription = 3,
        SoftwareTestDescription = 4
    }

    public sealed class RequirementItem
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public string SectionNumber { get; set; }

        public object BookmarkOrRange { get; set; }

        public int Start { get; set; }

        public bool IsMapped { get; set; }

        public string DisplayText
        {
            get
            {
                string sectionPrefix = string.IsNullOrWhiteSpace(SectionNumber)
                    ? string.Empty
                    : "[" + SectionNumber + "] ";
                string shortId = GetDisplayRequirementId(Id);

                if (!string.IsNullOrWhiteSpace(shortId) && !string.IsNullOrWhiteSpace(Name))
                {
                    return sectionPrefix + shortId + " " + Name;
                }

                if (!string.IsNullOrWhiteSpace(shortId))
                {
                    return sectionPrefix + shortId;
                }

                return sectionPrefix + (Name ?? string.Empty);
            }
        }

        public override string ToString()
        {
            return DisplayText;
        }

        private static string GetShortRequirementId(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return string.Empty;
            }

            MatchCollection matches = Regex.Matches(id, @"(?<prefix>SSS|SRS|SDS|SDD)\s*[-－–—]\s*\d+", RegexOptions.IgnoreCase);
            if (matches.Count == 0)
            {
                return id;
            }

            return Regex.Replace(matches[matches.Count - 1].Value, @"\s*[-－–—]\s*", "-").ToUpperInvariant();
        }

        internal static string GetDisplayRequirementId(string id)
        {
            string shortId = GetShortRequirementId(id);
            if (string.IsNullOrWhiteSpace(shortId))
            {
                return id ?? string.Empty;
            }

            string raw = (id ?? string.Empty).Trim();
            return string.Equals(raw, shortId, StringComparison.OrdinalIgnoreCase)
                ? shortId
                : "..." + shortId;
        }

        internal static bool ContainsRequirementPrefix(string id, string prefix)
        {
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(prefix))
            {
                return false;
            }

            return id.IndexOf(prefix, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }

    internal sealed class RequirementTrackingDocumentOption
    {
        public string FullName { get; set; }

        public string DisplayName { get; set; }

        public bool IsCurrentlyOpen { get; set; }

        public override string ToString()
        {
            return DisplayName ?? string.Empty;
        }
    }

    internal sealed class RequirementTrackingDocumentSnapshot
    {
        public string FullName { get; set; }

        public string DisplayName { get; set; }

        public List<RequirementItem> Requirements { get; set; } = new List<RequirementItem>();
    }

    internal sealed class RequirementTraceMapping
    {
        public string SourceRequirementId { get; set; }

        public List<string> TargetRequirementIds { get; set; } = new List<string>();
    }
}
