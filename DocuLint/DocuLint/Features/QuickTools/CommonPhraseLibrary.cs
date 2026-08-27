using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;

namespace DocuLint
{
    internal static class CommonPhraseLibrary
    {
        private static readonly object CacheLock = new object();
        private static string cachedPath = string.Empty;
        private static DateTime cachedLastWriteUtc = DateTime.MinValue;
        private static List<string> cachedPhrases = new List<string>();

        internal sealed class Suggestion
        {
            internal string Phrase { get; set; }

            internal int Score { get; set; }
        }

        internal static string ConfiguredPath
        {
            get
            {
                try
                {
                    string configuredPath = (Properties.Settings.Default.CommonPhraseLibraryPath ?? string.Empty).Trim();
                    return !string.IsNullOrWhiteSpace(configuredPath) && File.Exists(configuredPath)
                        ? configuredPath
                        : PluginDataStore.CommonPhrasesPath;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        internal static void SaveConfiguredPath(string filePath)
        {
            Properties.Settings.Default.CommonPhraseLibraryPath = (filePath ?? string.Empty).Trim();
            Properties.Settings.Default.Save();
            InvalidateCache();
        }

        internal static IReadOnlyList<string> LoadConfiguredPhrases()
        {
            string path = ConfiguredPath;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                InvalidateCache();
                return new List<string>();
            }

            DateTime lastWriteUtc;
            try
            {
                lastWriteUtc = File.GetLastWriteTimeUtc(path);
            }
            catch
            {
                lastWriteUtc = DateTime.MinValue;
            }

            lock (CacheLock)
            {
                if (string.Equals(cachedPath, path, StringComparison.OrdinalIgnoreCase)
                    && cachedLastWriteUtc == lastWriteUtc)
                {
                    return cachedPhrases.ToList();
                }
            }

            if (!TryLoad(path, out List<string> phrases, out string _))
            {
                InvalidateCache();
                return new List<string>();
            }

            lock (CacheLock)
            {
                cachedPath = path;
                cachedLastWriteUtc = lastWriteUtc;
                cachedPhrases = phrases.ToList();
                return cachedPhrases.ToList();
            }
        }

        internal static IReadOnlyList<Suggestion> FindSimilar(string input, int maximum = 6)
        {
            string seed = Normalize(input);
            if (seed.Length < 2 || maximum <= 0)
            {
                return new List<Suggestion>();
            }

            List<Suggestion> matches = new List<Suggestion>();
            foreach (string phrase in LoadConfiguredPhrases())
            {
                string normalizedPhrase = Normalize(phrase);
                if (normalizedPhrase.Length == 0)
                {
                    continue;
                }

                int score = Score(seed, normalizedPhrase);
                if (score >= 35)
                {
                    matches.Add(new Suggestion { Phrase = phrase, Score = score });
                }
            }

            return matches
                .OrderByDescending(item => item.Score)
                .ThenBy(item => item.Phrase, StringComparer.Ordinal)
                .Take(maximum)
                .ToList();
        }

        internal static void InvalidateCache()
        {
            lock (CacheLock)
            {
                cachedPath = string.Empty;
                cachedLastWriteUtc = DateTime.MinValue;
                cachedPhrases = new List<string>();
            }
        }

        private static string Normalize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return new string(value
                .Where(character => !char.IsWhiteSpace(character) && !char.IsPunctuation(character))
                .ToArray())
                .ToLowerInvariant();
        }

        private static int Score(string seed, string phrase)
        {
            if (phrase.StartsWith(seed, StringComparison.Ordinal))
            {
                return 100 + Math.Min(20, seed.Length);
            }

            int containsAt = phrase.IndexOf(seed, StringComparison.Ordinal);
            if (containsAt >= 0)
            {
                return 78 - Math.Min(20, containsAt * 2);
            }

            int prefixLength = Math.Min(seed.Length, phrase.Length);
            int distance = LevenshteinDistance(seed, phrase.Substring(0, prefixLength));
            int maximumDistance = Math.Max(seed.Length, prefixLength);
            return maximumDistance == 0
                ? 0
                : 70 - (distance * 40 / maximumDistance);
        }

        private static int LevenshteinDistance(string first, string second)
        {
            int[] previous = new int[second.Length + 1];
            int[] current = new int[second.Length + 1];
            for (int index = 0; index <= second.Length; index++)
            {
                previous[index] = index;
            }

            for (int row = 1; row <= first.Length; row++)
            {
                current[0] = row;
                for (int column = 1; column <= second.Length; column++)
                {
                    int substitution = previous[column - 1] + (first[row - 1] == second[column - 1] ? 0 : 1);
                    current[column] = Math.Min(
                        Math.Min(previous[column] + 1, current[column - 1] + 1),
                        substitution);
                }

                int[] swap = previous;
                previous = current;
                current = swap;
            }

            return previous[second.Length];
        }

        internal static bool TryLoad(string filePath, out List<string> phrases, out string errorMessage)
        {
            phrases = new List<string>();
            errorMessage = string.Empty;
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                errorMessage = "未找到常用语库文件。";
                return false;
            }

            try
            {
                string json = File.ReadAllText(filePath);
                object value = new JavaScriptSerializer().DeserializeObject(json);
                IEnumerable values = value as IEnumerable;
                if (values == null || value is string)
                {
                    errorMessage = "常用语库必须是 JSON 字符串数组。";
                    return false;
                }

                foreach (object item in values)
                {
                    string phrase = item as string;
                    if (phrase == null)
                    {
                        errorMessage = "常用语库目前仅支持纯文本 JSON 字符串数组。";
                        phrases.Clear();
                        return false;
                    }

                    phrase = phrase.Trim();
                    if (phrase.Length > 0)
                    {
                        phrases.Add(phrase);
                    }
                }

                phrases = phrases
                    .Distinct(StringComparer.Ordinal)
                    .ToList();
                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "常用语库不是有效的 JSON 文件：" + ex.Message;
                return false;
            }
        }
    }
}
