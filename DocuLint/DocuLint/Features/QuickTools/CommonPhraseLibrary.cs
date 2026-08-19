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
        internal static string ConfiguredPath
        {
            get
            {
                try
                {
                    return (Properties.Settings.Default.CommonPhraseLibraryPath ?? string.Empty).Trim();
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
        }

        internal static IReadOnlyList<string> LoadConfiguredPhrases()
        {
            return TryLoad(ConfiguredPath, out List<string> phrases, out string _) ? phrases : new List<string>();
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
