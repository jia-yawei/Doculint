using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;

namespace DocuLint
{
    internal static class PluginDataStore
    {
        private const string ProductFolder = "DocuLint";
        private const string CommonPhrasesFileName = "common-phrases.json";
        private const string CalibrationFileName = "text-check-calibration.txt";
        private const string PluginSettingsFileName = "plugin-settings.json";
        private const string CommonPhrasesResourceName = "DocuLint.DefaultData.CommonPhrases";
        private const string CalibrationResourceName = "DocuLint.DefaultData.Calibration";

        internal static string DataRoot => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            ProductFolder,
            "Data");

        internal static string CommonPhrasesFolder => Path.Combine(DataRoot, "CommonPhrases");

        internal static string ConfigFolder => Path.Combine(DataRoot, "Config");

        internal static string CommonPhrasesPath
        {
            get
            {
                EnsureDefaultFiles();
                return Path.Combine(CommonPhrasesFolder, CommonPhrasesFileName);
            }
        }

        internal static string CalibrationPath
        {
            get
            {
                EnsureDefaultFiles();
                return Path.Combine(ConfigFolder, CalibrationFileName);
            }
        }

        internal static string PluginSettingsPath => Path.Combine(ConfigFolder, PluginSettingsFileName);

        internal static string GetStandardLibraryDirectory()
        {
            try
            {
                if (!File.Exists(PluginSettingsPath))
                {
                    return string.Empty;
                }

                string json = File.ReadAllText(PluginSettingsPath, Encoding.UTF8);
                Dictionary<string, object> settings = new JavaScriptSerializer()
                    .Deserialize<Dictionary<string, object>>(json);
                object directory;
                return settings != null
                    && settings.TryGetValue("StandardLibraryDirectory", out directory)
                    ? (directory as string ?? string.Empty).Trim()
                    : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        internal static void SaveStandardLibraryDirectory(string directory)
        {
            string fullPath = Path.GetFullPath((directory ?? string.Empty).Trim());
            if (!Directory.Exists(fullPath))
            {
                throw new DirectoryNotFoundException("标准文件夹不存在：" + fullPath);
            }

            Directory.CreateDirectory(ConfigFolder);
            Dictionary<string, object> settings = new Dictionary<string, object>
            {
                { "StandardLibraryDirectory", fullPath }
            };
            string json = new JavaScriptSerializer().Serialize(settings);
            File.WriteAllText(PluginSettingsPath, json, new UTF8Encoding(false));
        }

        internal static void EnsureDefaultFiles()
        {
            EnsureEmbeddedFile(
                Path.Combine(CommonPhrasesFolder, CommonPhrasesFileName),
                CommonPhrasesResourceName);
            EnsureEmbeddedFile(
                Path.Combine(ConfigFolder, CalibrationFileName),
                CalibrationResourceName);
        }

        private static void EnsureEmbeddedFile(string target, string resourceName)
        {
            if (File.Exists(target))
            {
                return;
            }

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(target));
                using (Stream input = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    if (input == null)
                    {
                        return;
                    }

                    using (FileStream output = File.Create(target))
                    {
                        input.CopyTo(output);
                    }
                }
            }
            catch
            {
                // Optional defaults must not prevent the add-in from starting.
            }
        }
    }
}
