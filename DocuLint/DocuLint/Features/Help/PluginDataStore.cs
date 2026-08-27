using System;
using System.IO;
using System.Reflection;

namespace DocuLint
{
    internal static class PluginDataStore
    {
        private const string ProductFolder = "DocuLint";
        private const string CommonPhrasesFileName = "common-phrases.json";
        private const string CalibrationFileName = "text-check-calibration.txt";
        private const string StandardFileName = "GJB438C.pdf";
        private const string CommonPhrasesResourceName = "DocuLint.DefaultData.CommonPhrases";
        private const string CalibrationResourceName = "DocuLint.DefaultData.Calibration";

        internal static string DataRoot => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            ProductFolder,
            "Data");

        internal static string CommonPhrasesFolder => Path.Combine(DataRoot, "CommonPhrases");

        internal static string ConfigFolder => Path.Combine(DataRoot, "Config");

        internal static string StandardsFolder => Path.Combine(DataRoot, "Standards");

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

        internal static string StandardPath => Path.Combine(StandardsFolder, StandardFileName);

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
