using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace DocuLint
{
    internal static class CapturedContentPreviewRenderer
    {
        private static readonly object SyncRoot = new object();
        private static readonly Dictionary<string, string> HtmlCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private static readonly string PreviewRoot = Path.Combine(Path.GetTempPath(), "DocuLint", "captured-content-preview");

        public static bool TryRenderHtml(string wordOpenXml, out string htmlPath)
        {
            htmlPath = string.Empty;
            if (string.IsNullOrWhiteSpace(wordOpenXml))
            {
                return false;
            }

            string key = ComputeHash(wordOpenXml);
            lock (SyncRoot)
            {
                if (HtmlCache.TryGetValue(key, out string cachedPath) && File.Exists(cachedPath))
                {
                    htmlPath = cachedPath;
                    return true;
                }
            }

            try
            {
                Directory.CreateDirectory(PreviewRoot);
                CleanupOldDirectories();
            }
            catch
            {
                return false;
            }

            string targetDirectory = Path.Combine(PreviewRoot, key);
            string targetHtmlPath = Path.Combine(targetDirectory, "preview.html");
            if (File.Exists(targetHtmlPath))
            {
                lock (SyncRoot)
                {
                    HtmlCache[key] = targetHtmlPath;
                }

                htmlPath = targetHtmlPath;
                return true;
            }

            try
            {
                Directory.CreateDirectory(targetDirectory);
            }
            catch
            {
                return false;
            }

            Word.Application app = Globals.ThisAddIn?.Application;
            if (app == null)
            {
                return false;
            }

            object missing = Type.Missing;
            Word.Document tempDocument = null;
            try
            {
                object visible = false;
                tempDocument = app.Documents.Add(ref missing, ref missing, ref missing, ref visible);
                Word.Range range = tempDocument.Range(0, 0);
                object transform = Type.Missing;
                range.InsertXML(wordOpenXml, ref transform);

                object fileName = targetHtmlPath;
                object fileFormat = Word.WdSaveFormat.wdFormatFilteredHTML;
                tempDocument.SaveAs2(
                    ref fileName,
                    ref fileFormat,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing,
                    ref missing);

                lock (SyncRoot)
                {
                    HtmlCache[key] = targetHtmlPath;
                }

                htmlPath = targetHtmlPath;
                return File.Exists(targetHtmlPath);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (tempDocument != null)
                {
                    object saveOption = Word.WdSaveOptions.wdDoNotSaveChanges;
                    try
                    {
                        tempDocument.Close(ref saveOption, ref missing, ref missing);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static string ComputeHash(string text)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] bytes = Encoding.UTF8.GetBytes(text ?? string.Empty);
                byte[] hash = sha.ComputeHash(bytes);
                StringBuilder sb = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                {
                    sb.Append(b.ToString("x2"));
                }

                return sb.ToString();
            }
        }

        private static void CleanupOldDirectories()
        {
            string[] directories;
            try
            {
                directories = Directory.GetDirectories(PreviewRoot);
            }
            catch
            {
                return;
            }

            DateTime threshold = DateTime.Now.AddDays(-2);
            foreach (string directory in directories.Where(path =>
            {
                try
                {
                    return Directory.GetLastWriteTime(path) < threshold;
                }
                catch
                {
                    return false;
                }
            }))
            {
                try
                {
                    Directory.Delete(directory, true);
                }
                catch
                {
                }
            }
        }
    }
}
