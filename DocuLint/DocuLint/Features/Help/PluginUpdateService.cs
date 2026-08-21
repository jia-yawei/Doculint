using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace DocuLint
{
    internal sealed class PluginUpdateManifest
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string Notes { get; set; }
        public string PackageUrl { get; set; }
        public string Sha256 { get; set; }
        public string PackageFileName { get; set; }

        internal Version ParsedVersion
        {
            get
            {
                return System.Version.TryParse(Version, out System.Version parsed) ? parsed : new System.Version(0, 0, 0, 0);
            }
        }
    }

    internal static class PluginUpdateService
    {
        internal const string DefaultGitHubManifestUrl =
            "https://raw.githubusercontent.com/jia-yawei/Doculint/main/update/latest.json";
        internal const string ManifestFileName = "latest.json";

        static PluginUpdateService()
        {
            // GitHub requires TLS 1.2 or newer. Explicitly select TLS 1.2 for
            // .NET Framework installations whose system default is obsolete.
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        }

        internal static string CurrentVersionText =>
            typeof(PluginUpdateService).Assembly.GetName().Version?.ToString() ?? "0.0.0.0";

        internal static Version CurrentVersion =>
            System.Version.TryParse(CurrentVersionText, out System.Version version) ? version : new System.Version(0, 0, 0, 0);

        internal static PluginUpdateManifest ReadManifest(string manifestText)
        {
            if (string.IsNullOrWhiteSpace(manifestText))
            {
                return null;
            }

            try
            {
                return new JavaScriptSerializer().Deserialize<PluginUpdateManifest>(
                    manifestText.TrimStart('\uFEFF'));
            }
            catch
            {
                return null;
            }
        }

        internal static PluginUpdateManifest LoadFromGitHub(string manifestUrl, out string error)
        {
            error = string.Empty;
            try
            {
                using (WebClient client = CreateClient())
                {
                    PluginUpdateManifest manifest = ReadManifest(client.DownloadString(manifestUrl));
                    if (manifest == null)
                    {
                        error = "GitHub 更新清单格式无效。";
                    }

                    return manifest;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }

        internal static PluginUpdateManifest LoadFromFolder(string folder, out string error)
        {
            error = string.Empty;
            try
            {
                string path = Path.Combine(folder ?? string.Empty, ManifestFileName);
                if (!File.Exists(path))
                {
                    error = "文件夹中未找到 latest.json。";
                    return null;
                }

                PluginUpdateManifest manifest = ReadManifest(File.ReadAllText(path, Encoding.UTF8));
                if (manifest == null)
                {
                    error = "更新清单格式无效。";
                }

                return manifest;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }

        internal static string ResolvePackagePath(PluginUpdateManifest manifest, string localFolder)
        {
            if (manifest == null)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(localFolder))
            {
                string packageName = manifest.PackageFileName;
                if (string.IsNullOrWhiteSpace(packageName) && !string.IsNullOrWhiteSpace(manifest.PackageUrl))
                {
                    packageName = Path.GetFileName(manifest.PackageUrl);
                }

                string localPath = Path.Combine(localFolder, packageName ?? string.Empty);
                if (File.Exists(localPath))
                {
                    return localPath;
                }
            }

            return null;
        }

        internal static string DownloadPackage(PluginUpdateManifest manifest, string targetFolder, out string error)
        {
            return DownloadPackage(manifest, targetFolder, null, out error);
        }

        internal static string DownloadPackage(
            PluginUpdateManifest manifest,
            string targetFolder,
            Action<long, long> progress,
            out string error)
        {
            error = string.Empty;
            if (manifest == null || string.IsNullOrWhiteSpace(manifest.PackageUrl))
            {
                error = "更新清单中没有安装包地址。";
                return null;
            }

            try
            {
                Directory.CreateDirectory(targetFolder);
                string name = string.IsNullOrWhiteSpace(manifest.PackageFileName)
                    ? Path.GetFileName(new Uri(manifest.PackageUrl).LocalPath)
                    : manifest.PackageFileName;
                string target = Path.Combine(targetFolder, string.IsNullOrWhiteSpace(name) ? "DocuLint-update.vsto" : name);
                using (WebClient client = CreateClient())
                using (Stream input = client.OpenRead(manifest.PackageUrl))
                using (FileStream output = File.Create(target))
                {
                    long totalBytes = -1;
                    long.TryParse(client.ResponseHeaders[HttpResponseHeader.ContentLength], out totalBytes);
                    long receivedBytes = 0;
                    byte[] buffer = new byte[64 * 1024];
                    int read;
                    progress?.Invoke(0, totalBytes);
                    while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        output.Write(buffer, 0, read);
                        receivedBytes += read;
                        progress?.Invoke(receivedBytes, totalBytes);
                    }
                }

                if (!VerifySha256(target, manifest.Sha256))
                {
                    File.Delete(target);
                    error = "安装包 SHA-256 校验失败。";
                    return null;
                }

                return target;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return null;
            }
        }

        internal static bool VerifySha256(string path, string expected)
        {
            if (string.IsNullOrWhiteSpace(expected))
            {
                return true;
            }

            using (SHA256 sha = SHA256.Create())
            using (FileStream stream = File.OpenRead(path))
            {
                string actual = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", string.Empty);
                return string.Equals(actual, expected.Trim(), StringComparison.OrdinalIgnoreCase);
            }
        }

        private static WebClient CreateClient()
        {
            WebClient client = new WebClient();
            client.Headers[HttpRequestHeader.UserAgent] = "DocuLint-UpdateClient/1.0";
            client.Encoding = Encoding.UTF8;
            return client;
        }
    }
}
