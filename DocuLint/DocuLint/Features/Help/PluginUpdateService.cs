using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
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
        private const string RawGitHubContentPrefix =
            "https://raw.githubusercontent.com/jia-yawei/Doculint/main/";
        private const string JsDelivrContentPrefix =
            "https://cdn.jsdelivr.net/gh/jia-yawei/Doculint@main/";
        private const string GitHubReleasePrefix = "https://github.com/";
        private const string GhFastPrefix = "https://ghfast.top/";
        private static readonly string[] DefaultManifestUrls =
        {
            RawGitHubContentPrefix + "update/latest.json",
            JsDelivrContentPrefix + "update/latest.json"
        };
        private const int ConnectionTimeoutMilliseconds = 30000;
        private const int ReadWriteTimeoutMilliseconds = 600000;

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

        internal static PluginUpdateManifest LoadFromGitHub(out string error)
        {
            StringBuilder errors = new StringBuilder();
            PluginUpdateManifest newestManifest = null;
            foreach (string manifestUrl in DefaultManifestUrls)
            {
                string sourceError;
                PluginUpdateManifest manifest = LoadFromGitHub(manifestUrl, out sourceError);
                if (manifest != null)
                {
                    if (newestManifest == null || manifest.ParsedVersion > newestManifest.ParsedVersion)
                    {
                        newestManifest = manifest;
                    }

                    continue;
                }

                if (!string.IsNullOrWhiteSpace(sourceError))
                {
                    if (errors.Length > 0)
                    {
                        errors.AppendLine();
                    }

                    errors.Append(sourceError);
                }
            }

            if (newestManifest != null)
            {
                error = string.Empty;
                return newestManifest;
            }

            error = errors.ToString();
            return null;
        }

        private static PluginUpdateManifest LoadFromGitHub(string manifestUrl, out string error)
        {
            error = string.Empty;
            try
            {
                using (WebClient client = CreateClient())
                {
                    PluginUpdateManifest manifest = ReadManifest(client.DownloadString(AddCacheBuster(manifestUrl)));
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

        internal static string DownloadPackage(PluginUpdateManifest manifest, string targetFolder, out string error)
        {
            return DownloadPackage(manifest, targetFolder, null, out error);
        }

        internal static PluginUpdateManifest LoadFromPackageDirectory(string directoryPath, out string error)
        {
            error = string.Empty;
            if (string.IsNullOrWhiteSpace(directoryPath) || !Directory.Exists(directoryPath))
            {
                error = "请先选择存在的本地升级目录。";
                return null;
            }

            string selectedPath = null;
            Version selectedVersion = null;
            foreach (string packagePath in Directory.GetFiles(directoryPath, "DocuLint-*-setup.exe", SearchOption.TopDirectoryOnly))
            {
                Match versionMatch = Regex.Match(
                    Path.GetFileNameWithoutExtension(packagePath),
                    @"(?<!\d)(\d+\.\d+\.\d+(?:\.\d+)?)(?!\d)");
                if (!versionMatch.Success || !Version.TryParse(versionMatch.Groups[1].Value, out Version version))
                {
                    continue;
                }

                if (selectedVersion == null || version > selectedVersion)
                {
                    selectedPath = packagePath;
                    selectedVersion = version;
                }
            }

            if (selectedPath == null)
            {
                error = "目录中未找到 DocuLint-版本号-setup.exe 格式的升级包。";
                return null;
            }

            return new PluginUpdateManifest
            {
                Version = selectedVersion.ToString(),
                PackageUrl = selectedPath,
                PackageFileName = Path.GetFileName(selectedPath),
                Notes = "已从本地升级目录中选择最新安装包。"
            };
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

            foreach (string packageUrl in GetPackageUrls(manifest.PackageUrl))
            {
                try
                {
                    Directory.CreateDirectory(targetFolder);
                    string name = string.IsNullOrWhiteSpace(manifest.PackageFileName)
                        ? Path.GetFileName(new Uri(packageUrl).LocalPath)
                        : manifest.PackageFileName;
                    string target = Path.Combine(targetFolder, string.IsNullOrWhiteSpace(name) ? "DocuLint-update.vsto" : name);
                    DownloadPackageFile(packageUrl, target, progress);

                    if (!VerifySha256(target, manifest.Sha256))
                    {
                        File.Delete(target);
                        error = "安装包 SHA-256 校验失败。";
                        continue;
                    }

                    return target;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                }
            }

            return null;
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

        private static string[] GetPackageUrls(string packageUrl)
        {
            if (packageUrl.StartsWith(GhFastPrefix, StringComparison.OrdinalIgnoreCase))
            {
                string originalUrl = packageUrl.Substring(GhFastPrefix.Length);
                return originalUrl.StartsWith(GitHubReleasePrefix, StringComparison.OrdinalIgnoreCase)
                    ? new[] { packageUrl, originalUrl }
                    : new[] { packageUrl };
            }

            if (packageUrl.StartsWith(GitHubReleasePrefix, StringComparison.OrdinalIgnoreCase))
            {
                return new[]
                {
                    GhFastPrefix + packageUrl,
                    packageUrl
                };
            }

            if (packageUrl.StartsWith(RawGitHubContentPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return new[]
                {
                    JsDelivrContentPrefix + packageUrl.Substring(RawGitHubContentPrefix.Length),
                    packageUrl
                };
            }

            if (packageUrl.StartsWith(JsDelivrContentPrefix, StringComparison.OrdinalIgnoreCase))
            {
                return new[]
                {
                    packageUrl,
                    RawGitHubContentPrefix + packageUrl.Substring(JsDelivrContentPrefix.Length)
                };
            }

            return new[] { packageUrl };
        }

        private static void DownloadPackageFile(string packageUrl, string target, Action<long, long> progress)
        {
            using (WebClient client = CreateClient())
            using (Stream input = client.OpenRead(packageUrl))
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
        }

        private static WebClient CreateClient()
        {
            WebClient client = new TimeoutWebClient();
            client.Headers[HttpRequestHeader.UserAgent] = "DocuLint-UpdateClient/1.0";
            client.Encoding = Encoding.UTF8;
            return client;
        }

        private static string AddCacheBuster(string url)
        {
            return url + (url.IndexOf('?') >= 0 ? "&" : "?") + "v=" + DateTime.UtcNow.Ticks;
        }

        private sealed class TimeoutWebClient : WebClient
        {
            protected override WebRequest GetWebRequest(Uri address)
            {
                WebRequest request = base.GetWebRequest(address);
                request.Timeout = ConnectionTimeoutMilliseconds;

                HttpWebRequest httpRequest = request as HttpWebRequest;
                if (httpRequest != null)
                {
                    httpRequest.ReadWriteTimeout = ReadWriteTimeoutMilliseconds;
                }

                return request;
            }
        }
    }
}
