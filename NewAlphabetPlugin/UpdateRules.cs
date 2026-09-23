using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Xml;

namespace NewAlphabetPlugin
{
    /// <summary>
    /// The parts of looking for a new version that need neither Word nor the network: when to look, how to read
    /// GitHub's answer and whether that release is newer than the running add-in.
    /// </summary>
    internal static class UpdateRules
    {
        /// <summary>
        /// GitHub's description of the newest release.
        /// </summary>
        public const string LatestReleaseUrl = "https://api.github.com/repos/discoveruz/NewAlphabetPlugin/releases/latest";

        /// <summary>
        /// The installer every release carries. Its name has no version in it, so a link to the newest one never changes.
        /// </summary>
        public const string SetupFileName = "NewAlphabetPlugin-Setup.exe";

        // Only installers from this repository's own releases are downloaded and run.
        private const string DownloadPrefix = "https://github.com/discoveruz/NewAlphabetPlugin/releases/download/";

        private static readonly TimeSpan CheckInterval = TimeSpan.FromDays(1);

        /// <summary>
        /// Word looks for a new version by itself at most once a day. A last check that is in the future, because the
        /// clock was set back, counts as due.
        /// </summary>
        public static bool IsCheckDue(DateTime lastCheckUtc, DateTime nowUtc)
        {
            return lastCheckUtc > nowUtc || nowUtc - lastCheckUtc >= CheckInterval;
        }

        /// <summary>
        /// Reads GitHub's answer about the newest release: the version in its tag, such as v0.3.0, and where to download
        /// its installer. Returns null when the answer has no such version or installer, or the installer is not in this
        /// repository's releases.
        /// </summary>
        public static ReleaseInfo ReadRelease(string json)
        {
            GitHubRelease release;
            try
            {
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
                {
                    release = (GitHubRelease)new DataContractJsonSerializer(typeof(GitHubRelease)).ReadObject(stream);
                }
            }
            catch (SerializationException)
            {
                return null;
            }
            catch (XmlException)
            {
                return null;
            }

            Version version;
            if (release == null || release.TagName == null || release.Assets == null
                || !Version.TryParse(release.TagName.TrimStart('v', 'V'), out version))
            {
                return null;
            }

            foreach (GitHubAsset asset in release.Assets)
            {
                if (asset != null
                    && string.Equals(asset.Name, SetupFileName, StringComparison.OrdinalIgnoreCase)
                    && asset.DownloadUrl != null
                    && asset.DownloadUrl.StartsWith(DownloadPrefix, StringComparison.Ordinal))
                {
                    return new ReleaseInfo(version, asset.DownloadUrl);
                }
            }
            return null;
        }

        /// <summary>
        /// Whether <paramref name="latest"/> is newer than <paramref name="current"/>. Only the first three numbers count,
        /// because the add-in's AssemblyVersion always ends in .0.
        /// </summary>
        public static bool IsNewer(Version latest, Version current)
        {
            return ThreeNumbers(latest) > ThreeNumbers(current);
        }

        private static Version ThreeNumbers(Version version)
        {
            return new Version(version.Major, version.Minor, Math.Max(version.Build, 0));
        }

        [DataContract]
        private sealed class GitHubRelease
        {
            [DataMember(Name = "tag_name")]
            public string TagName { get; set; }

            [DataMember(Name = "assets")]
            public GitHubAsset[] Assets { get; set; }
        }

        [DataContract]
        private sealed class GitHubAsset
        {
            [DataMember(Name = "name")]
            public string Name { get; set; }

            [DataMember(Name = "browser_download_url")]
            public string DownloadUrl { get; set; }
        }
    }

    /// <summary>
    /// A release's version and the address of its installer.
    /// </summary>
    internal sealed class ReleaseInfo
    {
        public ReleaseInfo(Version version, string setupUrl)
        {
            Version = version;
            SetupUrl = setupUrl;
        }

        public Version Version { get; private set; }

        public string SetupUrl { get; private set; }
    }
}
