using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace NewAlphabetPlugin
{
    /// <summary>
    /// Looks on GitHub for a newer release and offers to install it. Word looks by itself when it starts, at most once
    /// a day and in the background; the Maʼlumot button looks straight away.
    /// </summary>
    internal static class Updater
    {
        private const int BackgroundTimeout = 30000;
        private const int ManualTimeout = 5000;

        private static Version CurrentVersion
        {
            get { return typeof(Updater).Assembly.GetName().Version; }
        }

        // GitHub refuses requests without a user agent.
        private static string UserAgent
        {
            get { return "NewAlphabetPlugin/" + CurrentVersion.ToString(3); }
        }

        // The installer's "Look for new versions" choice. A copy installed before the choice existed has no value and
        // looks, as it did then; when the choice cannot be read, it does not look.
        private static bool AutomaticCheckChosen
        {
            get
            {
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\NewAlphabetPlugin"))
                    {
                        object value = key == null ? null : key.GetValue("CheckForUpdates");
                        return value == null || (value is int && (int)value != 0);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                    return false;
                }
            }
        }

        /// <summary>
        /// Looks for a new version in the background when a day has passed since the last look, unless it was turned
        /// off during installation. Call it on Word's thread, where the offer is then shown.
        /// </summary>
        public static void CheckInBackground()
        {
#if DEBUG
            // A development build keeps the repository's AssemblyVersion, so every release would look new to it.
            return;
#else
            try
            {
                if (!AutomaticCheckChosen || !UpdateRules.IsCheckDue(UserSettings.LastUpdateCheck, DateTime.UtcNow))
                {
                    return;
                }
                UserSettings.LastUpdateCheck = DateTime.UtcNow;

                string skipped = UserSettings.SkippedUpdate;
                SynchronizationContext wordThread = WordThread();
                Task.Run(() => FetchLatest(BackgroundTimeout)).ContinueWith(task =>
                {
                    if (task.IsFaulted)
                    {
                        Debug.WriteLine(task.Exception);
                        return;
                    }
                    ReleaseInfo release = task.Result;
                    if (release != null && UpdateRules.IsNewer(release.Version, CurrentVersion)
                        && release.Version.ToString(3) != skipped)
                    {
                        wordThread.Post(_ => Offer(release), null);
                    }
                });
            }
            catch (Exception ex)
            {
                // Starting Word must never fail because of the update check.
                Debug.WriteLine(ex);
            }
#endif
        }

        /// <summary>
        /// The newest release when it is newer than this add-in; null when it is not, or GitHub cannot be reached within
        /// a few seconds. It waits for GitHub, so only call it when the user has just asked.
        /// </summary>
        public static ReleaseInfo FindNewerRelease()
        {
            try
            {
                ReleaseInfo release = FetchLatest(ManualTimeout);
                return release != null && UpdateRules.IsNewer(release.Version, CurrentVersion) ? release : null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                return null;
            }
        }

        /// <summary>
        /// The question that offers a release, such as "Yangi versiya çiqdi: 0.3.0 (sizda 0.2.0). Hozir örnatilsinmi?".
        /// </summary>
        public static string OfferText(ReleaseInfo release)
        {
            return string.Format(
                "Yangi versiya çiqdi: {0} (sizda {1}).\nHozir örnatilsinmi?",
                release.Version.ToString(3),
                CurrentVersion.ToString(3));
        }

        /// <summary>
        /// Downloads the release's installer in the background, then starts it. Call it on Word's thread.
        /// </summary>
        public static void Install(ReleaseInfo release)
        {
            SynchronizationContext wordThread = WordThread();
            SetStatusBar("Yangi versiya yuklab olinmoqda...");
            Task.Run(() => Download(release)).ContinueWith(task => wordThread.Post(_ =>
            {
                SetStatusBar("");
                try
                {
                    // The installer closes Word if it has to; the new version runs from the next time Word starts.
                    Process.Start(task.Result);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                    MessageBox.Show(
                        UpdateRules.InstallFailureMessage(ex),
                        AlphabetRibbon.MessageTitle,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }, null));
        }

        // Asked once for each new version: after a no, Word does not offer that version again by itself.
        private static void Offer(ReleaseInfo release)
        {
            try
            {
                DialogResult answer = MessageBox.Show(
                    OfferText(release),
                    AlphabetRibbon.MessageTitle,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);
                if (answer == DialogResult.Yes)
                {
                    Install(release);
                }
                else
                {
                    UserSettings.SkippedUpdate = release.Version.ToString(3);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private static ReleaseInfo FetchLatest(int timeout)
        {
            UseTls12();
            var request = (HttpWebRequest)WebRequest.Create(UpdateRules.LatestReleaseUrl);
            request.UserAgent = UserAgent;
            request.Accept = "application/vnd.github+json";
            request.Timeout = timeout;
            request.ReadWriteTimeout = timeout;
            using (var response = (HttpWebResponse)request.GetResponse())
            using (var reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8))
            {
                return UpdateRules.ReadRelease(reader.ReadToEnd());
            }
        }

        private static string Download(ReleaseInfo release)
        {
            UseTls12();
            string path = Path.Combine(
                Path.GetTempPath(),
                "NewAlphabetPlugin-Setup-" + release.Version.ToString(3) + ".exe");
            using (var client = new WebClient())
            {
                client.Headers[HttpRequestHeader.UserAgent] = UserAgent;
                client.DownloadFile(release.SetupUrl, path);
            }
            return path;
        }

        // Word's process can still be limited to TLS 1.0, which GitHub refuses. When it leaves the choice to Windows,
        // which also allows newer versions, it is left alone.
        private static void UseTls12()
        {
            if (ServicePointManager.SecurityProtocol != SecurityProtocolType.SystemDefault)
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            }
        }

        // Background work shows its results through this. Call it on Word's thread.
        private static SynchronizationContext WordThread()
        {
            return SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();
        }

        private static void SetStatusBar(string text)
        {
            try
            {
                Globals.ThisAddIn.Application.StatusBar = text;
            }
            catch (COMException ex)
            {
                Debug.WriteLine(ex);
            }
        }
    }
}
