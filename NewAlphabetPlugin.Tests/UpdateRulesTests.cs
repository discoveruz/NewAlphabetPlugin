using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace NewAlphabetPlugin.Tests
{
    [TestClass]
    public class UpdateRulesTests
    {
        private const string SetupUrl =
            "https://github.com/discoveruz/NewAlphabetPlugin/releases/download/v0.3.0/NewAlphabetPlugin-Setup.exe";

        // Shaped like GitHub's answer, with some of the fields the add-in does not use.
        private static string ReleaseJson(string tag, string assetName, string assetUrl)
        {
            return @"{
  ""url"": ""https://api.github.com/repos/discoveruz/NewAlphabetPlugin/releases/1"",
  ""tag_name"": """ + tag + @""",
  ""name"": ""v0.3.0"",
  ""draft"": false,
  ""author"": { ""login"": ""discoveruz"", ""id"": 1 },
  ""body"": ""* Fix \""quotes\"" by @someone in #5"",
  ""assets"": [
    {
      ""name"": ""notes.txt"",
      ""size"": 12,
      ""browser_download_url"": ""https://github.com/discoveruz/NewAlphabetPlugin/releases/download/v0.3.0/notes.txt""
    },
    {
      ""name"": """ + assetName + @""",
      ""size"": 2400000,
      ""uploader"": { ""login"": ""github-actions[bot]"" },
      ""browser_download_url"": """ + assetUrl + @"""
    }
  ]
}";
        }

        [TestMethod]
        public void ReadRelease_FindsTheVersionAndTheInstaller()
        {
            ReleaseInfo release = UpdateRules.ReadRelease(ReleaseJson("v0.3.0", "NewAlphabetPlugin-Setup.exe", SetupUrl));

            Assert.IsNotNull(release);
            Assert.AreEqual(new Version(0, 3, 0), release.Version);
            Assert.AreEqual(SetupUrl, release.SetupUrl);
        }

        [TestMethod]
        public void ReadRelease_WithoutAnInstaller_IsNull()
        {
            // v0.1.0 and v0.2.0 carry a zipped ClickOnce setup instead.
            string zipUrl = "https://github.com/discoveruz/NewAlphabetPlugin/releases/download/v0.2.0/NewAlphabetPlugin-v0.2.0.zip";

            Assert.IsNull(UpdateRules.ReadRelease(ReleaseJson("v0.2.0", "NewAlphabetPlugin-v0.2.0.zip", zipUrl)));
        }

        [TestMethod]
        [DataRow("https://github.com/someone-else/NewAlphabetPlugin/releases/download/v0.3.0/NewAlphabetPlugin-Setup.exe")]
        [DataRow("http://github.com/discoveruz/NewAlphabetPlugin/releases/download/v0.3.0/NewAlphabetPlugin-Setup.exe")]
        [DataRow("https://example.com/discoveruz/NewAlphabetPlugin/releases/download/v0.3.0/NewAlphabetPlugin-Setup.exe")]
        public void ReadRelease_AnInstallerFromElsewhere_IsNull(string url)
        {
            Assert.IsNull(UpdateRules.ReadRelease(ReleaseJson("v0.3.0", "NewAlphabetPlugin-Setup.exe", url)));
        }

        [TestMethod]
        [DataRow("latest")]
        [DataRow("")]
        [DataRow("v")]
        public void ReadRelease_ATagThatIsNotAVersion_IsNull(string tag)
        {
            Assert.IsNull(UpdateRules.ReadRelease(ReleaseJson(tag, "NewAlphabetPlugin-Setup.exe", SetupUrl)));
        }

        [TestMethod]
        [DataRow("")]
        [DataRow("{")]
        [DataRow("<html>rate limited</html>")]
        [DataRow("{\"message\": \"Not Found\"}")]
        public void ReadRelease_AnAnswerThatIsNotARelease_IsNull(string json)
        {
            Assert.IsNull(UpdateRules.ReadRelease(json));
        }

        [TestMethod]
        [DataRow("0.3.0", "0.2.0.0", true)]
        [DataRow("0.2.1", "0.2.0.0", true)]
        [DataRow("1.0.0", "0.9.9.0", true)]
        [DataRow("0.10.0", "0.9.0.0", true, DisplayName = "0.10.0 is newer than 0.9.0")]
        [DataRow("0.2.0", "0.2.0.0", false, DisplayName = "the same version")]
        [DataRow("0.2.0", "0.3.0.0", false, DisplayName = "an older version")]
        [DataRow("0.2", "0.2.0.0", false, DisplayName = "a two-number tag")]
        public void IsNewer_ComparesTheFirstThreeNumbers(string latest, string current, bool expected)
        {
            Assert.AreEqual(expected, UpdateRules.IsNewer(Version.Parse(latest), Version.Parse(current)));
        }

        [TestMethod]
        [DataRow(4551, DisplayName = "4551 An Application Control policy has blocked this file")]
        [DataRow(4556, DisplayName = "4556 blocked: malicious binary reputation")]
        [DataRow(4557, DisplayName = "4557 blocked: potentially unwanted application")]
        [DataRow(4558, DisplayName = "4558 blocked: dangerous file extension from the web")]
        [DataRow(4559, DisplayName = "4559 blocked: unable to contact reputation service")]
        public void InstallFailureMessage_WindowsBlockedTheInstaller(int errorCode)
        {
            string message = UpdateRules.InstallFailureMessage(new Win32Exception(errorCode));

            StringAssert.StartsWith(message, "Windows yangi versiyaning örnatuvçisini hozirça töxtatdi");
            StringAssert.Contains(message, "«Maʼlumot» tugmasi orqali qayta urinib köring");
        }

        [TestMethod]
        public void InstallFailureMessage_TheDownloadFailed()
        {
            // The download runs in a task, so its error arrives wrapped.
            var error = new AggregateException(new WebException("The remote name could not be resolved: 'github.com'"));

            StringAssert.StartsWith(UpdateRules.InstallFailureMessage(error), "Yangi versiyani yuklab olib bölmadi.");
        }

        [TestMethod]
        public void InstallFailureMessage_SavingTheInstallerFailed()
        {
            var error = new AggregateException(new IOException("There is not enough space on the disk."));

            StringAssert.StartsWith(UpdateRules.InstallFailureMessage(error), "Yangi versiyani yuklab olib bölmadi.");
        }

        [TestMethod]
        [DataRow(2, DisplayName = "2 file not found")]
        [DataRow(1223, DisplayName = "1223 the user cancelled")]
        public void InstallFailureMessage_AnythingElse(int errorCode)
        {
            string message = UpdateRules.InstallFailureMessage(new Win32Exception(errorCode));

            StringAssert.StartsWith(message, "Yangi versiyani örnatib bölmadi.");
            Assert.IsFalse(message.Contains(new Win32Exception(errorCode).Message), "Windows' own English text is left out");
        }

        [TestMethod]
        public void IsCheckDue_OncePerDay()
        {
            var now = new DateTime(2026, 9, 23, 12, 0, 0, DateTimeKind.Utc);

            Assert.IsTrue(UpdateRules.IsCheckDue(DateTime.MinValue, now), "never checked");
            Assert.IsFalse(UpdateRules.IsCheckDue(now.AddHours(-1), now), "an hour ago");
            Assert.IsFalse(UpdateRules.IsCheckDue(now.AddHours(-23), now), "23 hours ago");
            Assert.IsTrue(UpdateRules.IsCheckDue(now.AddHours(-24), now), "a day ago");
            Assert.IsTrue(UpdateRules.IsCheckDue(now.AddDays(3), now), "in the future, after the clock was set back");
        }
    }
}
