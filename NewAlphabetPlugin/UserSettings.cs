using System;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace NewAlphabetPlugin
{
    /// <summary>
    /// The choices the user makes on the ribbon, kept between Word sessions. When the settings file cannot be read
    /// the defaults are used, and when it cannot be written the choice still applies until Word closes.
    /// </summary>
    internal static class UserSettings
    {
        /// <summary>
        /// The highlight colour for previews; turquoise when nothing valid is saved.
        /// </summary>
        public static Word.WdColorIndex PreviewColor
        {
            get
            {
                var saved = (Word.WdColorIndex)Read(settings => settings.PreviewColor, (int)AlphabetConverter.DefaultPreviewColor);
                return HighlightColors.All.Any(color => color.Index == saved) ? saved : AlphabetConverter.DefaultPreviewColor;
            }
            set { Write(settings => settings.PreviewColor = (int)value); }
        }

        /// <summary>
        /// Whether previews look for old Latin letters.
        /// </summary>
        public static bool ConvertOldLatin
        {
            get { return Read(settings => settings.ConvertOldLatin, true); }
            set { Write(settings => settings.ConvertOldLatin = value); }
        }

        /// <summary>
        /// Whether previews look for Uzbek Cyrillic words.
        /// </summary>
        public static bool ConvertCyrillic
        {
            get { return Read(settings => settings.ConvertCyrillic, true); }
            set { Write(settings => settings.ConvertCyrillic = value); }
        }

        private static T Read<T>(Func<Properties.Settings, T> read, T fallback)
        {
            try
            {
                UpgradeAfterWordUpdate();
                return read(Properties.Settings.Default);
            }
            catch (ConfigurationErrorsException ex)
            {
                Debug.WriteLine(ex);
                return fallback;
            }
        }

        private static void Write(Action<Properties.Settings> write)
        {
            try
            {
                write(Properties.Settings.Default);
                Properties.Settings.Default.Save();
            }
            catch (ConfigurationErrorsException ex)
            {
                Debug.WriteLine(ex);
            }
        }

        // Settings are stored per Word version, so after a Word update they start empty; carry the previous
        // version's settings over once.
        private static void UpgradeAfterWordUpdate()
        {
            Properties.Settings settings = Properties.Settings.Default;
            if (settings.UpgradeRequired)
            {
                settings.Upgrade();
                settings.UpgradeRequired = false;
                settings.Save();
            }
        }
    }
}
