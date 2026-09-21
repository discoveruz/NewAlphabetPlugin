using System;
using System.Text.RegularExpressions;

namespace NewAlphabetPlugin
{
    /// <summary>
    /// Which old Uzbek Latin letters become which new letters. Contains no Word code.
    /// </summary>
    internal static class AlphabetRules
    {
        /// <summary>
        /// Every character people type as the ʻ in Oʻ and Gʻ.
        /// </summary>
        public const string Apostrophes =
            "\u02BB" + // ʻ modifier letter turned comma (the official one)
            "\u0027" + // ' apostrophe
            "\u0060" + // ` grave accent
            "\u2019" + // ’ right single quotation mark
            "\u2018" + // ‘ left single quotation mark
            "\u02BC" + // ʼ modifier letter apostrophe
            "\u00B4";  // ´ acute accent

        /// <summary>
        /// Word wildcard patterns that find every old two-character letter.
        /// Wildcard search is case-sensitive, so both cases are listed.
        /// </summary>
        public static readonly string[] WordFindPatterns =
        {
            "[OoGg][" + Apostrophes + "]",
            "[SsCc][Hh]",
        };

        // The Cyrillic domains are .уз and .рф.
        private static readonly Regex LinkPattern = new Regex(
            @"://|www\.|@|\.(uz|com|ru|org|net|\u0443\u0437|\u0440\u0444)\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // What follows an address after an apostrophe when it is a word ending, as in Kun.uz'ga: letters, then
        // perhaps punctuation such as a comma or a closing quote.
        private static readonly Regex WordEnding = new Regex(@"^\p{L}+\p{P}*\z", RegexOptions.CultureInvariant);

        // Word wildcard classes such as [OoGg] mean the same in a .NET regex, so the Word patterns are reused.
        private static readonly Regex OnlyOldLettersPattern = new Regex(
            "^(?:" + string.Join("|", WordFindPatterns) + ")+\\z",
            RegexOptions.CultureInvariant);

        /// <summary>
        /// Returns the new letter for an old two-character letter such as "Sh" or "oʻ",
        /// or null when the text is not one. The first character decides upper or lower case.
        /// </summary>
        public static string GetReplacement(string oldLetter)
        {
            if (oldLetter == null || oldLetter.Length != 2)
            {
                return null;
            }

            char second = oldLetter[1];
            bool isApostrophe = Apostrophes.IndexOf(second) >= 0;
            bool isH = second == 'h' || second == 'H';

            switch (oldLetter[0])
            {
                case 'O': return isApostrophe ? "\u00D6" : null; // Ö
                case 'o': return isApostrophe ? "\u00F6" : null; // ö
                case 'G': return isApostrophe ? "\u011E" : null; // Ğ
                case 'g': return isApostrophe ? "\u011F" : null; // ğ
                case 'S': return isH ? "\u015E" : null;          // Ş
                case 's': return isH ? "\u015F" : null;          // ş
                case 'C': return isH ? "\u00C7" : null;          // Ç
                case 'c': return isH ? "\u00E7" : null;          // ç
                default: return null;
            }
        }

        /// <summary>
        /// True when the text is made of old two-character letters and nothing else, such as "sh" or "gʻoʻ".
        /// </summary>
        public static bool IsOnlyOldLetters(string text)
        {
            return !string.IsNullOrEmpty(text) && OnlyOldLettersPattern.IsMatch(text);
        }

        /// <summary>
        /// True when the text contains a web address, e-mail address or domain name, which must stay as typed.
        /// </summary>
        public static bool LooksLikeLink(string text)
        {
            return !string.IsNullOrEmpty(text) && LinkPattern.IsMatch(text);
        }

        /// <summary>
        /// True when the characters of <paramref name="word"/> (text between spaces) from <paramref name="start"/> on
        /// belong to a web address, e-mail address or domain name. A word ending joined to an address with an
        /// apostrophe, as in Kun.uz'ga, is not part of the address.
        /// </summary>
        public static bool IsInLink(string word, int start)
        {
            if (!LooksLikeLink(word))
            {
                return false;
            }

            start = Math.Min(start, word.Length);
            int apostrophe = start > 0 ? word.LastIndexOfAny(Apostrophes.ToCharArray(), start - 1) : -1;
            if (apostrophe < 0)
            {
                return true;
            }
            bool isWordEnding = LooksLikeLink(word.Substring(0, apostrophe)) && WordEnding.IsMatch(word.Substring(apostrophe + 1));
            return !isWordEnding;
        }
    }
}
