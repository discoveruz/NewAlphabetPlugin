using System.Text.RegularExpressions;

namespace NewAlphabetPlugin.Tests
{
    /// <summary>
    /// Runs the add-in's rules over plain text without Word, the way the add-in applies them to a document:
    /// web addresses stay as typed. AlphabetConverterWordTests runs the real Word search.
    /// </summary>
    internal static class LikeTheAddIn
    {
        // Word wildcard classes such as [OoGg] mean the same in .NET regex, so the add-in's own patterns are used.
        private static readonly Regex OldLetter = new Regex(string.Join("|", AlphabetRules.WordFindPatterns));

        // The Word pattern as a .NET regex: Word writes "one or more" as @.
        private static readonly Regex CyrillicRun = new Regex(CyrillicRules.WordFindPattern.Replace("]@", "]+"));

        private static readonly Regex Token = new Regex(@"\S+");

        public static string ConvertOldLatin(string text)
        {
            return Token.Replace(text, token => OldLetter.Replace(token.Value, match =>
                AlphabetRules.IsInLink(token.Value, match.Index) ? match.Value : AlphabetRules.GetReplacement(match.Value)));
        }

        public static string ConvertCyrillic(string text)
        {
            return Token.Replace(text, token => CyrillicRun.Replace(token.Value, run =>
                AlphabetRules.IsInLink(token.Value, run.Index) ? run.Value : Transliterate(token.Value, run)));
        }

        private static string Transliterate(string token, Match run)
        {
            if (!CyrillicRules.IsUzbekWord(run.Value))
            {
                return run.Value;
            }
            int end = run.Index + run.Length;
            char before = run.Index > 0 ? token[run.Index - 1] : '\0';
            char after = end < token.Length ? token[end] : '\0';
            return string.Concat(CyrillicRules.Transliterate(run.Value, before, after));
        }
    }
}
