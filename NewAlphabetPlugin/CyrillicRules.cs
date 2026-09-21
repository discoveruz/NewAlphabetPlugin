using System;
using System.Collections.Generic;

namespace NewAlphabetPlugin
{
    /// <summary>
    /// Spells Uzbek Cyrillic words in the new Latin alphabet. Follows the rules of the lotin-kirill transliterator
    /// (github.com/diyorbek/lotin-kirill), except that Ш, Ч, Ў and Ғ become the new letters Ş, Ç, Ö and Ğ.
    /// Contains no Word code.
    /// </summary>
    internal static class CyrillicRules
    {
        /// <summary>
        /// Word wildcard pattern that finds a run of Cyrillic characters.
        /// </summary>
        public const string WordFindPattern = "[\u0400-\u04FF]@";

        // Written for Ъ, as in maʼno.
        private const string Tutuq = "\u02BC"; // ʼ

        // Cyrillic letters the rules below look at.
        private const char Ie = 'е';
        private const char O = 'о';
        private const char Es = 'с';
        private const char Tse = 'ц';
        private const char HardSign = 'ъ';
        private const char SoftSign = 'ь';
        private const char ShortU = 'ў';
        private const char GheWithStroke = 'ғ';
        private const char HaWithDescender = 'ҳ';

        // Cyrillic vowels, and the Latin ones that can stand next to a Cyrillic letter.
        private const string Vowels = "аоэеиуўёюя" + "aeiou\u00F6"; // ö

        // Every letter of the Uzbek Cyrillic alphabet in lower case, with its usual spelling.
        // Е, О, С, Ц and Ъ are sometimes spelled differently depending on their neighbours; see Spell.
        private static readonly Dictionary<char, string> Letters = new Dictionary<char, string>
        {
            { 'а', "a" }, { 'б', "b" }, { 'в', "v" }, { 'г', "g" }, { 'д', "d" }, { 'е', "e" },
            { 'ё', "yo" }, { 'ж', "j" }, { 'з', "z" }, { 'и', "i" }, { 'й', "y" }, { 'к', "k" },
            { 'л', "l" }, { 'м', "m" }, { 'н', "n" }, { 'о', "o" }, { 'п', "p" }, { 'р', "r" },
            { 'с', "s" }, { 'т', "t" }, { 'у', "u" }, { 'ф', "f" }, { 'х', "x" }, { 'ц', "s" },
            { 'э', "e" }, { 'ю', "yu" }, { 'я', "ya" }, { 'қ', "q" }, { 'ҳ', "h" }, { 'ь', "" },
            { 'ъ', Tutuq },
            { 'ч', "\u00E7" }, // ç
            { 'ш', "\u015F" }, // ş
            { 'ў', "\u00F6" }, // ö
            { 'ғ', "\u011F" }, // ğ
        };

        // Words the rules would spell wrongly (byudjet, oktyabr), given by how they start so that forms such as
        // бюджетга and октябрда match too. One Latin letter per Cyrillic letter, so each letter keeps its formatting.
        private static readonly KeyValuePair<string, string>[] Exceptions =
        {
            new KeyValuePair<string, string>("бюджет", "budjet"),
            new KeyValuePair<string, string>("октябр", "oktabr"),
            new KeyValuePair<string, string>("сентябр", "sentabr"),
        };

        /// <summary>
        /// True when the text is made only of letters of the Uzbek Cyrillic alphabet. Words with letters Uzbek does
        /// not use, such as the Russian Щ and Ы, are left alone.
        /// </summary>
        public static bool IsUzbekWord(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }
            foreach (char c in text)
            {
                if (!Letters.ContainsKey(char.ToLowerInvariant(c)))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Spells an Uzbek Cyrillic word in the new alphabet, one piece per Cyrillic letter, so that each letter can be
        /// replaced on its own. A piece is empty when the letter is dropped, like Ь.
        /// <paramref name="before"/> and <paramref name="after"/> are the characters around the word, or '\0'.
        /// </summary>
        public static string[] Transliterate(string word, char before, char after)
        {
            var pieces = new string[word.Length];
            bool previousIsYeAfterVowel = false;
            for (int i = SpellException(word, before, pieces); i < word.Length; i++)
            {
                char previous = i > 0 ? word[i - 1] : before;
                char next = i + 1 < word.Length ? word[i + 1] : after;
                bool isYeAfterVowel = false;
                pieces[i] = Spell(word, i, previous, next, previousIsYeAfterVowel, ref isYeAfterVowel);
                previousIsYeAfterVowel = isYeAfterVowel;
            }
            return pieces;
        }

        private static string Spell(string word, int i, char previous, char next, bool previousIsYeAfterVowel, ref bool isYeAfterVowel)
        {
            char letter = word[i];
            char lower = char.ToLowerInvariant(letter);
            bool upper = letter != lower;
            switch (lower)
            {
                case Ie:
                    if (IsSign(previous))
                    {
                        return Case("ye", upper); // съезд → syezd
                    }
                    if (!IsLetter(previous))
                    {
                        return CaseLike("ye", letter, previous, next); // ер → yer
                    }
                    if (IsInRunOfThreeIe(word, i))
                    {
                        return Case("e", upper); // деее → deee
                    }
                    // Mirrors lotin-kirill's non-overlapping search: in аее only the first е follows a vowel.
                    if (IsVowel(previous) && !previousIsYeAfterVowel)
                    {
                        isYeAfterVowel = true;
                        return Case("ye", upper); // Мирзиёев → Mirziyoyev
                    }
                    return Case("e", upper);

                case O:
                    return Case(char.ToLowerInvariant(previous) == SoftSign ? "yo" : "o", upper); // медальон → medalyon

                case Tse:
                    return Case(IsVowel(previous) ? "ts" : "s", upper); // лицей → litsey, цирк → sirk

                case HardSign:
                    char previousLower = char.ToLowerInvariant(previous);
                    if (previousLower == ShortU || previousLower == GheWithStroke || char.ToLowerInvariant(next) == Ie)
                    {
                        return ""; // мўъжиза → möjiza, съезд → syezd
                    }
                    return Tutuq;

                case Es:
                    // Keeps s and h apart, as the old Latin alphabet did: Исҳоқ → Isʼhoq.
                    return char.ToLowerInvariant(next) == HaWithDescender ? Case("s", upper) + Tutuq : Case("s", upper);

                default:
                    string spelling = Letters[lower];
                    return spelling.Length > 1 ? CaseLike(spelling, letter, previous, next) : Case(spelling, upper);
            }
        }

        // Spells a word that starts like one of the exceptions and returns how many letters that took.
        private static int SpellException(string word, char before, string[] pieces)
        {
            if (IsLetter(before))
            {
                return 0;
            }

            foreach (KeyValuePair<string, string> exception in Exceptions)
            {
                string start = exception.Key;
                if (word.Length >= start.Length
                    && string.Compare(word, 0, start, 0, start.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    for (int i = 0; i < start.Length; i++)
                    {
                        pieces[i] = Case(exception.Value[i].ToString(), char.IsUpper(word[i]));
                    }
                    return start.Length;
                }
            }
            return 0;
        }

        // Ё, Ю, Я and a word-initial Е become two letters. In capitals the second one follows the next letter
        // (Ёш → Yoş, ЁШ → YOŞ); with no next letter it follows the previous one (ДЁ → DYO), and alone it is lower (Ё → Yo).
        private static string CaseLike(string spelling, char letter, char previous, char next)
        {
            if (!char.IsUpper(letter))
            {
                return spelling;
            }
            if (IsLetter(next))
            {
                return char.IsUpper(next) ? spelling.ToUpperInvariant() : Capitalize(spelling);
            }
            return IsLetter(previous) ? spelling.ToUpperInvariant() : Capitalize(spelling);
        }

        private static string Case(string spelling, bool upper)
        {
            return upper ? spelling.ToUpperInvariant() : spelling;
        }

        private static string Capitalize(string spelling)
        {
            return char.ToUpperInvariant(spelling[0]) + spelling.Substring(1);
        }

        private static bool IsInRunOfThreeIe(string word, int i)
        {
            int start = i;
            while (start > 0 && char.ToLowerInvariant(word[start - 1]) == Ie)
            {
                start--;
            }
            int end = i + 1;
            while (end < word.Length && char.ToLowerInvariant(word[end]) == Ie)
            {
                end++;
            }
            return end - start >= 3;
        }

        private static bool IsSign(char c)
        {
            char lower = char.ToLowerInvariant(c);
            return lower == HardSign || lower == SoftSign;
        }

        private static bool IsVowel(char c)
        {
            return c != '\0' && Vowels.IndexOf(char.ToLowerInvariant(c)) >= 0;
        }

        // Apostrophes count as letters in Unicode, but not here.
        private static bool IsLetter(char c)
        {
            return char.IsLetter(c) && AlphabetRules.Apostrophes.IndexOf(c) < 0;
        }
    }
}
