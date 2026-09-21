namespace NewAlphabetPlugin.Tests
{
    /// <summary>
    /// A text with every old letter in every spelling, shared by the rules tests and the Word tests.
    /// The curly and modifier apostrophes are written as escapes because they look alike; ' and ` are typed as is.
    /// </summary>
    internal static class Samples
    {
        public const string OldText =
            "O\u02BBzbekiston O'zbekiston O`zbekiston O\u2019zbekiston O\u2018zbekiston " +
            "O\u02BCzbekiston O\u00B4zbekiston O\u02BBZBEKISTON " +
            "tog\u02BB tog' bog\u02BB G'ALABA " +
            "Toshkent TOSHKENT Cho\u02BBl CHIROQ " +
            "Is\u02BChoq https://shop.uz info@school.uz";

        public const string NewText =
            "Özbekiston Özbekiston Özbekiston Özbekiston Özbekiston " +
            "Özbekiston Özbekiston ÖZBEKISTON " +
            "toğ toğ boğ ĞALABA " +
            "Toşkent TOŞKENT Çöl ÇIROQ " +
            "Is\u02BChoq https://shop.uz info@school.uz";

        // 8 × Oʻ, 4 × Gʻ/gʻ, then sh, SH, Ch, oʻ and CH.
        public const int OldLetterCount = 17;
    }
}
