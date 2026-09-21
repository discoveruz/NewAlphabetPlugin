using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace NewAlphabetPlugin.Tests
{
    [TestClass]
    public class AlphabetRulesTests
    {
        [TestMethod]
        [DataRow('\u02BB', DisplayName = "ʻ U+02BB turned comma (official)")]
        [DataRow('\'', DisplayName = "' U+0027 apostrophe")]
        [DataRow('`', DisplayName = "` U+0060 grave accent")]
        [DataRow('\u2019', DisplayName = "’ U+2019 right single quote")]
        [DataRow('\u2018', DisplayName = "‘ U+2018 left single quote")]
        [DataRow('\u02BC', DisplayName = "ʼ U+02BC modifier apostrophe")]
        [DataRow('\u00B4', DisplayName = "´ U+00B4 acute accent")]
        public void GetReplacement_OAndG_WithEveryApostrophe(char apostrophe)
        {
            Assert.AreEqual("Ö", AlphabetRules.GetReplacement("O" + apostrophe));
            Assert.AreEqual("ö", AlphabetRules.GetReplacement("o" + apostrophe));
            Assert.AreEqual("Ğ", AlphabetRules.GetReplacement("G" + apostrophe));
            Assert.AreEqual("ğ", AlphabetRules.GetReplacement("g" + apostrophe));
        }

        [TestMethod]
        [DataRow("Sh", "Ş")]
        [DataRow("SH", "Ş")]
        [DataRow("sh", "ş")]
        [DataRow("sH", "ş")]
        [DataRow("Ch", "Ç")]
        [DataRow("CH", "Ç")]
        [DataRow("ch", "ç")]
        [DataRow("cH", "ç")]
        public void GetReplacement_ShAndCh_TakeTheCaseOfTheFirstLetter(string oldLetter, string expected)
        {
            Assert.AreEqual(expected, AlphabetRules.GetReplacement(oldLetter));
        }

        [TestMethod]
        [DataRow("oh")]
        [DataRow("gh")]
        [DataRow("s'")]
        [DataRow("c'")]
        [DataRow("Oo")]
        [DataRow("o\"", DisplayName = "o followed by a double quote")]
        [DataRow("o-")]
        [DataRow("s\u02BC", DisplayName = "sʼ, the tutuq in Isʼhoq")]
        [DataRow("Ş", DisplayName = "Ş, already new")]
        [DataRow("O")]
        [DataRow("O\u02BBz", DisplayName = "three characters")]
        [DataRow("")]
        public void GetReplacement_ReturnsNull_ForAnythingElse(string text)
        {
            Assert.IsNull(AlphabetRules.GetReplacement(text));
        }

        [TestMethod]
        public void GetReplacement_ReturnsNull_ForNull()
        {
            Assert.IsNull(AlphabetRules.GetReplacement(null));
        }

        [TestMethod]
        [DataRow("https://shop.uz")]
        [DataRow("http://example.com/shosh")]
        [DataRow("www.chess.com")]
        [DataRow("WWW.Chess.Com")]
        [DataRow("info@school.uz")]
        [DataRow("@shohruh")]
        [DataRow("shosh.uz")]
        [DataRow("kun.uz,")]
        [DataRow("(www.shop.uz)")]
        [DataRow("gazeta.uz/news")]
        public void LooksLikeLink_IsTrue_ForAddresses(string text)
        {
            Assert.IsTrue(AlphabetRules.LooksLikeLink(text));
        }

        [TestMethod]
        [DataRow("Toshkent")]
        [DataRow("Toshkent.Uzbekiston", DisplayName = "missing space after a full stop")]
        [DataRow("Shahar.")]
        [DataRow("kitob.Ruscha")]
        [DataRow("O\u02BBzbekiston")]
        [DataRow("")]
        public void LooksLikeLink_IsFalse_ForOrdinaryWords(string text)
        {
            Assert.IsFalse(AlphabetRules.LooksLikeLink(text));
        }

        [TestMethod]
        [DataRow("Kun.uz'га", 7, false, DisplayName = "Kun.uz'га: the ending after the apostrophe")]
        [DataRow("Kun.uz'га,", 7, false, DisplayName = "Kun.uz'га, with a comma")]
        [DataRow("«kun.uz\u2019дан»", 8, false, DisplayName = "«kun.uz’дан» in quotes")]
        [DataRow("Kun.uz'га", 0, true, DisplayName = "Kun.uz'га: the address itself")]
        [DataRow("shosh.uz'da", 3, true, DisplayName = "shosh.uz'da: sh inside the address")]
        [DataRow("https://сайт.уз", 8, true, DisplayName = "a Cyrillic address")]
        [DataRow("https://example.com/it's-sh", 25, true, DisplayName = "an apostrophe inside an address")]
        [DataRow("Toshkent", 3, false, DisplayName = "not an address")]
        public void IsInLink_LeavesWordEndingsAfterAnAddressOut(string word, int start, bool expected)
        {
            Assert.AreEqual(expected, AlphabetRules.IsInLink(word, start));
        }

        [TestMethod]
        public void LooksLikeLink_IsFalse_ForNull()
        {
            Assert.IsFalse(AlphabetRules.LooksLikeLink(null));
        }

        [TestMethod]
        [DataRow("sh")]
        [DataRow("O\u02BB")]
        [DataRow("g\u02BBo\u02BB", DisplayName = "gʻoʻ")]
        [DataRow("chch")]
        [DataRow("o'sh")]
        public void IsOnlyOldLetters_IsTrue_ForOldLettersAlone(string text)
        {
            Assert.IsTrue(AlphabetRules.IsOnlyOldLetters(text));
        }

        [TestMethod]
        [DataRow("Toshkent")]
        [DataRow("sha")]
        [DataRow("s")]
        [DataRow("sh ")]
        [DataRow("sh\r", DisplayName = "sh and a paragraph mark")]
        [DataRow("")]
        public void IsOnlyOldLetters_IsFalse_ForAnythingMore(string text)
        {
            Assert.IsFalse(AlphabetRules.IsOnlyOldLetters(text));
        }

        [TestMethod]
        public void IsOnlyOldLetters_IsFalse_ForNull()
        {
            Assert.IsFalse(AlphabetRules.IsOnlyOldLetters(null));
        }

        [TestMethod]
        public void WordFindPatterns_ListEveryApostropheAndBothCases()
        {
            CollectionAssert.AreEqual(
                new[] { "[OoGg][\u02BB'`\u2019\u2018\u02BC\u00B4]", "[SsCc][Hh]" },
                AlphabetRules.WordFindPatterns);
        }

        [TestMethod]
        [DataRow(Samples.OldText, Samples.NewText, DisplayName = "every apostrophe and case")]
        [DataRow("O\u02BBzbekiston Respublikasi poytaxti Toshkent shahri", "Özbekiston Respublikasi poytaxti Toşkent şahri")]
        [DataRow("CHORSHANBA", "ÇORŞANBA")]
        [DataRow("Chiroyli qo\u02BBshiq", "Çiroyli qöşiq")]
        [DataRow("G\u02BBafur G\u02BBulom", "Ğafur Ğulom")]
        [DataRow("g\u02BBisht, tog\u02BB, bog\u02BB", "ğişt, toğ, boğ")]
        [DataRow("Is\u02BChoq", "Is\u02BChoq", DisplayName = "the tutuq in Isʼhoq stays")]
        [DataRow("«Birorta g\u02BBo\u02BBdaygan boyvachcha paydo bo\u02BBlsa, kechirmayman!» – prezident",
                 "«Birorta ğödaygan boyvaçça paydo bölsa, keçirmayman!» – prezident", DisplayName = "old letters that touch (gʻoʻ, chch)")]
        [DataRow("Batafsil: https://shop.uz, shikoyat: info@school.uz", "Batafsil: https://shop.uz, şikoyat: info@school.uz")]
        [DataRow("Kun.uz'chi, shop.uz'dagi", "Kun.uz'çi, shop.uz'dagi", DisplayName = "an ending after an address converts, the address stays")]
        public void Rules_ConvertWholeSentences(string oldText, string expected)
        {
            Assert.AreEqual(expected, LikeTheAddIn.ConvertOldLatin(oldText));
        }
    }
}
