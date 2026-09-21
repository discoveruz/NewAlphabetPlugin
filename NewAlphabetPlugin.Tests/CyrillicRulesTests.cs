using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace NewAlphabetPlugin.Tests
{
    [TestClass]
    public class CyrillicRulesTests
    {
        private const string UzbekLetters =
            "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЪЬЭЮЯЎҚҒҲ" +
            "абвгдеёжзийклмнопрстуфхцчшъьэюяўқғҳ";

        // ---- the alphabet ----

        [TestMethod]
        public void IsUzbekWord_AcceptsTheUzbekCyrillicLetters_AndNoOtherCyrillicCharacter()
        {
            for (int code = 0x0400; code <= 0x04FF; code++)
            {
                string character = ((char)code).ToString();
                Assert.AreEqual(UzbekLetters.Contains(character), CyrillicRules.IsUzbekWord(character), "U+" + code.ToString("X4"));
            }
        }

        [TestMethod]
        [DataRow("Щука", DisplayName = "Russian Щ")]
        [DataRow("мы", DisplayName = "Russian Ы")]
        [DataRow("Україна", DisplayName = "Ukrainian І and Ї")]
        [DataRow("Toshkent", DisplayName = "Latin")]
        [DataRow("Тош кент", DisplayName = "a space")]
        [DataRow("")]
        public void IsUzbekWord_IsFalse_ForAnythingElse(string text)
        {
            Assert.IsFalse(CyrillicRules.IsUzbekWord(text));
        }

        [TestMethod]
        public void IsUzbekWord_IsFalse_ForNull()
        {
            Assert.IsFalse(CyrillicRules.IsUzbekWord(null));
        }

        [TestMethod]
        public void Transliterate_SpellsEveryLetterWithoutCyrillic()
        {
            foreach (char letter in UzbekLetters)
            {
                string latin = string.Concat(CyrillicRules.Transliterate(letter.ToString(), '\0', '\0'));
                Assert.IsFalse(latin.Any(c => c >= '\u0400' && c <= '\u04FF'), letter + " became " + latin);
            }
        }

        [TestMethod]
        public void Transliterate_UsesTheSameNewLettersAsTheOldLatinRules()
        {
            string oldLatin = string.Concat(new[] { "Sh", "Ch", "O\u02BB", "G\u02BB", "sh", "ch", "o\u02BB", "g\u02BB" }
                .Select(AlphabetRules.GetReplacement));

            Assert.AreEqual(oldLatin, LikeTheAddIn.ConvertCyrillic("Ш Ч Ў Ғ ш ч ў ғ").Replace(" ", ""));
        }

        [TestMethod]
        public void Transliterate_ReturnsOnePiecePerLetter_AndAnEmptyPieceForADroppedOne()
        {
            CollectionAssert.AreEqual(
                new[] { "o", "k", "t", "a", "b", "r", "" },
                CyrillicRules.Transliterate("октябрь", '\0', '\0'));
            CollectionAssert.AreEqual(
                new[] { "S", "yo", "m", "k", "a" },
                CyrillicRules.Transliterate("Сёмка", '\0', '\0'));
        }

        // ---- the cases lotin-kirill tests, with the new letters ----

        [TestMethod]
        [DataRow("дае", "daye")]
        [DataRow("даЕ", "daYE")]
        [DataRow("дАе", "dAye")]
        [DataRow("дое", "doye")]
        [DataRow("дэе", "deye")]
        [DataRow("дее", "deye")]
        [DataRow("деее", "deee")]
        [DataRow("дееее", "deeee")]
        [DataRow("дие", "diye")]
        [DataRow("дуе", "duye")]
        [DataRow("ёе", "yoye")]
        [DataRow("юе", "yuye")]
        [DataRow("яе", "yaye")]
        public void Ye_AfterAVowel_InAnyCase_AndRepeated(string cyrillic, string latin)
        {
            Assert.AreEqual(latin, LikeTheAddIn.ConvertCyrillic(cyrillic));
            Assert.AreEqual(latin.ToUpperInvariant(), LikeTheAddIn.ConvertCyrillic(cyrillic.ToUpperInvariant()));
            Assert.AreEqual(latin + latin, LikeTheAddIn.ConvertCyrillic(cyrillic + cyrillic));
            Assert.AreEqual((latin + latin).ToUpperInvariant(), LikeTheAddIn.ConvertCyrillic((cyrillic + cyrillic).ToUpperInvariant()));
        }

        [TestMethod]
        [DataRow("Е", "Ye")]
        [DataRow("е", "ye")]
        [DataRow("ЕД", "YED")]
        [DataRow("ДЕ", "DE")]
        [DataRow("ДЬЕ", "DYE")]
        [DataRow("ДЪЕ", "DYE")]
        [DataRow("Ед", "Yed")]
        [DataRow("дьЕ", "dYE")]
        [DataRow("дъЕ", "dYE")]
        [DataRow("Ё", "Yo")]
        [DataRow("ё", "yo")]
        [DataRow("ЁД", "YOD")]
        [DataRow("ДЁ", "DYO")]
        [DataRow("Ёд", "Yod")]
        [DataRow("дЁ", "dYO")]
        [DataRow("ьо", "yo")]
        [DataRow("ьО", "YO")]
        [DataRow("Ьо", "yo")]
        [DataRow("ЬО", "YO")]
        [DataRow("Ю", "Yu")]
        [DataRow("ю", "yu")]
        [DataRow("ЮБ", "YUB")]
        [DataRow("БЮ", "BYU")]
        [DataRow("Юб", "Yub")]
        [DataRow("бЮ", "bYU")]
        [DataRow("Я", "Ya")]
        [DataRow("я", "ya")]
        [DataRow("ЯБ", "YAB")]
        [DataRow("БЯ", "BYA")]
        [DataRow("Яб", "Yab")]
        [DataRow("бЯ", "bYA")]
        public void TwoLetterSpellings_FollowTheCaseAroundThem(string cyrillic, string latin)
        {
            Assert.AreEqual(latin, LikeTheAddIn.ConvertCyrillic(cyrillic));
        }

        [TestMethod]
        [DataRow("Ч", "Ç")]
        [DataRow("ч", "ç")]
        [DataRow("ЧА", "ÇA")]
        [DataRow("АЧ", "AÇ")]
        [DataRow("Ча", "Ça")]
        [DataRow("аЧ", "aÇ")]
        [DataRow("Ш", "Ş")]
        [DataRow("ш", "ş")]
        [DataRow("ША", "ŞA")]
        [DataRow("АШ", "AŞ")]
        [DataRow("Ша", "Şa")]
        [DataRow("аШ", "aŞ")]
        [DataRow("Ў", "Ö")]
        [DataRow("ў", "ö")]
        [DataRow("Ўъ", "Ö", DisplayName = "Ўъ: the hard sign after Ў is dropped")]
        [DataRow("ўъ", "ö")]
        [DataRow("Ғ", "Ğ")]
        [DataRow("ғ", "ğ")]
        [DataRow("Ғъ", "Ğ")]
        [DataRow("ғъ", "ğ")]
        public void NewLetters_KeepTheCaseOfTheOldOne(string cyrillic, string latin)
        {
            Assert.AreEqual(latin, LikeTheAddIn.ConvertCyrillic(cyrillic));
        }

        [TestMethod]
        [DataRow("ц", "s")]
        [DataRow("цц", "ss")]
        [DataRow("Ц", "S")]
        [DataRow("ЦЦ", "SS")]
        [DataRow("пица", "pitsa")]
        [DataRow("аца", "atsa")]
        [DataRow("aцi", "atsi", DisplayName = "aцi with Latin a and i around it")]
        [DataRow("пицца", "pitssa")]
        [DataRow("iцц", "itss")]
        [DataRow("iцЦ", "itsS")]
        [DataRow("aЦц", "aTSs")]
        [DataRow("aЦЦ", "aTSS")]
        public void Tse_IsTsAfterAVowel_AndSOtherwise(string cyrillic, string latin)
        {
            Assert.AreEqual(latin, LikeTheAddIn.ConvertCyrillic(cyrillic));
        }

        [TestMethod]
        [DataRow("СҲ", "S\u02BCH")]
        [DataRow("Сҳ", "S\u02BCh")]
        [DataRow("сҲ", "s\u02BCH")]
        [DataRow("сҳ", "s\u02BCh")]
        public void SAndH_StayApartWithATutuq(string cyrillic, string latin)
        {
            Assert.AreEqual(latin, LikeTheAddIn.ConvertCyrillic(cyrillic));
        }

        [TestMethod]
        [DataRow("бюджет", "budjet")]
        [DataRow("Бюджетга", "Budjetga")]
        [DataRow("БЮДЖЕТ", "BUDJET")]
        [DataRow("октябрь", "oktabr")]
        [DataRow("октябрда", "oktabrda")]
        [DataRow("Октябрда", "Oktabrda")]
        [DataRow("ОКТЯБРДАГИ", "OKTABRDAGI")]
        [DataRow("сентябрь", "sentabr")]
        [DataRow("ноябрь", "noyabr", DisplayName = "ноябрь follows the rules")]
        public void Exceptions_AlsoMatchLongerForms(string cyrillic, string latin)
        {
            Assert.AreEqual(latin, LikeTheAddIn.ConvertCyrillic(cyrillic));
        }

        // ---- words and sentences ----

        [TestMethod]
        [DataRow("Ўзбекистон Республикаси пойтахти Тошкент шаҳри", "Özbekiston Respublikasi poytaxti Toşkent şahri")]
        [DataRow("ЎЗБЕКИСТОН", "ÖZBEKISTON")]
        [DataRow("ғишт, тоғ, боғ", "ğişt, toğ, boğ")]
        [DataRow("Чиройли қўшиқ", "Çiroyli qöşiq")]
        [DataRow("Ғафур Ғулом", "Ğafur Ğulom")]
        [DataRow("«Бирорта ғўдайган бойвачча пайдо бўлса, кечирмайман!» – президент",
                 "«Birorta ğödaygan boyvaçça paydo bölsa, keçirmayman!» – prezident", DisplayName = "the sentence from the old-letters bug")]
        [DataRow("Мирзиёев", "Mirziyoyev")]
        [DataRow("съезд, объект", "syezd, obyekt")]
        [DataRow("маъно, шеър", "ma\u02BCno, şe\u02BCr", DisplayName = "маъно, шеър: Ъ becomes the tutuq")]
        [DataRow("мўъжиза", "möjiza")]
        [DataRow("Исҳоқ", "Is\u02BChoq")]
        [DataRow("«Ер» ва ер-сув", "«Yer» va yer-suv")]
        [DataRow("бор-йўғи", "bor-yöği")]
        [DataRow("цирк, лицей, пицца", "sirk, litsey, pitssa")]
        [DataRow("медальон", "medalyon")]
        [DataRow("Ёш, ЁШ", "Yoş, YOŞ")]
        [DataRow("XXI аср", "XXI asr", DisplayName = "Latin text around Cyrillic stays")]
        [DataRow("Щука ва Україна", "Щука va Україна", DisplayName = "words with letters Uzbek does not use stay")]
        [DataRow("Батафсил: президент.уз, почта: info@мактаб.уз", "Batafsil: президент.уз, poçta: info@мактаб.уз", DisplayName = "links stay")]
        [DataRow("Kun.uz'га мурожаат қилган", "Kun.uz'ga murojaat qilgan", DisplayName = "Kun.uz'га: the ending after an address converts")]
        [DataRow("kun.uz\u2019дан, Gazeta.uz'нинг «Telegram'даги»", "kun.uz\u2019dan, Gazeta.uz'ning «Telegram'dagi»", DisplayName = "more endings after addresses")]
        [DataRow("https://сайт.уз/it's-сайт", "https://сайт.уз/it's-сайт", DisplayName = "an address with an apostrophe inside stays")]
        public void WholeSentences(string cyrillic, string latin)
        {
            Assert.AreEqual(latin, LikeTheAddIn.ConvertCyrillic(cyrillic));
        }

        // lotin-kirill tests the same news texts in Cyrillic and in old Latin. Converting the Cyrillic one with both
        // scripts ticked must give what the old Latin one gives. The Cyrillic long text quotes some old Latin, which is
        // why its old letters are converted too. Two differences between the texts are not about the rules: they write
        // the tutuq with different apostrophes, and the Latin one spells the months sentyabr and oktyabr, where the
        // add-in uses the official sentabr and oktabr. One difference is on purpose: lotin-kirill half converts the
        // Kyrgyz name Чыӈгыз, while the add-in leaves words with letters Uzbek does not use as they are.
        [TestMethod]
        [DataRow("cyrillicText.txt", "latinText.txt")]
        [DataRow("cyrillicLongText.txt", "latinLongText.txt")]
        public void SampleTexts_MatchTheOldLatinVersionConvertedToTheNewAlphabet(string cyrillicFile, string latinFile)
        {
            string expected = SameApostrophes(LikeTheAddIn.ConvertOldLatin(ReadFixture(latinFile)))
                .Replace("sentyabr", "sentabr").Replace("oktyabr", "oktabr")
                .Replace("Çыӈgыz", "Чыӈгыз");
            string actual = SameApostrophes(LikeTheAddIn.ConvertCyrillic(LikeTheAddIn.ConvertOldLatin(ReadFixture(cyrillicFile))));

            string[] expectedWords = expected.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            string[] actualWords = actual.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
            Assert.AreEqual(expectedWords.Length, actualWords.Length, "word count");
            var differences = new List<string>();
            for (int i = 0; i < expectedWords.Length; i++)
            {
                if (expectedWords[i] != actualWords[i])
                {
                    differences.Add(expectedWords[i] + " <> " + actualWords[i]);
                }
            }
            Assert.AreEqual(0, differences.Count, "expected <> actual:\n" + string.Join("\n", differences.Take(40)));
            Assert.AreEqual(expected, actual);
        }

        private static string ReadFixture(string name)
        {
            return File.ReadAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures", name));
        }

        private static string SameApostrophes(string text)
        {
            return new string(text.Select(c => AlphabetRules.Apostrophes.IndexOf(c) >= 0 ? '\'' : c).ToArray());
        }
    }
}
