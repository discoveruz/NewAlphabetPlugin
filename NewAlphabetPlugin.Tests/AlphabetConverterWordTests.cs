using System;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace NewAlphabetPlugin.Tests
{
    /// <summary>
    /// Runs <see cref="AlphabetConverter"/> against a real, hidden Word; each test gets a new blank document.
    /// On a machine without Word every test is reported as skipped (inconclusive), not failed.
    /// </summary>
    [TestClass]
    public class AlphabetConverterWordTests
    {
        private const Word.WdHeaderFooterIndex Primary = Word.WdHeaderFooterIndex.wdHeaderFooterPrimary;

        private static Word.Application word;
        private static string wordMissingReason;

        private Word.Document doc;
        private AlphabetConverter converter;

        [ClassInitialize]
        public static void StartWord(TestContext context)
        {
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                wordMissingReason = "Microsoft Word is not installed, so the Word tests cannot run here.";
                return;
            }

            try
            {
                word = (Word.Application)Activator.CreateInstance(wordType);
                word.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            }
            catch (Exception ex)
            {
                wordMissingReason = "Microsoft Word could not be started: " + ex.Message;
            }
        }

        [ClassCleanup]
        public static void QuitWord()
        {
            if (word != null)
            {
                ((Word._Application)word).Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                word = null;
            }
        }

        [TestInitialize]
        public void CreateDocument()
        {
            if (word == null)
            {
                Assert.Inconclusive(wordMissingReason);
            }
            doc = word.Documents.Add();
            converter = new AlphabetConverter();
        }

        [TestCleanup]
        public void CloseDocument()
        {
            if (doc != null)
            {
                ((Word._Document)doc).Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                doc = null;
            }
        }

        // ---- Preview ----

        [TestMethod]
        public void Preview_HighlightsOnlyTheOldLetters_AndKeepsTheText()
        {
            SetBody("O\u02BBzbekiston tog\u02BB Toshkent chiroyli");

            int count = converter.Preview(doc, null).OldLetters;

            Assert.AreEqual(4, count);
            Assert.AreEqual("O\u02BB|g\u02BB|sh|ch", Highlighted(doc.Content));
            Assert.AreEqual("O\u02BBzbekiston tog\u02BB Toshkent chiroyli", BodyText);
            Assert.IsTrue(converter.HasPreview(doc));
        }

        [TestMethod]
        public void Preview_FindsNothing_InNewAlphabetText()
        {
            SetBody(Samples.NewText);

            Assert.AreEqual(0, converter.Preview(doc, null).Total);
            Assert.IsFalse(converter.HasPreview(doc));
        }

        [TestMethod]
        public void Preview_WithASelection_OnlyTouchesTheSelection()
        {
            SetBody("shahar\rshahar");

            int count = converter.Preview(doc, doc.Paragraphs[1].Range).OldLetters;
            converter.Apply(doc);

            Assert.AreEqual(1, count);
            Assert.AreEqual("şahar\rshahar", BodyText);
        }

        [TestMethod]
        public void Preview_IgnoresALetterThatTheSelectionCutsInHalf()
        {
            SetBody("Toshkent");

            int count = converter.Preview(doc, doc.Range(0, 3)).OldLetters; // "Tos"

            Assert.AreEqual(0, count);
        }

        [TestMethod]
        public void Preview_Twice_StillRestoresTheOriginalHighlight()
        {
            SetBody("Toshkent");
            converter.Preview(doc, null);
            converter.Preview(doc, null);

            converter.Cancel(doc);

            Assert.AreEqual("", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Preview_SkipsTextThatIsATrackedDeletion()
        {
            SetBody("shahar Toshkent");
            doc.TrackRevisions = true;
            RangeOf("shahar ").Delete();
            doc.TrackRevisions = false;

            int count = converter.Preview(doc, null).OldLetters;

            Assert.AreEqual(1, count, "only the sh in Toshkent; the deleted word does not count");
        }

        // ---- Apply ----

        [TestMethod]
        public void Apply_ConvertsEveryApostropheAndCase()
        {
            SetBody(Samples.OldText);

            int highlighted = converter.Preview(doc, null).OldLetters;
            int replaced = converter.Apply(doc).OldLetters;

            Assert.AreEqual(Samples.OldLetterCount, highlighted);
            Assert.AreEqual(Samples.OldLetterCount, replaced);
            Assert.AreEqual(Samples.NewText, BodyText);
            Assert.AreEqual("", Highlighted(doc.Content), "the preview highlights must be gone");
            Assert.IsFalse(converter.HasPreview(doc));
        }

        [TestMethod]
        [DataRow("g\u02BBo\u02BBdaygan", "ğödaygan", 2, DisplayName = "gʻoʻ: same kind of letter twice in a row")]
        [DataRow("boyvachcha", "boyvaçça", 2, DisplayName = "chch")]
        [DataRow("qo\u02BBshiq", "qöşiq", 2, DisplayName = "oʻ then sh")]
        [DataRow("Cho\u02BBl", "Çöl", 2, DisplayName = "Ch then oʻ")]
        [DataRow("«Birorta g\u02BBo\u02BBdaygan boyvachcha paydo bo\u02BBlsa, kechirmayman!» – prezident",
                 "«Birorta ğödaygan boyvaçça paydo bölsa, keçirmayman!» – prezident", 6, DisplayName = "the reported sentence")]
        public void Apply_ConvertsOldLettersThatTouchEachOther(string oldText, string expected, int letters)
        {
            SetBody(oldText);

            Assert.AreEqual(letters, converter.Preview(doc, null).OldLetters, "letters highlighted");
            Assert.AreEqual(letters, converter.Apply(doc).OldLetters, "letters replaced");

            Assert.AreEqual(expected, BodyText);
            Assert.AreEqual("", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Apply_SkipsALetterWhoseHighlightWasRemoved()
        {
            SetBody("English Toshkent");
            converter.Preview(doc, null);
            RangeOf("sh").HighlightColorIndex = Word.WdColorIndex.wdNoHighlight; // the sh in "English"

            int replaced = converter.Apply(doc).OldLetters;

            Assert.AreEqual(1, replaced);
            Assert.AreEqual("English Toşkent", BodyText);
        }

        [TestMethod]
        public void Apply_KeepsExistingHighlightAndBold()
        {
            SetBody("Toshkent shahri");
            Word.Range city = RangeOf("Toshkent");
            city.HighlightColorIndex = Word.WdColorIndex.wdYellow;
            city.Bold = 1;
            converter.Preview(doc, null);

            converter.Apply(doc);

            Assert.AreEqual("Toşkent şahri", BodyText);
            Assert.AreEqual("Toşkent", Highlighted(doc.Content, Word.WdColorIndex.wdYellow));
            Assert.AreNotEqual(0, RangeOf("ş").Bold, "the new letter keeps the bold of the old one");
            Assert.AreEqual("", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Apply_LeavesLinksAndEmailsAlone()
        {
            SetBody("Toshkent https://shop.uz info@school.uz www.chess.com shosh.uz");

            int count = converter.Preview(doc, null).OldLetters;
            converter.Apply(doc);

            Assert.AreEqual(1, count);
            Assert.AreEqual("Toşkent https://shop.uz info@school.uz www.chess.com shosh.uz", BodyText);
        }

        [TestMethod]
        public void Apply_ConvertsTextAfterAHyperlink_AndKeepsTheLinkAddress()
        {
            SetBody("x Toshkent");
            doc.Hyperlinks.Add(RangeOf("x"), "https://example.com/shop", TextToDisplay: "bu yerda");

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("bu yerda Toşkent", BodyText);
            Assert.AreEqual("https://example.com/shop", doc.Hyperlinks[1].Address);
        }

        [TestMethod]
        public void Apply_LeavesAHyperlinkThatShowsAnAddress()
        {
            SetBody("x Toshkent");
            doc.Hyperlinks.Add(RangeOf("x"), "https://shop.uz", TextToDisplay: "https://shop.uz");

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("https://shop.uz Toşkent", BodyText);
        }

        [TestMethod]
        public void Apply_ConvertsHeaderAndFooter()
        {
            SetBody("matn");
            Word.Section section = doc.Sections[1];
            section.Headers[Primary].Range.Text = "shahar";
            section.Footers[Primary].Range.Text = "Cho\u02BBl";

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("şahar", TextOf(section.Headers[Primary].Range));
            Assert.AreEqual("Çöl", TextOf(section.Footers[Primary].Range));
        }

        [TestMethod]
        public void Apply_ConvertsFootnote()
        {
            SetBody("Toshkent");
            Word.Footnote footnote = doc.Footnotes.Add(doc.Range(8, 8), Text: "Izoh: bog\u02BB");

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("Izoh: boğ", TextOf(footnote.Range));
            Assert.IsTrue(BodyText.StartsWith("Toşkent", StringComparison.Ordinal), BodyText);
        }

        [TestMethod]
        public void Apply_ConvertsTableCells()
        {
            Word.Table table = doc.Tables.Add(doc.Range(0, 0), 1, 2);
            table.Cell(1, 1).Range.Text = "g\u02BBisht";
            table.Cell(1, 2).Range.Text = "Choy";

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("ğişt", TextOf(table.Cell(1, 1).Range));
            Assert.AreEqual("Çoy", TextOf(table.Cell(1, 2).Range));
        }

        [TestMethod]
        public void Apply_ConvertsTextBox()
        {
            Word.Shape box = doc.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 50, 200, 40);
            box.TextFrame.TextRange.Text = "O\u02BBrik";

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("Örik", TextOf(box.TextFrame.TextRange));
        }

        [TestMethod]
        public void Apply_ConvertsTextBoxInHeader()
        {
            Word.HeaderFooter header = doc.Sections[1].Headers[Primary];
            Word.Shape box = header.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 20, 200, 30, header.Range);
            box.TextFrame.TextRange.Text = "bosh";

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("boş", TextOf(box.TextFrame.TextRange));
            Assert.AreEqual("", Highlighted(box.TextFrame.TextRange));
        }

        [TestMethod]
        public void Apply_SkipsLettersDeletedAfterThePreview()
        {
            SetBody("Toshkent chiroyli");
            converter.Preview(doc, null);
            RangeOf(" chiroyli").Delete();

            int replaced = converter.Apply(doc).OldLetters;

            Assert.AreEqual(1, replaced);
            Assert.AreEqual("Toşkent", BodyText);
        }

        [TestMethod]
        public void Apply_DoesNotReplaceLettersRetypedAfterThePreview()
        {
            SetBody("Toshkent");
            converter.Preview(doc, null);
            RangeOf("sh").Text = "ss";

            converter.Apply(doc);

            Assert.AreEqual("Tosskent", BodyText);
        }

        [TestMethod]
        public void ApplyAndCancel_WithoutAPreview_DoNothing()
        {
            SetBody("Toshkent");

            Assert.AreEqual(0, converter.Apply(doc).Total);
            converter.Cancel(doc);

            Assert.AreEqual("Toshkent", BodyText);
        }

        // ---- Track Changes and undo ----

        [TestMethod]
        public void TrackChanges_PreviewIsNotTracked_ButTheReplacementIs()
        {
            SetBody("Toshkent");
            doc.TrackRevisions = true;

            converter.Preview(doc, null);
            Assert.AreEqual(0, doc.Revisions.Count, "the preview highlight must not become a tracked change");

            converter.Apply(doc);
            Assert.IsTrue(doc.TrackRevisions, "Track Changes must be back on");
            Assert.IsTrue(doc.Revisions.Count > 0, "the replacement must be tracked");

            doc.Revisions.AcceptAll();
            Assert.AreEqual("Toşkent", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Undo_RevertsApplyInOneStep()
        {
            SetBody("Toshkent chiroyli");
            converter.Preview(doc, null);
            converter.Apply(doc);

            doc.Undo(1);

            Assert.AreEqual("Toshkent chiroyli", BodyText);
            Assert.AreEqual("sh|ch", Highlighted(doc.Content), "one undo goes back to the preview");
        }

        [TestMethod]
        public void Cancel_AfterAnUndoneApply_RemovesTheLeftoverHighlights()
        {
            SetBody("Toshkent chiroyli");
            converter.Preview(doc, null);
            converter.Apply(doc);
            doc.Undo(1); // the old letters return with their preview highlight, but that preview is already forgotten

            converter.Preview(doc, null);
            converter.Cancel(doc);

            Assert.AreEqual("Toshkent chiroyli", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Apply_AfterAnUndoneApply_LeavesNoHighlights()
        {
            SetBody("Toshkent chiroyli");
            converter.Preview(doc, null);
            converter.Apply(doc);
            doc.Undo(1);

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("Toşkent çiroyli", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Cancel_AfterAnUndoneApply_RemovesLeftoversOnTouchingLetters()
        {
            SetBody("g\u02BBo\u02BBdaygan boyvachcha");
            converter.Preview(doc, null);
            converter.Apply(doc);
            doc.Undo(1);

            converter.Preview(doc, null);
            converter.Cancel(doc);

            Assert.AreEqual("", Highlighted(doc.Content));
        }

        // ---- Chosen colour ----

        [TestMethod]
        public void Preview_UsesTheChosenColour()
        {
            SetBody("Toshkent chiroyli");
            converter.PreviewColor = Word.WdColorIndex.wdYellow;

            converter.Preview(doc, null);

            Assert.AreEqual("sh|ch", Highlighted(doc.Content, Word.WdColorIndex.wdYellow));
            Assert.AreEqual("", Highlighted(doc.Content), "no turquoise");
        }

        [TestMethod]
        public void Apply_RemovesTheChosenColour()
        {
            SetBody("Toshkent chiroyli");
            converter.PreviewColor = Word.WdColorIndex.wdYellow;
            converter.Preview(doc, null);

            converter.Apply(doc);

            Assert.AreEqual("Toşkent çiroyli", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content, Word.WdColorIndex.wdYellow));
        }

        [TestMethod]
        public void Cancel_RemovesTheChosenColour()
        {
            SetBody("Toshkent chiroyli");
            converter.PreviewColor = Word.WdColorIndex.wdYellow;
            converter.Preview(doc, null);

            converter.Cancel(doc);

            Assert.AreEqual("Toshkent chiroyli", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content, Word.WdColorIndex.wdYellow));
        }

        [TestMethod]
        public void Recolor_RepaintsAPendingPreview_AndKeepsSkippedLettersSkipped()
        {
            SetBody("English Toshkent chiroyli");
            converter.Preview(doc, null);
            RangeOf("sh").HighlightColorIndex = Word.WdColorIndex.wdNoHighlight; // skip the sh in "English"

            converter.PreviewColor = Word.WdColorIndex.wdBrightGreen;
            converter.Recolor(doc);

            Assert.AreEqual("sh|ch", Highlighted(doc.Content, Word.WdColorIndex.wdBrightGreen));
            Assert.AreEqual("", Highlighted(doc.Content), "no turquoise left");

            converter.Apply(doc);

            Assert.AreEqual("English Toşkent çiroyli", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content, Word.WdColorIndex.wdBrightGreen));
        }

        [TestMethod]
        public void Apply_KeepsTheUsersOwnHighlightInThePreviewColour()
        {
            SetBody("Toshkent shahri");
            RangeOf("Toshkent").HighlightColorIndex = Word.WdColorIndex.wdYellow;
            converter.PreviewColor = Word.WdColorIndex.wdYellow;
            converter.Preview(doc, null);

            converter.Apply(doc);

            Assert.AreEqual("Toşkent şahri", BodyText);
            Assert.AreEqual("Toşkent", Highlighted(doc.Content, Word.WdColorIndex.wdYellow));
        }

        [TestMethod]
        public void Cancel_AfterAnUndoneApply_RemovesLeftoversOfAnEarlierColour()
        {
            SetBody("Toshkent chiroyli");
            converter.Preview(doc, null);
            converter.Apply(doc);
            doc.Undo(1); // turquoise leftovers on the old letters

            converter.PreviewColor = Word.WdColorIndex.wdYellow;
            converter.Preview(doc, null);
            converter.Cancel(doc);

            Assert.AreEqual("", Highlighted(doc.Content), "no turquoise");
            Assert.AreEqual("", Highlighted(doc.Content, Word.WdColorIndex.wdYellow), "no yellow");
        }

        // ---- Cancel ----

        [TestMethod]
        public void Cancel_RemovesTheHighlights_AndKeepsTheText()
        {
            SetBody("Toshkent chiroyli");
            converter.Preview(doc, null);

            converter.Cancel(doc);

            Assert.AreEqual("Toshkent chiroyli", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content));
            Assert.IsFalse(converter.HasPreview(doc));
        }

        [TestMethod]
        public void Cancel_KeepsAnExistingHighlight()
        {
            SetBody("Toshkent shahri");
            RangeOf("Toshkent").HighlightColorIndex = Word.WdColorIndex.wdYellow;
            converter.Preview(doc, null);

            converter.Cancel(doc);

            Assert.AreEqual("Toshkent", Highlighted(doc.Content, Word.WdColorIndex.wdYellow));
            Assert.AreEqual("", Highlighted(doc.Content));
        }

        // ---- Cyrillic ----

        [TestMethod]
        public void Preview_HighlightsWholeCyrillicWords_AndCountsThemApart()
        {
            SetBody("Тошкент шаҳри, Toshkent");

            ConversionCount count = converter.Preview(doc, null);

            Assert.AreEqual(2, count.CyrillicWords);
            Assert.AreEqual(1, count.OldLetters);
            Assert.AreEqual("Тошкент|шаҳри|sh", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Apply_ConvertsCyrillicText()
        {
            SetBody("«Бирорта ғўдайган бойвачча пайдо бўлса, кечирмайман!» – президент");

            converter.Preview(doc, null);
            ConversionCount count = converter.Apply(doc);

            Assert.AreEqual(7, count.CyrillicWords);
            Assert.AreEqual("«Birorta ğödaygan boyvaçça paydo bölsa, keçirmayman!» – prezident", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content));
            Assert.IsFalse(converter.HasPreview(doc));
        }

        [TestMethod]
        [DataRow(true, true, "Özbekiston va Özbekiston", DisplayName = "both scripts")]
        [DataRow(true, false, "Özbekiston ва Ўзбекистон", DisplayName = "old Latin only")]
        [DataRow(false, true, "O'zbekiston va Özbekiston", DisplayName = "Cyrillic only")]
        public void Preview_LooksOnlyForTheTickedScripts(bool oldLatin, bool cyrillic, string expected)
        {
            SetBody("O'zbekiston ва Ўзбекистон");
            converter.ConvertOldLatin = oldLatin;
            converter.ConvertCyrillic = cyrillic;

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual(expected, BodyText);
        }

        [TestMethod]
        public void Apply_UsesTheLettersAroundACyrillicWord()
        {
            SetBody("aцi Цирк");

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("atsi Sirk", BodyText, "Ц after the Latin a is ts; at the start of a word it is s");
        }

        [TestMethod]
        public void Apply_KeepsTheFormattingOfEachLetter()
        {
            SetBody("медальон");
            RangeOf("д").Bold = 1;
            RangeOf("м").Font.Color = Word.WdColor.wdColorRed;

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("medalyon", BodyText, "ь is dropped and о after it becomes yo");
            Assert.AreNotEqual(0, RangeOf("d").Bold, "d keeps the bold of д");
            Assert.AreEqual(0, RangeOf("me").Bold);
            Assert.AreEqual(0, RangeOf("alyon").Bold);
            Assert.AreEqual(Word.WdColor.wdColorRed, RangeOf("m").Font.Color, "m keeps the colour of м");
            Assert.AreEqual(Word.WdColor.wdColorAutomatic, RangeOf("edalyon").Font.Color);
        }

        [TestMethod]
        public void TrackChanges_ShowEachCyrillicWordAsOneReplacement()
        {
            SetBody("Тошкент шаҳри");
            doc.TrackRevisions = true;

            converter.Preview(doc, null);
            Assert.AreEqual(0, doc.Revisions.Count, "the preview highlight must not become a tracked change");
            converter.Apply(doc);

            Assert.AreEqual(4, doc.Revisions.Count, "one deletion and one insertion per word, not per letter");
            doc.Revisions.AcceptAll();
            Assert.AreEqual("Toşkent şahri", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Preview_WithASelection_ConvertsTheWholeWordItCutsInHalf()
        {
            SetBody("Тошкент шаҳри");

            ConversionCount count = converter.Preview(doc, doc.Range(0, 3)); // "Тош"
            converter.Apply(doc);

            Assert.AreEqual(1, count.CyrillicWords);
            Assert.AreEqual("Toşkent шаҳри", BodyText);
        }

        [TestMethod]
        public void Preview_LeavesNonUzbekWordsAndLinksAlone()
        {
            SetBody("Тошкент щука https://сайт.уз мактаб.уз");

            ConversionCount count = converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual(1, count.CyrillicWords);
            Assert.AreEqual("Toşkent щука https://сайт.уз мактаб.уз", BodyText);
        }

        [TestMethod]
        public void Apply_ConvertsAWordEndingJoinedToAnAddress()
        {
            SetBody("Kun.uz'га мурожаат қилган, https://сайт.уз");

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("Kun.uz'ga murojaat qilgan, https://сайт.уз", BodyText);
        }

        [TestMethod]
        public void Apply_SkipsCyrillicWordsWhoseHighlightWasRemoved_EvenInPart()
        {
            SetBody("Тошкент Москва Самарқанд");
            converter.Preview(doc, null);
            RangeOf("Москва").HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            RangeOf("Сам").HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;

            ConversionCount count = converter.Apply(doc);

            Assert.AreEqual(1, count.CyrillicWords);
            Assert.AreEqual("Toşkent Москва Самарқанд", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content), "no preview colour is left on the rest of Самарқанд");
        }

        [TestMethod]
        public void Apply_ConvertsCyrillicInAHeaderAndAFootnote()
        {
            SetBody("матн");
            Word.Section section = doc.Sections[1];
            section.Headers[Primary].Range.Text = "Тошкент";
            Word.Footnote footnote = doc.Footnotes.Add(doc.Range(4, 4), Text: "Изоҳ: боғ");

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("Toşkent", TextOf(section.Headers[Primary].Range));
            Assert.AreEqual("Izoh: boğ", TextOf(footnote.Range));
            Assert.IsTrue(BodyText.StartsWith("matn", StringComparison.Ordinal), BodyText);
        }

        [TestMethod]
        public void Cancel_AfterAnUndoneApply_RemovesLeftoversOnCyrillicWords()
        {
            SetBody("Тошкент шаҳри");
            converter.Preview(doc, null);
            converter.Apply(doc);
            doc.Undo(1); // the Cyrillic words return with their preview highlight, but that preview is already forgotten
            Assert.AreEqual("Тошкент|шаҳри", Highlighted(doc.Content), "one undo goes back to the preview");

            converter.Preview(doc, null);
            converter.Cancel(doc);

            Assert.AreEqual("Тошкент шаҳри", BodyText);
            Assert.AreEqual("", Highlighted(doc.Content));
        }

        [TestMethod]
        public void Apply_AfterAnUndoneApply_KeepsTheUsersOwnWordHighlightInThePreviewColour()
        {
            SetBody("Тошкент шаҳри");
            RangeOf("шаҳри").HighlightColorIndex = AlphabetConverter.DefaultPreviewColor; // the user's own
            converter.Preview(doc, null);
            converter.Apply(doc);
            doc.Undo(1);

            converter.Preview(doc, null);
            converter.Apply(doc);

            Assert.AreEqual("Toşkent şahri", BodyText);
            Assert.AreEqual("şahri", Highlighted(doc.Content));
        }

        // ---- helpers ----

        private string BodyText
        {
            get { return TextOf(doc.Content); }
        }

        private void SetBody(string text)
        {
            doc.Content.Text = text;
        }

        // The body range of the first occurrence of a fragment. Only used on plain text, where positions equal string indexes.
        private Word.Range RangeOf(string fragment)
        {
            int start = BodyText.IndexOf(fragment, StringComparison.Ordinal);
            Assert.IsTrue(start >= 0, "'" + fragment + "' is not in the document");
            return doc.Range(start, start + fragment.Length);
        }

        private static string TextOf(Word.Range range)
        {
            return range.Text.TrimEnd('\r', '\a');
        }

        // The characters carrying a highlight colour (the preview colour by default), with "|" between separate runs.
        private static string Highlighted(Word.Range range, Word.WdColorIndex color = AlphabetConverter.DefaultPreviewColor)
        {
            var text = new StringBuilder();
            bool inRun = false;
            foreach (Word.Range character in range.Characters)
            {
                bool match = character.HighlightColorIndex == color;
                if (match && !inRun && text.Length > 0)
                {
                    text.Append('|');
                }
                if (match)
                {
                    text.Append(character.Text);
                }
                inRun = match;
            }
            return text.ToString();
        }
    }
}
