using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace NewAlphabetPlugin
{
    /// <summary>
    /// How many old letters and Cyrillic words a preview highlighted or an apply converted.
    /// </summary>
    internal struct ConversionCount
    {
        public int OldLetters;
        public int CyrillicWords;

        public int Total
        {
            get { return OldLetters + CyrillicWords; }
        }
    }

    /// <summary>
    /// Finds old Latin letters and Uzbek Cyrillic words in a Word document, highlights them for review and writes
    /// them in the new alphabet on request. Keeps at most one pending preview per document.
    /// </summary>
    internal sealed class AlphabetConverter
    {
        /// <summary>
        /// The highlight colour a preview uses until the user picks another one.
        /// </summary>
        public const Word.WdColorIndex DefaultPreviewColor = Word.WdColorIndex.wdTurquoise;

        // A highlight made of old letters only and at most this long is a leftover from an earlier preview.
        // Touching old letters such as gʻoʻ or chch make up 4 characters.
        private const int LeftoverMaxLength = 6;

        // What separates one "word" from the next when checking for links: space, tab, paragraph mark,
        // line break, page break, no-break space and end-of-cell mark.
        private const string TokenSeparators = " \t\r\n\v\f\u00A0\u0007";

        // Every character of the Cyrillic block, for growing a hit to the whole word.
        private static readonly string CyrillicCharacters =
            new string(Enumerable.Range(0x0400, 0x100).Select(code => (char)code).ToArray());

        private readonly Dictionary<Word.Document, PendingPreview> previews =
            new Dictionary<Word.Document, PendingPreview>();

        // What the last Apply or Cancel of each document took the preview colour off.
        private readonly Dictionary<Word.Document, RemovedHighlights> lastRemoved =
            new Dictionary<Word.Document, RemovedHighlights>();

        /// <summary>
        /// The highlight colour the next preview uses.
        /// </summary>
        public Word.WdColorIndex PreviewColor { get; set; } = DefaultPreviewColor;

        /// <summary>
        /// Whether a preview looks for old Latin letters: Oʻ, Gʻ, Sh and Ch.
        /// </summary>
        public bool ConvertOldLatin { get; set; } = true;

        /// <summary>
        /// Whether a preview looks for Uzbek Cyrillic words.
        /// </summary>
        public bool ConvertCyrillic { get; set; } = true;

        public bool HasPreview(Word.Document doc)
        {
            return previews.ContainsKey(doc);
        }

        /// <summary>
        /// Highlights every old letter and Cyrillic word in <paramref name="selection"/>, or in the whole document when
        /// it is null. Returns how many were highlighted.
        /// </summary>
        public ConversionCount Preview(Word.Document doc, Word.Range selection)
        {
            using (var batch = new BatchEdit(doc, "Yangi alifbo: körib çiqiş"))
            {
                // Highlighting must not show up as tracked formatting changes.
                batch.PauseTrackChanges();
                RemovePreview(doc);

                // Undo may have brought back highlights an earlier Apply or Cancel took off.
                RemovedHighlights leftovers;
                if (lastRemoved.TryGetValue(doc, out leftovers))
                {
                    lastRemoved.Remove(doc);
                }

                bool hasRevisions = doc.Revisions.Count > 0;
                var found = new List<PendingChange>();
                IEnumerable<Word.Range> scopes = selection != null ? new[] { selection } : GetStories(doc);
                foreach (Word.Range scope in scopes)
                {
                    if (ConvertOldLatin)
                    {
                        foreach (string pattern in AlphabetRules.WordFindPatterns)
                        {
                            FindChanges(scope, pattern, ChangeKind.OldLetter, hasRevisions, leftovers, found);
                        }
                    }
                    if (ConvertCyrillic)
                    {
                        FindChanges(scope, CyrillicRules.WordFindPattern, ChangeKind.CyrillicWord, hasRevisions, leftovers, found);
                    }
                }

                // Highlight only after everything is found, so every original highlight is read before it changes.
                var highlighted = new List<PendingChange>(found.Count);
                foreach (PendingChange change in found)
                {
                    try
                    {
                        change.Range.HighlightColorIndex = PreviewColor;
                        highlighted.Add(change);
                    }
                    catch (COMException)
                    {
                        // Locked text, such as a content control that cannot be edited: leave it out.
                    }
                }

                if (highlighted.Count > 0)
                {
                    previews[doc] = new PendingPreview(PreviewColor, highlighted);
                }

                var count = new ConversionCount();
                foreach (PendingChange change in highlighted)
                {
                    Tally(ref count, change.Kind);
                }
                return count;
            }
        }

        /// <summary>
        /// Converts every previewed letter and word that is still highlighted and removes the preview highlights.
        /// Those whose highlight the user removed are skipped. Returns how many were converted.
        /// </summary>
        public ConversionCount Apply(Word.Document doc)
        {
            var count = new ConversionCount();
            PendingPreview preview;
            if (!previews.TryGetValue(doc, out preview))
            {
                return count;
            }
            previews.Remove(doc);

            using (var batch = new BatchEdit(doc, "Yangi alifbo: qöllaş"))
            {
                // Put the original highlight back first; the new text then inherits it along with the font.
                batch.PauseTrackChanges();
                var removed = new RemovedHighlights(preview.Color);
                var accepted = new List<(int Start, PendingChange Change)>(preview.Changes.Count);
                foreach (PendingChange change in preview.Changes)
                {
                    if (TryRestoreHighlight(change, preview.Color))
                    {
                        removed.Add(change);
                        accepted.Add((change.Range.Start, change));
                    }
                }
                lastRemoved[doc] = removed;

                // Replace from the end of the text backwards. Each replacement changes the length of the text, and when
                // an old letter directly follows one that was already replaced (gʻoʻ, chch, oʻsh), Word would stretch
                // its range over the new letter, so it would no longer be recognised.
                accepted.Sort((a, b) => b.Start.CompareTo(a.Start));

                // The replacement itself follows the user's Track Changes setting.
                batch.ResumeTrackChanges();
                foreach (var item in accepted)
                {
                    try
                    {
                        if (Replace(item.Change))
                        {
                            Tally(ref count, item.Change.Kind);
                        }
                    }
                    catch (COMException)
                    {
                        // Locked text: it stays as it was.
                    }
                }
                return count;
            }
        }

        /// <summary>
        /// Removes the preview highlights without changing any text.
        /// </summary>
        public void Cancel(Word.Document doc)
        {
            if (!previews.ContainsKey(doc))
            {
                return;
            }

            using (var batch = new BatchEdit(doc, "Yangi alifbo: bekor qiliş"))
            {
                batch.PauseTrackChanges();
                RemovePreview(doc);
            }
        }

        /// <summary>
        /// Takes the preview highlights off a document that is closing and forgets the document.
        /// </summary>
        public void Close(Word.Document doc)
        {
            Cancel(doc);
            lastRemoved.Remove(doc);
        }

        /// <summary>
        /// Repaints the document's pending preview in <see cref="PreviewColor"/> after the user picked another colour.
        /// </summary>
        public void Recolor(Word.Document doc)
        {
            PendingPreview preview;
            if (!previews.TryGetValue(doc, out preview) || preview.Color == PreviewColor)
            {
                return;
            }

            using (var batch = new BatchEdit(doc, "Yangi alifbo: rang"))
            {
                batch.PauseTrackChanges();
                foreach (PendingChange change in preview.Changes)
                {
                    try
                    {
                        // Whatever the user un-highlighted to skip it stays skipped.
                        SwapHighlight(change.Range, preview.Color, PreviewColor);
                    }
                    catch (COMException)
                    {
                        // Locked text: it keeps the old colour.
                    }
                }
                preview.Color = PreviewColor;
            }
        }

        // Restores the original highlights of the document's preview and forgets the preview.
        private void RemovePreview(Word.Document doc)
        {
            PendingPreview preview;
            if (previews.TryGetValue(doc, out preview))
            {
                previews.Remove(doc);
                var removed = new RemovedHighlights(preview.Color);
                foreach (PendingChange change in preview.Changes)
                {
                    if (TryRestoreHighlight(change, preview.Color))
                    {
                        removed.Add(change);
                    }
                }
                lastRemoved[doc] = removed;
            }
        }

        // Puts the original highlight back if the preview colour is still there. Returns false when the user
        // removed the highlight to skip this change, or the text is gone. When the user un-highlighted only part of a
        // word, the word is skipped too and the preview colour comes off the rest of it.
        private static bool TryRestoreHighlight(PendingChange change, Word.WdColorIndex previewColor)
        {
            try
            {
                Word.Range range = change.Range;
                return range.Start < range.End && SwapHighlight(range, previewColor, change.OriginalHighlight);
            }
            catch (COMException)
            {
                return false;
            }
        }

        // Gives the characters of the range that are highlighted in one colour another colour. Returns true when all
        // of them were.
        private static bool SwapHighlight(Word.Range range, Word.WdColorIndex from, Word.WdColorIndex to)
        {
            Word.WdColorIndex highlight = range.HighlightColorIndex;
            if (highlight == from)
            {
                range.HighlightColorIndex = to;
                return true;
            }
            if ((int)highlight == (int)Word.WdConstants.wdUndefined)
            {
                foreach (Word.Range character in range.Characters)
                {
                    if (character.HighlightColorIndex == from)
                    {
                        character.HighlightColorIndex = to;
                    }
                }
            }
            return false;
        }

        private static void Tally(ref ConversionCount count, ChangeKind kind)
        {
            if (kind == ChangeKind.CyrillicWord)
            {
                count.CyrillicWords++;
            }
            else
            {
                count.OldLetters++;
            }
        }

        // Runs one wildcard search through the scope and records every hit that should change.
        // Word's Find is used instead of searching Range.Text because text offsets drift at fields and table cells.
        private static void FindChanges(Word.Range scope, string pattern, ChangeKind kind, bool hasRevisions,
            RemovedHighlights leftovers, List<PendingChange> found)
        {
            int scopeEnd = scope.End;
            var links = new LinkDetector();
            Word.Range hit = scope.Duplicate;
            Word.Find find = hit.Find;
            find.ClearFormatting();
            find.Text = pattern;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.Format = false;
            find.MatchCase = true;
            find.MatchWholeWord = false;
            find.MatchWildcards = true;
            find.MatchSoundsLike = false;
            find.MatchAllWordForms = false;

            int lastStart = -1;
            while (find.Execute())
            {
                // After the first hit Word keeps searching to the end of the story, so stop at the end of the scope.
                int start = hit.Start;
                if (hit.End > scopeEnd || start <= lastStart)
                {
                    break;
                }
                lastStart = start;

                if (kind == ChangeKind.CyrillicWord)
                {
                    // A selection can cut a word in half. The whole word is converted, because how each letter is
                    // spelled depends on the letters around it.
                    hit.MoveStartWhile(CyrillicCharacters, Word.WdConstants.wdBackward);
                    hit.MoveEndWhile(CyrillicCharacters, Word.WdConstants.wdForward);
                }

                if (IsConvertible(kind, hit.Text)
                    && !links.IsInLink(hit)
                    && !(hasRevisions && IsTrackedDeletion(hit)))
                {
                    found.Add(new PendingChange(kind, hit.Duplicate, GetHighlight(hit, leftovers)));
                }
                hit.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            }
        }

        private static bool IsConvertible(ChangeKind kind, string text)
        {
            return kind == ChangeKind.CyrillicWord
                ? CyrillicRules.IsUzbekWord(text)
                : AlphabetRules.GetReplacement(text) != null;
        }

        // The highlight to put back later; when the characters are highlighted differently, the first character's.
        // A preview highlight that Undo brought back after the add-in took it off is left over, so it counts as what
        // was there before that preview. So does a highlight that covers old letters and nothing else, as after a
        // restart of Word. The user's own highlights, even in the preview colour, cover other characters too and are kept.
        private static Word.WdColorIndex GetHighlight(Word.Range hit, RemovedHighlights leftovers)
        {
            Word.WdColorIndex highlight = hit.HighlightColorIndex;
            if ((int)highlight == (int)Word.WdConstants.wdUndefined)
            {
                return hit.Characters.First.HighlightColorIndex;
            }
            if (highlight == Word.WdColorIndex.wdNoHighlight)
            {
                return highlight;
            }

            Word.WdColorIndex original;
            if (leftovers != null && leftovers.TryGetOriginal(hit, highlight, out original))
            {
                return original;
            }
            return IsLeftoverPreview(hit, highlight) ? Word.WdColorIndex.wdNoHighlight : highlight;
        }

        // True when the stretch of text around the hit that carries this highlight is made of old letters only.
        private static bool IsLeftoverPreview(Word.Range hit, Word.WdColorIndex highlight)
        {
            Word.Range run = hit.Duplicate;
            GrowWhileHighlighted(run, highlight, backwards: true);
            GrowWhileHighlighted(run, highlight, backwards: false);
            return run.End - run.Start <= LeftoverMaxLength && AlphabetRules.IsOnlyOldLetters(run.Text);
        }

        // Grows the range one character at a time while the next character carries the highlight, stopping once
        // it is longer than any leftover could be.
        private static void GrowWhileHighlighted(Word.Range run, Word.WdColorIndex highlight, bool backwards)
        {
            while (run.End - run.Start <= LeftoverMaxLength)
            {
                Word.Range next = run.Duplicate;
                if (backwards)
                {
                    if (next.MoveStart(Word.WdUnits.wdCharacter, -1) == 0)
                    {
                        return;
                    }
                    next.End = next.Start + 1;
                }
                else
                {
                    if (next.MoveEnd(Word.WdUnits.wdCharacter, 1) == 0)
                    {
                        return;
                    }
                    next.Start = next.End - 1;
                }

                if (next.HighlightColorIndex != highlight)
                {
                    return;
                }
                if (backwards)
                {
                    run.Start = next.Start;
                }
                else
                {
                    run.End = next.End;
                }
            }
        }

        // True for text that Track Changes shows as deleted.
        private static bool IsTrackedDeletion(Word.Range range)
        {
            Word.Revisions revisions = range.Revisions;
            return revisions.Count > 0 && revisions[1].Type == Word.WdRevisionType.wdRevisionDelete;
        }

        // Writes a change in the new alphabet. Returns false when its text was edited since the preview and no
        // longer needs it.
        private static bool Replace(PendingChange change)
        {
            if (change.Kind == ChangeKind.CyrillicWord)
            {
                return ReplaceCyrillicWord(change.Range);
            }

            string newLetter = AlphabetRules.GetReplacement(change.Range.Text);
            if (newLetter == null)
            {
                return false;
            }
            change.Range.Text = newLetter;
            return true;
        }

        private static bool ReplaceCyrillicWord(Word.Range word)
        {
            string text = word.Text;
            if (!CyrillicRules.IsUzbekWord(text))
            {
                return false;
            }

            string[] pieces = CyrillicRules.Transliterate(
                text, CharacterNextTo(word, before: true), CharacterNextTo(word, before: false));
            int start = word.Start;
            if (word.End - start == text.Length && !HasUniformFormatting(word))
            {
                // Letter by letter from the end, so that each new letter keeps the formatting of its old one.
                for (int i = text.Length - 1; i >= 0; i--)
                {
                    Word.Range letter = word.Duplicate;
                    letter.SetRange(start + i, start + i + 1);
                    letter.Text = pieces[i];
                }
            }
            else
            {
                // One replacement, which Track Changes shows as one change. The new word takes the old one's formatting.
                word.Text = string.Concat(pieces);
            }
            return true;
        }

        // True when every letter of the range is formatted alike, so replacing it in one go loses nothing.
        // Word reports a property that varies within the range as undefined, and a font name as empty.
        private static bool HasUniformFormatting(Word.Range range)
        {
            const int Mixed = (int)Word.WdConstants.wdUndefined;
            Word.Font font = range.Font;
            return font.Name.Length > 0
                && font.Size != Mixed
                && font.Bold != Mixed
                && font.Italic != Mixed
                && (int)font.Underline != Mixed
                && (int)font.Color != Mixed
                && font.Superscript != Mixed
                && font.Subscript != Mixed
                && font.StrikeThrough != Mixed
                && font.SmallCaps != Mixed
                && font.AllCaps != Mixed
                && font.Hidden != Mixed;
        }

        // The character just before or after the range in its story, or '\0' at the edge of the story.
        private static char CharacterNextTo(Word.Range range, bool before)
        {
            Word.Range next = range.Duplicate;
            if (before)
            {
                next.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (next.MoveStart(Word.WdUnits.wdCharacter, -1) == 0)
                {
                    return '\0';
                }
            }
            else
            {
                next.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                if (next.MoveEnd(Word.WdUnits.wdCharacter, 1) == 0)
                {
                    return '\0';
                }
            }
            string text = next.Text;
            return string.IsNullOrEmpty(text) ? '\0' : text[0];
        }

        // Every part of the document a reader sees: body, footnotes, endnotes, text boxes, headers and footers.
        private static IEnumerable<Word.Range> GetStories(Word.Document doc)
        {
            // Reading a header first makes Word list headers and footers that were never opened.
            _ = doc.Sections.First.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.StoryType;

            foreach (Word.Range firstOfType in doc.StoryRanges)
            {
                if (!IsReadableStory(firstOfType.StoryType))
                {
                    continue;
                }

                for (Word.Range story = firstOfType; story != null; story = story.NextStoryRange)
                {
                    yield return story;
                    if (IsHeaderOrFooter(story.StoryType))
                    {
                        foreach (Word.Range textBox in GetTextBoxes(story))
                        {
                            yield return textBox;
                        }
                    }
                }
            }
        }

        private static bool IsReadableStory(Word.WdStoryType type)
        {
            switch (type)
            {
                case Word.WdStoryType.wdMainTextStory:
                case Word.WdStoryType.wdFootnotesStory:
                case Word.WdStoryType.wdEndnotesStory:
                case Word.WdStoryType.wdTextFrameStory:
                    return true;
                default:
                    return IsHeaderOrFooter(type);
            }
        }

        private static bool IsHeaderOrFooter(Word.WdStoryType type)
        {
            switch (type)
            {
                case Word.WdStoryType.wdPrimaryHeaderStory:
                case Word.WdStoryType.wdEvenPagesHeaderStory:
                case Word.WdStoryType.wdFirstPageHeaderStory:
                case Word.WdStoryType.wdPrimaryFooterStory:
                case Word.WdStoryType.wdEvenPagesFooterStory:
                case Word.WdStoryType.wdFirstPageFooterStory:
                    return true;
                default:
                    return false;
            }
        }

        // Text boxes anchored in a header or footer are not reached through the text-frame stories.
        private static List<Word.Range> GetTextBoxes(Word.Range headerOrFooter)
        {
            var textBoxes = new List<Word.Range>();
            try
            {
                foreach (Word.Shape shape in headerOrFooter.ShapeRange)
                {
                    try
                    {
                        if (shape.TextFrame.HasText != 0)
                        {
                            textBoxes.Add(shape.TextFrame.TextRange);
                        }
                    }
                    catch (COMException)
                    {
                        // Pictures and other shapes that cannot hold text.
                    }
                }
            }
            catch (COMException)
            {
                // The header or footer has no shapes Word can list.
            }
            return textBoxes;
        }

        // The letters a preview highlighted, and the colour it used for them.
        private sealed class PendingPreview
        {
            public readonly List<PendingChange> Changes;
            public Word.WdColorIndex Color;

            public PendingPreview(Word.WdColorIndex color, List<PendingChange> changes)
            {
                Color = color;
                Changes = changes;
            }
        }

        private enum ChangeKind
        {
            OldLetter,
            CyrillicWord,
        }

        // An old letter or a Cyrillic word waiting to be converted: its live range and the highlight it had before the preview.
        private sealed class PendingChange
        {
            public readonly ChangeKind Kind;
            public readonly Word.Range Range;
            public readonly Word.WdColorIndex OriginalHighlight;

            public PendingChange(ChangeKind kind, Word.Range range, Word.WdColorIndex originalHighlight)
            {
                Kind = kind;
                Range = range;
                OriginalHighlight = originalHighlight;
            }
        }

        // The preview highlights an Apply or Cancel took off, and where. Undo can bring them back; the next preview
        // must then treat them as the add-in's leftovers rather than as the user's own highlights.
        private sealed class RemovedHighlights
        {
            private readonly Word.WdColorIndex color;
            private readonly Dictionary<(int Start, int End), List<PendingChange>> byPosition =
                new Dictionary<(int Start, int End), List<PendingChange>>();

            public RemovedHighlights(Word.WdColorIndex color)
            {
                this.color = color;
            }

            // Call before the change's text is replaced, while it is still where Undo puts it back.
            public void Add(PendingChange change)
            {
                var position = (change.Range.Start, change.Range.End);
                List<PendingChange> changes;
                if (!byPosition.TryGetValue(position, out changes))
                {
                    changes = new List<PendingChange>();
                    byPosition.Add(position, changes);
                }
                changes.Add(change);
            }

            // True when the hit has the removed highlight's colour and place; gives the highlight from before that preview.
            public bool TryGetOriginal(Word.Range hit, Word.WdColorIndex highlight, out Word.WdColorIndex original)
            {
                List<PendingChange> changes;
                if (highlight == color && byPosition.TryGetValue((hit.Start, hit.End), out changes))
                {
                    foreach (PendingChange change in changes)
                    {
                        // Positions count from the start of each story, so the story must match too.
                        if (change.Range.InStory(hit))
                        {
                            original = change.OriginalHighlight;
                            return true;
                        }
                    }
                }
                original = Word.WdColorIndex.wdNoHighlight;
                return false;
            }
        }

        // Tells whether a hit is part of a web address or e-mail. Reads each paragraph's text once and looks
        // closer only at hits in paragraphs that contain something link-like. Hits must come in document order.
        private sealed class LinkDetector
        {
            private int paragraphEnd = -1;
            private bool paragraphHasLink;

            public bool IsInLink(Word.Range hit)
            {
                if (hit.Start >= paragraphEnd)
                {
                    Word.Range paragraph = hit.Paragraphs.First.Range;
                    paragraphEnd = paragraph.End;
                    paragraphHasLink = AlphabetRules.LooksLikeLink(paragraph.Text);
                }
                if (!paragraphHasLink)
                {
                    return false;
                }

                Word.Range token = hit.Duplicate;
                token.MoveStartUntil(TokenSeparators, Word.WdConstants.wdBackward);
                token.MoveEndUntil(TokenSeparators, Word.WdConstants.wdForward);
                return AlphabetRules.IsInLink(token.Text, hit.Start - token.Start);
            }
        }

        // Makes a batch of edits a single undo step, hides screen redraws while it runs and can pause
        // Track Changes. Everything is restored on Dispose, even after an error.
        private sealed class BatchEdit : IDisposable
        {
            private readonly Word.Application app;
            private readonly Word.Document doc;
            private readonly bool screenUpdating;
            private readonly bool trackRevisions;

            public BatchEdit(Word.Document doc, string undoName)
            {
                this.doc = doc;
                app = doc.Application;
                screenUpdating = app.ScreenUpdating;
                trackRevisions = doc.TrackRevisions;
                app.UndoRecord.StartCustomRecord(undoName);
                app.ScreenUpdating = false;
                app.System.Cursor = Word.WdCursorType.wdCursorWait;
            }

            public void PauseTrackChanges()
            {
                if (doc.TrackRevisions)
                {
                    doc.TrackRevisions = false;
                }
            }

            public void ResumeTrackChanges()
            {
                if (doc.TrackRevisions != trackRevisions)
                {
                    doc.TrackRevisions = trackRevisions;
                }
            }

            public void Dispose()
            {
                try
                {
                    ResumeTrackChanges();
                }
                finally
                {
                    // Always give the screen back, or Word looks frozen.
                    app.UndoRecord.EndCustomRecord();
                    app.System.Cursor = Word.WdCursorType.wdCursorNormal;
                    app.ScreenUpdating = screenUpdating;
                    app.ScreenRefresh();
                }
            }
        }
    }
}
