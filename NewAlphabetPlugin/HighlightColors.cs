using System.Drawing;
using Word = Microsoft.Office.Interop.Word;

namespace NewAlphabetPlugin
{
    /// <summary>
    /// The highlight colours the preview can use.
    /// </summary>
    internal static class HighlightColors
    {
        /// <summary>
        /// Word's highlight colours in the order of Word's own picker, with Uzbek names in the new alphabet.
        /// </summary>
        public static readonly HighlightColor[] All =
        {
            new HighlightColor(Word.WdColorIndex.wdYellow, "Sariq", Color.FromArgb(255, 255, 0)),
            new HighlightColor(Word.WdColorIndex.wdBrightGreen, "Oç yaşil", Color.FromArgb(0, 255, 0)),
            new HighlightColor(Word.WdColorIndex.wdTurquoise, "Feruza", Color.FromArgb(0, 255, 255)),
            new HighlightColor(Word.WdColorIndex.wdPink, "Puşti", Color.FromArgb(255, 0, 255)),
            new HighlightColor(Word.WdColorIndex.wdBlue, "Kök", Color.FromArgb(0, 0, 255)),
            new HighlightColor(Word.WdColorIndex.wdRed, "Qizil", Color.FromArgb(255, 0, 0)),
            new HighlightColor(Word.WdColorIndex.wdDarkBlue, "Töq kök", Color.FromArgb(0, 0, 128)),
            new HighlightColor(Word.WdColorIndex.wdTeal, "Töq feruza", Color.FromArgb(0, 128, 128)),
            new HighlightColor(Word.WdColorIndex.wdGreen, "Yaşil", Color.FromArgb(0, 128, 0)),
            new HighlightColor(Word.WdColorIndex.wdViolet, "Binafşa", Color.FromArgb(128, 0, 128)),
            new HighlightColor(Word.WdColorIndex.wdDarkRed, "Töq qizil", Color.FromArgb(128, 0, 0)),
            new HighlightColor(Word.WdColorIndex.wdDarkYellow, "Töq sariq", Color.FromArgb(128, 128, 0)),
            new HighlightColor(Word.WdColorIndex.wdGray50, "Töq kulrang", Color.FromArgb(128, 128, 128)),
            new HighlightColor(Word.WdColorIndex.wdGray25, "Oç kulrang", Color.FromArgb(192, 192, 192)),
            new HighlightColor(Word.WdColorIndex.wdBlack, "Qora", Color.FromArgb(0, 0, 0)),
        };
    }

    internal sealed class HighlightColor
    {
        public readonly Word.WdColorIndex Index;
        public readonly string Name;
        public readonly Color Swatch;

        public HighlightColor(Word.WdColorIndex index, string name, Color swatch)
        {
            Index = index;
            Name = name;
            Swatch = swatch;
        }
    }
}
