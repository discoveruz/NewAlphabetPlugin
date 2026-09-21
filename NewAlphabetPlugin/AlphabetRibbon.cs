using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace NewAlphabetPlugin
{
    public partial class AlphabetRibbon
    {
        private const string MessageTitle = "Yangi alifbo";

        // {0} is the version.
        private const string AboutText =
            "Muallif: Diyorbek Satimboyev\n" +
            "Versiya: {0}\n\n" +
            "IT yeçimlar va biznesni avtomatlaştiriş uçun telegram orqali @devdiyorbek ga boğlaning";

        private static AlphabetConverter Converter
        {
            get { return Globals.ThisAddIn.Converter; }
        }

        private void AlphabetRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            chkOldLatin.Checked = UserSettings.ConvertOldLatin;
            chkCyrillic.Checked = UserSettings.ConvertCyrillic;
            btnAbout.Image = DrawInfoIcon();

            // The colour list is filled here because its swatches are drawn in code.
            Word.WdColorIndex saved = UserSettings.PreviewColor;
            foreach (HighlightColor color in HighlightColors.All)
            {
                RibbonDropDownItem item = Factory.CreateRibbonDropDownItem();
                item.Label = color.Name;
                item.Image = DrawSwatch(color.Swatch);
                item.Tag = color.Index;
                ddColor.Items.Add(item);
                if (color.Index == saved)
                {
                    ddColor.SelectedItem = item;
                }
            }
        }

        private void ddColor_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var color = (Word.WdColorIndex)ddColor.SelectedItem.Tag;
            UserSettings.PreviewColor = color;
            Converter.PreviewColor = color;

            // Highlights already on screen change colour straight away.
            Word.Document doc = GetActiveDocument();
            if (doc != null && Converter.HasPreview(doc))
            {
                RunOnActiveDocument(Converter.Recolor);
            }
        }

        // The ticks apply to the next preview.
        private void chkScript_Click(object sender, RibbonControlEventArgs e)
        {
            Converter.ConvertOldLatin = chkOldLatin.Checked;
            Converter.ConvertCyrillic = chkCyrillic.Checked;
            UserSettings.ConvertOldLatin = chkOldLatin.Checked;
            UserSettings.ConvertCyrillic = chkCyrillic.Checked;
        }

        private void btnPreview_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Converter.ConvertOldLatin && !Converter.ConvertCyrillic)
            {
                ShowMessage("Avval «Eski lotin» yoki «Kirill»ni tanlang.", MessageBoxIcon.Information);
                return;
            }

            RunOnActiveDocument(doc =>
            {
                ConversionCount count = Converter.Preview(doc, GetSelectedText(doc));
                if (count.Total == 0)
                {
                    ShowMessage(NothingFoundMessage(), MessageBoxIcon.Information);
                }
                else
                {
                    doc.Application.StatusBar = Describe(count) + " belgilandi. Tekşirib, «Qöllaş» tugmasini bosing.";
                }
            });
        }

        private void btnApply_Click(object sender, RibbonControlEventArgs e)
        {
            RunOnActiveDocument(doc =>
            {
                ConversionCount count = Converter.Apply(doc);
                doc.Application.StatusBar = count.Total == 0
                    ? "Heç narsa özgartirilmadi."
                    : Describe(count) + " yangi alifboga ötkazildi.";
            });
        }

        // "3 ta harf", "12 ta kirillça söz" or "3 ta harf va 12 ta kirillça söz".
        private static string Describe(ConversionCount count)
        {
            string letters = string.Format("{0} ta harf", count.OldLetters);
            string words = string.Format("{0} ta kirillça söz", count.CyrillicWords);
            if (count.CyrillicWords == 0)
            {
                return letters;
            }
            return count.OldLetters == 0 ? words : letters + " va " + words;
        }

        private static string NothingFoundMessage()
        {
            if (!Converter.ConvertCyrillic)
            {
                return "Eski alifbodagi harflar topilmadi.";
            }
            return Converter.ConvertOldLatin
                ? "Eski alifbodagi harflar ham, kirillça sözlar ham topilmadi."
                : "Kirillça sözlar topilmadi.";
        }

        private void btnCancel_Click(object sender, RibbonControlEventArgs e)
        {
            RunOnActiveDocument(doc =>
            {
                Converter.Cancel(doc);
                doc.Application.StatusBar = "Belgilar olib taşlandi.";
            });
        }

        // The version is the AssemblyVersion in Properties\AssemblyInfo.cs.
        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            Version version = typeof(AlphabetRibbon).Assembly.GetName().Version;
            ShowMessage(string.Format(AboutText, version.ToString(3)), MessageBoxIcon.Information);
        }

        /// <summary>
        /// Enables Apply and Cancel only while the active document has a preview waiting.
        /// </summary>
        internal void RefreshButtons()
        {
            Word.Document doc = GetActiveDocument();
            bool hasPreview = doc != null && Converter.HasPreview(doc);
            btnApply.Enabled = hasPreview;
            btnCancel.Enabled = hasPreview;
        }

        // Runs an action on the active document. Problems are shown to the user because Office
        // silently drops exceptions thrown from ribbon buttons.
        private void RunOnActiveDocument(Action<Word.Document> action)
        {
            try
            {
                Word.Document doc = GetActiveDocument();
                if (doc == null)
                {
                    ShowMessage("Avval hujjatni oçing va tahrirlaşga ruxsat bering.", MessageBoxIcon.Warning);
                }
                else if (doc.ProtectionType != Word.WdProtectionType.wdNoProtection)
                {
                    ShowMessage("Hujjat himoyalangan. Avval himoyani olib taşlang.", MessageBoxIcon.Warning);
                }
                else
                {
                    action(doc);
                }
            }
            catch (Exception ex)
            {
                ShowMessage("Xatolik: " + ex.Message, MessageBoxIcon.Error);
            }
            finally
            {
                RefreshButtons();
            }
        }

        private static void ShowMessage(string text, MessageBoxIcon icon)
        {
            MessageBox.Show(text, MessageTitle, MessageBoxButtons.OK, icon);
        }

        // A small square of the colour with a grey border, for the colour list.
        private static Image DrawSwatch(Color color)
        {
            var swatch = new Bitmap(16, 16);
            using (Graphics graphics = Graphics.FromImage(swatch))
            {
                graphics.Clear(color);
                graphics.DrawRectangle(Pens.Gray, 0, 0, 15, 15);
            }
            return swatch;
        }

        // A blue circle with a white "i", for the information button. Drawn here so it does not depend on which
        // built-in Office icons the installed Word has.
        private static Image DrawInfoIcon()
        {
            var icon = new Bitmap(32, 32);
            using (Graphics graphics = Graphics.FromImage(icon))
            using (var blue = new SolidBrush(Color.FromArgb(0, 114, 198)))
            {
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                graphics.FillEllipse(blue, 2, 2, 28, 28);
                graphics.FillEllipse(Brushes.White, 13.5f, 7, 5, 5);
                graphics.FillRectangle(Brushes.White, 14, 14, 4, 11);
            }
            return icon;
        }

        // The active document, or null when none is open or it is shown in Protected View.
        private static Word.Document GetActiveDocument()
        {
            Word.Application app = Globals.ThisAddIn.Application;
            try
            {
                return app.Documents.Count > 0 ? app.ActiveDocument : null;
            }
            catch (COMException)
            {
                return null;
            }
        }

        // The selected text, or null when nothing is selected so that the whole document is used.
        private static Word.Range GetSelectedText(Word.Document doc)
        {
            Word.Selection selection = doc.Application.Selection;
            bool hasText = selection.Type == Word.WdSelectionType.wdSelectionNormal && selection.Start < selection.End;
            return hasText ? selection.Range : null;
        }
    }
}
