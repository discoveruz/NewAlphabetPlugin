using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace NewAlphabetPlugin
{
    public partial class ThisAddIn
    {
        // Created here rather than in Startup so it already exists when the ribbon loads.
        internal readonly AlphabetConverter Converter = new AlphabetConverter();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Converter.PreviewColor = UserSettings.PreviewColor;
            Converter.ConvertOldLatin = UserSettings.ConvertOldLatin;
            Converter.ConvertCyrillic = UserSettings.ConvertCyrillic;
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        // Apply and Cancel follow whichever document is active.
        private void Application_WindowActivate(Word.Document doc, Word.Window window)
        {
            try
            {
                Globals.Ribbons.AlphabetRibbon.RefreshButtons();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        // Preview highlights are temporary, so take them off before the document is saved and closed.
        private void Application_DocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            try
            {
                Converter.Close(doc);
            }
            catch (Exception ex)
            {
                // Closing a document must never fail because of the add-in.
                Debug.WriteLine(ex);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
