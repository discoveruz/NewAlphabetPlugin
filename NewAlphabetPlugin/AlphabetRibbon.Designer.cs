namespace NewAlphabetPlugin
{
    partial class AlphabetRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AlphabetRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabAlphabet = this.Factory.CreateRibbonTab();
            this.grpConvert = this.Factory.CreateRibbonGroup();
            this.btnPreview = this.Factory.CreateRibbonButton();
            this.btnApply = this.Factory.CreateRibbonButton();
            this.btnCancel = this.Factory.CreateRibbonButton();
            this.grpScripts = this.Factory.CreateRibbonGroup();
            this.chkOldLatin = this.Factory.CreateRibbonCheckBox();
            this.chkCyrillic = this.Factory.CreateRibbonCheckBox();
            this.grpColor = this.Factory.CreateRibbonGroup();
            this.ddColor = this.Factory.CreateRibbonDropDown();
            this.grpAbout = this.Factory.CreateRibbonGroup();
            this.btnAbout = this.Factory.CreateRibbonButton();
            this.tabAlphabet.SuspendLayout();
            this.grpConvert.SuspendLayout();
            this.grpScripts.SuspendLayout();
            this.grpColor.SuspendLayout();
            this.grpAbout.SuspendLayout();
            this.SuspendLayout();
            //
            // tabAlphabet
            //
            this.tabAlphabet.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Custom;
            this.tabAlphabet.Groups.Add(this.grpConvert);
            this.tabAlphabet.Groups.Add(this.grpScripts);
            this.tabAlphabet.Groups.Add(this.grpColor);
            this.tabAlphabet.Groups.Add(this.grpAbout);
            this.tabAlphabet.Label = "Yangi alifbo";
            this.tabAlphabet.Name = "tabAlphabet";
            this.tabAlphabet.Position = this.Factory.RibbonPosition.AfterOfficeId("TabHome");
            //
            // grpConvert
            //
            this.grpConvert.Items.Add(this.btnPreview);
            this.grpConvert.Items.Add(this.btnApply);
            this.grpConvert.Items.Add(this.btnCancel);
            this.grpConvert.Label = "Almaştiriş";
            this.grpConvert.Name = "grpConvert";
            //
            // btnPreview
            //
            this.btnPreview.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPreview.Label = "Körib çiqiş";
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.OfficeImageId = "TextHighlightColorPicker";
            this.btnPreview.ScreenTip = "Körib çiqiş";
            this.btnPreview.ShowImage = true;
            this.btnPreview.SuperTip = "Tanlangan matnni (heç narsa tanlanmagan bölsa, butun hujjatni) tekşiradi va özgaradigan " +
    "joylarni rang bilan belgilaydi: eski lotin harflarini (Oʻ→Ö, Gʻ→Ğ, Sh→Ş, Ch→Ç) va kirillça sözlarni (Тошкент→Toşkent).";
            this.btnPreview.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPreview_Click);
            //
            // btnApply
            //
            this.btnApply.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnApply.Enabled = false;
            this.btnApply.Label = "Qöllaş";
            this.btnApply.Name = "btnApply";
            this.btnApply.OfficeImageId = "ReviewAcceptChange";
            this.btnApply.ScreenTip = "Qöllaş";
            this.btnApply.ShowImage = true;
            this.btnApply.SuperTip = "Hali ham belgilangan harflar va sözlarni yangi alifboga ötkazadi va belgilarni olib taşlaydi. " +
    "Biror joyni özgartirmaslik uçun, avval undagi rangni olib taşlang (Text Highlight Color → No Color).";
            this.btnApply.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnApply_Click);
            //
            // btnCancel
            //
            this.btnCancel.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCancel.Enabled = false;
            this.btnCancel.Label = "Bekor qiliş";
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.OfficeImageId = "ReviewRejectChange";
            this.btnCancel.ScreenTip = "Bekor qiliş";
            this.btnCancel.ShowImage = true;
            this.btnCancel.SuperTip = "Belgilarni olib taşlaydi, matn özgarmaydi.";
            this.btnCancel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCancel_Click);
            //
            // grpScripts
            //
            this.grpScripts.Items.Add(this.chkOldLatin);
            this.grpScripts.Items.Add(this.chkCyrillic);
            this.grpScripts.Label = "Qaysi yozuvdan";
            this.grpScripts.Name = "grpScripts";
            //
            // chkOldLatin
            //
            this.chkOldLatin.Checked = true;
            this.chkOldLatin.Label = "Eski lotin";
            this.chkOldLatin.Name = "chkOldLatin";
            this.chkOldLatin.ScreenTip = "Eski lotin";
            this.chkOldLatin.SuperTip = "Körib çiqiş eski lotin harflarini topadi: Oʻ→Ö, Gʻ→Ğ, Sh→Ş, Ch→Ç.";
            this.chkOldLatin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkScript_Click);
            //
            // chkCyrillic
            //
            this.chkCyrillic.Checked = true;
            this.chkCyrillic.Label = "Kirill";
            this.chkCyrillic.Name = "chkCyrillic";
            this.chkCyrillic.ScreenTip = "Kirill";
            this.chkCyrillic.SuperTip = "Körib çiqiş özbekça kirill yozuvidagi sözlarni topadi: Тошкент→Toşkent. Özbek alifbosida " +
    "yöq harfli sözlar (masalan, Щ yoki Ы bilan) özgarmaydi. Rusça matnni özgartirmaslik uçun undagi rangni olib taşlang.";
            this.chkCyrillic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkScript_Click);
            //
            // grpColor
            //
            this.grpColor.Items.Add(this.ddColor);
            this.grpColor.Label = "Belgilaş rangi";
            this.grpColor.Name = "grpColor";
            //
            // ddColor
            //
            this.ddColor.Label = "Belgilaş rangi";
            this.ddColor.Name = "ddColor";
            this.ddColor.ScreenTip = "Belgilaş rangi";
            this.ddColor.ShowItemImage = true;
            this.ddColor.ShowLabel = false;
            this.ddColor.SizeString = "Töq kulrangWW";
            this.ddColor.SuperTip = "Körib çiqiş özgaradigan harflarni şu rang bilan belgilaydi.";
            this.ddColor.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddColor_SelectionChanged);
            //
            // grpAbout
            //
            this.grpAbout.Items.Add(this.btnAbout);
            this.grpAbout.Label = "Dastur haqida";
            this.grpAbout.Name = "grpAbout";
            //
            // btnAbout
            //
            this.btnAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAbout.Label = "Maʼlumot";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.ScreenTip = "Maʼlumot";
            this.btnAbout.ShowImage = true;
            this.btnAbout.SuperTip = "Dastur muallifi va u bilan boğlaniş yöli.";
            this.btnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAbout_Click);
            //
            // AlphabetRibbon
            //
            this.Name = "AlphabetRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabAlphabet);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AlphabetRibbon_Load);
            this.tabAlphabet.ResumeLayout(false);
            this.tabAlphabet.PerformLayout();
            this.grpConvert.ResumeLayout(false);
            this.grpConvert.PerformLayout();
            this.grpScripts.ResumeLayout(false);
            this.grpScripts.PerformLayout();
            this.grpColor.ResumeLayout(false);
            this.grpColor.PerformLayout();
            this.grpAbout.ResumeLayout(false);
            this.grpAbout.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabAlphabet;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPreview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApply;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCancel;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpScripts;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkOldLatin;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkCyrillic;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddColor;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
    }

    partial class ThisRibbonCollection
    {
        internal AlphabetRibbon AlphabetRibbon
        {
            get { return this.GetRibbon<AlphabetRibbon>(); }
        }
    }
}
