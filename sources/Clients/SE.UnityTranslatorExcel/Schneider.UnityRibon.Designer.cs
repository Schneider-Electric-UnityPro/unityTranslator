using Microsoft.Office.Tools.Ribbon;
using SchneiderElectric.UnityComments;
using System.Collections.Generic;

namespace SE.UnityCommentsExcel
{
    partial class UnityCommentRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public UnityCommentRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            foreach (var lang in UnityApplicationComments.Translator.LanguagesNames)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = lang;
                TargetLang.Items.Add(item);
            }
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UnityCommentRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.UnityCommentsGroup = this.Factory.CreateRibbonGroup();
            this.open = this.Factory.CreateRibbonButton();
            this.Save = this.Factory.CreateRibbonButton();
            this.xmlgroup = this.Factory.CreateRibbonGroup();
            this.Import = this.Factory.CreateRibbonButton();
            this.Export = this.Factory.CreateRibbonButton();
            this.AutoTranslate = this.Factory.CreateRibbonGroup();
            this.Tranlate = this.Factory.CreateRibbonButton();
            this.TargetLang = this.Factory.CreateRibbonComboBox();
            this.Erase = this.Factory.CreateRibbonCheckBox();
            this.selection = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.UnityCommentsGroup.SuspendLayout();
            this.xmlgroup.SuspendLayout();
            this.AutoTranslate.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.UnityCommentsGroup);
            this.tab1.Groups.Add(this.xmlgroup);
            this.tab1.Groups.Add(this.AutoTranslate);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // UnityCommentsGroup
            // 
            this.UnityCommentsGroup.Items.Add(this.open);
            this.UnityCommentsGroup.Items.Add(this.Save);
            this.UnityCommentsGroup.Label = "Unity application";
            this.UnityCommentsGroup.Name = "UnityCommentsGroup";
            // 
            // open
            // 
            this.open.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.open.Image = ((System.Drawing.Image)(resources.GetObject("open.Image")));
            this.open.Label = "Open";
            this.open.Name = "open";
            this.open.ScreenTip = "open unity file";
            this.open.ShowImage = true;
            this.open.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnOpen);
            // 
            // Save
            // 
            this.Save.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Save.Image = ((System.Drawing.Image)(resources.GetObject("Save.Image")));
            this.Save.Label = "Save";
            this.Save.Name = "Save";
            this.Save.ScreenTip = "saves the curent translations into the unity application";
            this.Save.ShowImage = true;
            this.Save.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnSave);
            // 
            // xmlgroup
            // 
            this.xmlgroup.Items.Add(this.Import);
            this.xmlgroup.Items.Add(this.Export);
            this.xmlgroup.Label = "XML project";
            this.xmlgroup.Name = "xmlgroup";
            // 
            // Import
            // 
            this.Import.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Import.Image = ((System.Drawing.Image)(resources.GetObject("Import.Image")));
            this.Import.Label = "Import XML";
            this.Import.Name = "Import";
            this.Import.ScreenTip = "import unity commments xml file";
            this.Import.ShowImage = true;
            this.Import.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnImport);
            // 
            // Export
            // 
            this.Export.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Export.Image = ((System.Drawing.Image)(resources.GetObject("Export.Image")));
            this.Export.Label = "Export XML";
            this.Export.Name = "Export";
            this.Export.ScreenTip = "export the curent project to xml";
            this.Export.ShowImage = true;
            this.Export.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnExport);
            // 
            // AutoTranslate
            // 
            this.AutoTranslate.Items.Add(this.Tranlate);
            this.AutoTranslate.Items.Add(this.TargetLang);
            this.AutoTranslate.Items.Add(this.Erase);
            this.AutoTranslate.Items.Add(this.selection);
            this.AutoTranslate.Label = "Automatic Translation";
            this.AutoTranslate.Name = "AutoTranslate";
            // 
            // Tranlate
            // 
            this.Tranlate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Tranlate.Image = ((System.Drawing.Image)(resources.GetObject("Tranlate.Image")));
            this.Tranlate.Label = "Translate";
            this.Tranlate.Name = "Tranlate";
            this.Tranlate.ScreenTip = "Translates the source to the selected language";
            this.Tranlate.ShowImage = true;
            this.Tranlate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnTranslate);
            // 
            // TargetLang
            // 
            this.TargetLang.Label = "Language";
            this.TargetLang.Name = "TargetLang";
            this.TargetLang.ScreenTip = "target language";
            this.TargetLang.Text = null;
            this.TargetLang.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnTargetLanguagechange);
            // 
            // Erase
            // 
            this.Erase.Label = "Erase existing translations";
            this.Erase.Name = "Erase";
            this.Erase.ScreenTip = "when checked, will translate all comments , even if a translation exists.";
            this.Erase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnEraseModeChange);
            // 
            // selection
            // 
            this.selection.Label = "Selection only";
            this.selection.Name = "selection";
            this.selection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.OnFilterSelectionChange);
            // 
            // UnityCommentRibbon
            // 
            this.Name = "UnityCommentRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.OnLoad);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.UnityCommentsGroup.ResumeLayout(false);
            this.UnityCommentsGroup.PerformLayout();
            this.xmlgroup.ResumeLayout(false);
            this.xmlgroup.PerformLayout();
            this.AutoTranslate.ResumeLayout(false);
            this.AutoTranslate.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup UnityCommentsGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton open;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup xmlgroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup AutoTranslate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Save;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Import;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Export;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Tranlate;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox TargetLang;
        internal RibbonCheckBox Erase;
        internal RibbonCheckBox selection;
    }

    partial class ThisRibbonCollection
    {
        internal UnityCommentRibbon Schneider
        {
            get { return this.GetRibbon<UnityCommentRibbon>(); }
        }
    }
}
