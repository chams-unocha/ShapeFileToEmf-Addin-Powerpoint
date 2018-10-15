namespace AddinPowerpointShapefileToEmf
{
    partial class AutoEmfRbbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AutoEmfRbbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupEmfGenerator = this.Factory.CreateRibbonGroup();
            this.buttonLaunch = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.groupEmfGenerator.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.groupEmfGenerator);
            this.tab1.Label = "Shapefile to Emf";
            this.tab1.Name = "tab1";
            // 
            // groupEmfGenerator
            // 
            this.groupEmfGenerator.Items.Add(this.buttonLaunch);
            this.groupEmfGenerator.Label = "Shapefile to Emf";
            this.groupEmfGenerator.Name = "groupEmfGenerator";
            // 
            // buttonLaunch
            // 
            this.buttonLaunch.Image = global::AddinPowerpointShapefileToEmf.Properties.Resources.map;
            this.buttonLaunch.Label = "Finalize convert";
            this.buttonLaunch.Name = "buttonLaunch";
            this.buttonLaunch.ShowImage = true;
            this.buttonLaunch.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLaunch_Click);
            // 
            // AutoEmfRbbon
            // 
            this.Name = "AutoEmfRbbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AutoEmfRbbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.groupEmfGenerator.ResumeLayout(false);
            this.groupEmfGenerator.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupEmfGenerator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLaunch;
    }

    partial class ThisRibbonCollection
    {
        internal AutoEmfRbbon AutoEmfRbbon
        {
            get { return this.GetRibbon<AutoEmfRbbon>(); }
        }
    }
}
