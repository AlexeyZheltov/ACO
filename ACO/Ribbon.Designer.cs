namespace ACO
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором компонентов

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon));
            this.RbnTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.BtnCreateProgect = this.Factory.CreateRibbonButton();
            this.BtnLoadKP = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.BtnLoadLvl12 = this.Factory.CreateRibbonButton();
            this.BtnLoadLvl11 = this.Factory.CreateRibbonButton();
            this.BtnLoadLvl0 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.BtnUpdateLvl12 = this.Factory.CreateRibbonButton();
            this.BtnUpdateLvl11 = this.Factory.CreateRibbonButton();
            this.BtnUpdateLvl0 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.BtnAbout = this.Factory.CreateRibbonButton();
            this.RbnTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // RbnTab
            // 
            this.RbnTab.Groups.Add(this.group1);
            this.RbnTab.Groups.Add(this.group2);
            this.RbnTab.Groups.Add(this.group3);
            this.RbnTab.Groups.Add(this.group4);
            this.RbnTab.Label = "Спектрум";
            this.RbnTab.Name = "RbnTab";
            this.RbnTab.Position = this.Factory.RibbonPosition.AfterOfficeId("TabAddIns");
            // 
            // group1
            // 
            this.group1.Items.Add(this.BtnCreateProgect);
            this.group1.Items.Add(this.BtnLoadKP);
            this.group1.Label = "Создание";
            this.group1.Name = "group1";
            // 
            // BtnCreateProgect
            // 
            this.BtnCreateProgect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnCreateProgect.Image = ((System.Drawing.Image)(resources.GetObject("BtnCreateProgect.Image")));
            this.BtnCreateProgect.Label = "Создать";
            this.BtnCreateProgect.Name = "BtnCreateProgect";
            this.BtnCreateProgect.ShowImage = true;
            this.BtnCreateProgect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCreateProgect_Click);
            // 
            // BtnLoadKP
            // 
            this.BtnLoadKP.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnLoadKP.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadKP.Image")));
            this.BtnLoadKP.Label = "Загрузить КП";
            this.BtnLoadKP.Name = "BtnLoadKP";
            this.BtnLoadKP.ShowImage = true;
            this.BtnLoadKP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadKP_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.BtnLoadLvl12);
            this.group2.Items.Add(this.BtnLoadLvl11);
            this.group2.Items.Add(this.BtnLoadLvl0);
            this.group2.Label = "Загрузить";
            this.group2.Name = "group2";
            // 
            // BtnLoadLvl12
            // 
            this.BtnLoadLvl12.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadLvl12.Image")));
            this.BtnLoadLvl12.Label = "Урв 12";
            this.BtnLoadLvl12.Name = "BtnLoadLvl12";
            this.BtnLoadLvl12.ShowImage = true;
            // 
            // BtnLoadLvl11
            // 
            this.BtnLoadLvl11.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadLvl11.Image")));
            this.BtnLoadLvl11.Label = "Урв 11";
            this.BtnLoadLvl11.Name = "BtnLoadLvl11";
            this.BtnLoadLvl11.ShowImage = true;
            // 
            // BtnLoadLvl0
            // 
            this.BtnLoadLvl0.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadLvl0.Image")));
            this.BtnLoadLvl0.Label = "Урв 0";
            this.BtnLoadLvl0.Name = "BtnLoadLvl0";
            this.BtnLoadLvl0.ShowImage = true;
            // 
            // group3
            // 
            this.group3.Items.Add(this.BtnUpdateLvl12);
            this.group3.Items.Add(this.BtnUpdateLvl11);
            this.group3.Items.Add(this.BtnUpdateLvl0);
            this.group3.Label = "Обновить";
            this.group3.Name = "group3";
            // 
            // BtnUpdateLvl12
            // 
            this.BtnUpdateLvl12.Image = ((System.Drawing.Image)(resources.GetObject("BtnUpdateLvl12.Image")));
            this.BtnUpdateLvl12.Label = "Урв 12";
            this.BtnUpdateLvl12.Name = "BtnUpdateLvl12";
            this.BtnUpdateLvl12.ShowImage = true;
            // 
            // BtnUpdateLvl11
            // 
            this.BtnUpdateLvl11.Image = ((System.Drawing.Image)(resources.GetObject("BtnUpdateLvl11.Image")));
            this.BtnUpdateLvl11.Label = "Урв 11";
            this.BtnUpdateLvl11.Name = "BtnUpdateLvl11";
            this.BtnUpdateLvl11.ShowImage = true;
            // 
            // BtnUpdateLvl0
            // 
            this.BtnUpdateLvl0.Image = ((System.Drawing.Image)(resources.GetObject("BtnUpdateLvl0.Image")));
            this.BtnUpdateLvl0.Label = "Урв 0";
            this.BtnUpdateLvl0.Name = "BtnUpdateLvl0";
            this.BtnUpdateLvl0.ShowImage = true;
            // 
            // group4
            // 
            this.group4.Items.Add(this.BtnAbout);
            this.group4.Label = "Информация";
            this.group4.Name = "group4";
            // 
            // BtnAbout
            // 
            this.BtnAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnAbout.Image = ((System.Drawing.Image)(resources.GetObject("BtnAbout.Image")));
            this.BtnAbout.Label = "О программе";
            this.BtnAbout.Name = "BtnAbout";
            this.BtnAbout.ShowImage = true;
            this.BtnAbout.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnAbout_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.RbnTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.RbnTab.ResumeLayout(false);
            this.RbnTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCreateProgect;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadLvl12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadLvl11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadLvl0;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUpdateLvl12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUpdateLvl11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUpdateLvl0;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadKP;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab RbnTab;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
