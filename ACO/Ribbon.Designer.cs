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
            this.BtnSpectrum = this.Factory.CreateRibbonButton();
            this.BtnLoadKP = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.BtnUpdateFormuls = this.Factory.CreateRibbonButton();
            this.SptBtnFormatComments = this.Factory.CreateRibbonSplitButton();
            this.BtnFormatComments = this.Factory.CreateRibbonButton();
            this.BtnCol = this.Factory.CreateRibbonButton();
            this.BtnClearFormateContions = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.BtnLoadLvl12 = this.Factory.CreateRibbonButton();
            this.BtnLoadLvl11 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.BtnUpdateLvl12 = this.Factory.CreateRibbonButton();
            this.BtnUpdateLvl11 = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.BtnProjectManager = this.Factory.CreateRibbonButton();
            this.BtnKP = this.Factory.CreateRibbonButton();
            this.BtnSettings = this.Factory.CreateRibbonButton();
            this.BtnExcelScreenUpdating = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.BtnAbout = this.Factory.CreateRibbonButton();
            this.RbnTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.group6.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group5.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // RbnTab
            // 
            this.RbnTab.Groups.Add(this.group1);
            this.RbnTab.Groups.Add(this.group6);
            this.RbnTab.Groups.Add(this.group2);
            this.RbnTab.Groups.Add(this.group3);
            this.RbnTab.Groups.Add(this.group5);
            this.RbnTab.Groups.Add(this.group4);
            this.RbnTab.Label = "Спектрум";
            this.RbnTab.Name = "RbnTab";
            this.RbnTab.Position = this.Factory.RibbonPosition.AfterOfficeId("TabAddIns");
            // 
            // group1
            // 
            this.group1.Items.Add(this.BtnCreateProgect);
            this.group1.Items.Add(this.BtnSpectrum);
            this.group1.Items.Add(this.BtnLoadKP);
            this.group1.Label = "Создание";
            this.group1.Name = "group1";
            // 
            // BtnCreateProgect
            // 
            this.BtnCreateProgect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnCreateProgect.Image = ((System.Drawing.Image)(resources.GetObject("BtnCreateProgect.Image")));
            this.BtnCreateProgect.Label = "Открыть шаблон";
            this.BtnCreateProgect.Name = "BtnCreateProgect";
            this.BtnCreateProgect.ShowImage = true;
            this.BtnCreateProgect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCreateProgect_Click);
            // 
            // BtnSpectrum
            // 
            this.BtnSpectrum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnSpectrum.Image = ((System.Drawing.Image)(resources.GetObject("BtnSpectrum.Image")));
            this.BtnSpectrum.Label = "Загрузить базовую оценку";
            this.BtnSpectrum.Name = "BtnSpectrum";
            this.BtnSpectrum.ShowImage = true;
            this.BtnSpectrum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnBaseEstimate_Click);
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
            // group6
            // 
            this.group6.Items.Add(this.BtnUpdateFormuls);
            this.group6.Items.Add(this.SptBtnFormatComments);
            this.group6.Label = "Формат";
            this.group6.Name = "group6";
            // 
            // BtnUpdateFormuls
            // 
            this.BtnUpdateFormuls.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnUpdateFormuls.Image = ((System.Drawing.Image)(resources.GetObject("BtnUpdateFormuls.Image")));
            this.BtnUpdateFormuls.Label = "Обновление формул";
            this.BtnUpdateFormuls.Name = "BtnUpdateFormuls";
            this.BtnUpdateFormuls.ShowImage = true;
            this.BtnUpdateFormuls.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUpdateFormuls_Click);
            // 
            // SptBtnFormatComments
            // 
            this.SptBtnFormatComments.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SptBtnFormatComments.Image = ((System.Drawing.Image)(resources.GetObject("SptBtnFormatComments.Image")));
            this.SptBtnFormatComments.Items.Add(this.BtnFormatComments);
            this.SptBtnFormatComments.Items.Add(this.BtnCol);
            this.SptBtnFormatComments.Items.Add(this.BtnClearFormateContions);
            this.SptBtnFormatComments.Label = "Форматирование комментариев";
            this.SptBtnFormatComments.Name = "SptBtnFormatComments";
            this.SptBtnFormatComments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SptBtnFormatComments_Click);
            // 
            // BtnFormatComments
            // 
            this.BtnFormatComments.Image = ((System.Drawing.Image)(resources.GetObject("BtnFormatComments.Image")));
            this.BtnFormatComments.Label = "Настройки формата";
            this.BtnFormatComments.Name = "BtnFormatComments";
            this.BtnFormatComments.ShowImage = true;
            this.BtnFormatComments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFormatComments_Click);
            // 
            // BtnCol
            // 
            this.BtnCol.Image = ((System.Drawing.Image)(resources.GetObject("BtnCol.Image")));
            this.BtnCol.Label = "Применить к выделеному";
            this.BtnCol.Name = "BtnCol";
            this.BtnCol.ShowImage = true;
            this.BtnCol.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCol_Click);
            // 
            // BtnClearFormateContions
            // 
            this.BtnClearFormateContions.Image = ((System.Drawing.Image)(resources.GetObject("BtnClearFormateContions.Image")));
            this.BtnClearFormateContions.Label = "Очистить";
            this.BtnClearFormateContions.Name = "BtnClearFormateContions";
            this.BtnClearFormateContions.ShowImage = true;
            this.BtnClearFormateContions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnClearFormateContions_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.BtnLoadLvl12);
            this.group2.Items.Add(this.BtnLoadLvl11);
            this.group2.Label = "Загрузить";
            this.group2.Name = "group2";
            // 
            // BtnLoadLvl12
            // 
            this.BtnLoadLvl12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnLoadLvl12.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadLvl12.Image")));
            this.BtnLoadLvl12.Label = "Урв 12";
            this.BtnLoadLvl12.Name = "BtnLoadLvl12";
            this.BtnLoadLvl12.ShowImage = true;
            this.BtnLoadLvl12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadLvl12_Click);
            // 
            // BtnLoadLvl11
            // 
            this.BtnLoadLvl11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnLoadLvl11.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadLvl11.Image")));
            this.BtnLoadLvl11.Label = "Урв 11";
            this.BtnLoadLvl11.Name = "BtnLoadLvl11";
            this.BtnLoadLvl11.ShowImage = true;
            this.BtnLoadLvl11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadLvl11_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.BtnUpdateLvl12);
            this.group3.Items.Add(this.BtnUpdateLvl11);
            this.group3.Label = "Обновить";
            this.group3.Name = "group3";
            // 
            // BtnUpdateLvl12
            // 
            this.BtnUpdateLvl12.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnUpdateLvl12.Image = ((System.Drawing.Image)(resources.GetObject("BtnUpdateLvl12.Image")));
            this.BtnUpdateLvl12.Label = "Урв 12";
            this.BtnUpdateLvl12.Name = "BtnUpdateLvl12";
            this.BtnUpdateLvl12.ShowImage = true;
            this.BtnUpdateLvl12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUpdateLvl12_Click);
            // 
            // BtnUpdateLvl11
            // 
            this.BtnUpdateLvl11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnUpdateLvl11.Image = ((System.Drawing.Image)(resources.GetObject("BtnUpdateLvl11.Image")));
            this.BtnUpdateLvl11.Label = "Урв 11";
            this.BtnUpdateLvl11.Name = "BtnUpdateLvl11";
            this.BtnUpdateLvl11.ShowImage = true;
            this.BtnUpdateLvl11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUpdateLvl11_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.BtnProjectManager);
            this.group5.Items.Add(this.BtnKP);
            this.group5.Items.Add(this.BtnSettings);
            this.group5.Items.Add(this.BtnExcelScreenUpdating);
            this.group5.Label = "Настройки";
            this.group5.Name = "group5";
            // 
            // BtnProjectManager
            // 
            this.BtnProjectManager.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnProjectManager.Image = ((System.Drawing.Image)(resources.GetObject("BtnProjectManager.Image")));
            this.BtnProjectManager.Label = "Диспетчер проекта";
            this.BtnProjectManager.Name = "BtnProjectManager";
            this.BtnProjectManager.ShowImage = true;
            this.BtnProjectManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnProjectManager_Click);
            // 
            // BtnKP
            // 
            this.BtnKP.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnKP.Image = ((System.Drawing.Image)(resources.GetObject("BtnKP.Image")));
            this.BtnKP.Label = "Диспетчер КП";
            this.BtnKP.Name = "BtnKP";
            this.BtnKP.ShowImage = true;
            this.BtnKP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnKP_Click);
            // 
            // BtnSettings
            // 
            this.BtnSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnSettings.Image = ((System.Drawing.Image)(resources.GetObject("BtnSettings.Image")));
            this.BtnSettings.Label = "Настройки";
            this.BtnSettings.Name = "BtnSettings";
            this.BtnSettings.ShowImage = true;
            this.BtnSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSettings_Click);
            // 
            // BtnExcelScreenUpdating
            // 
            this.BtnExcelScreenUpdating.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnExcelScreenUpdating.Image = ((System.Drawing.Image)(resources.GetObject("BtnExcelScreenUpdating.Image")));
            this.BtnExcelScreenUpdating.Label = "Включить обновление Excel";
            this.BtnExcelScreenUpdating.Name = "BtnExcelScreenUpdating";
            this.BtnExcelScreenUpdating.ShowImage = true;
            this.BtnExcelScreenUpdating.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnExcelScreenUpdating_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.BtnAbout);
            this.group4.Label = "Информация";
            this.group4.Name = "group4";
            this.group4.Visible = false;
            // 
            // BtnAbout
            // 
            this.BtnAbout.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnAbout.Image = ((System.Drawing.Image)(resources.GetObject("BtnAbout.Image")));
            this.BtnAbout.Label = "О программе";
            this.BtnAbout.Name = "BtnAbout";
            this.BtnAbout.ShowImage = true;
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
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUpdateLvl12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUpdateLvl11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadKP;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnAbout;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab RbnTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnProjectManager;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUpdateFormuls;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnKP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSpectrum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnExcelScreenUpdating;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton SptBtnFormatComments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFormatComments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCol;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnClearFormateContions;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
