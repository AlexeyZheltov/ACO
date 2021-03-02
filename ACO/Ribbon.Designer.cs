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
            this.RibbonTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.BtnCreateProgect = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.BtnLoadKP = this.Factory.CreateRibbonButton();
            this.RibbonTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // RibbonTab
            // 
            this.RibbonTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.RibbonTab.Groups.Add(this.group1);
            this.RibbonTab.Groups.Add(this.group2);
            this.RibbonTab.Groups.Add(this.group3);
            this.RibbonTab.Label = "Спектрум";
            this.RibbonTab.Name = "RibbonTab";
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
            this.BtnCreateProgect.Label = "button1";
            this.BtnCreateProgect.Name = "BtnCreateProgect";
            this.BtnCreateProgect.ShowImage = true;
            // 
            // group2
            // 
            this.group2.Items.Add(this.button2);
            this.group2.Items.Add(this.button3);
            this.group2.Items.Add(this.button4);
            this.group2.Label = "Загрузить";
            this.group2.Name = "group2";
            // 
            // button2
            // 
            this.button2.Label = "Урв 12";
            this.button2.Name = "button2";
            // 
            // button3
            // 
            this.button3.Label = "Урв 11";
            this.button3.Name = "button3";
            // 
            // button4
            // 
            this.button4.Label = "Урв 0";
            this.button4.Name = "button4";
            // 
            // group3
            // 
            this.group3.Items.Add(this.button5);
            this.group3.Items.Add(this.button6);
            this.group3.Items.Add(this.button7);
            this.group3.Label = "Обновить";
            this.group3.Name = "group3";
            // 
            // button5
            // 
            this.button5.Label = "Урв 12";
            this.button5.Name = "button5";
            // 
            // button6
            // 
            this.button6.Label = "Урв 11";
            this.button6.Name = "button6";
            // 
            // button7
            // 
            this.button7.Label = "Урв 0";
            this.button7.Name = "button7";
            // 
            // BtnLoadKP
            // 
            this.BtnLoadKP.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnLoadKP.Label = "button1";
            this.BtnLoadKP.Name = "BtnLoadKP";
            this.BtnLoadKP.ShowImage = true;
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.RibbonTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.RibbonTab.ResumeLayout(false);
            this.RibbonTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab RibbonTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCreateProgect;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadKP;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
