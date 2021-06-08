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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.RbnTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.BtnCreateProgect = this.Factory.CreateRibbonButton();
            this.BtnSpectrum = this.Factory.CreateRibbonButton();
            this.BtnLoadKP = this.Factory.CreateRibbonButton();
            this.SptBtn = this.Factory.CreateRibbonSplitButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.BtnLoadLvl12 = this.Factory.CreateRibbonButton();
            this.BtnLoadLvl11 = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.BtnUpdateLvl12 = this.Factory.CreateRibbonButton();
            this.BtnUpdateLvl11 = this.Factory.CreateRibbonButton();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.SptBtnUpdateFormate = this.Factory.CreateRibbonSplitButton();
            this.BtnFormatNumber = this.Factory.CreateRibbonButton();
            this.BtnDataFilter = this.Factory.CreateRibbonButton();
            this.SptBtnGroup = this.Factory.CreateRibbonSplitButton();
            this.BtnGroupColumns = this.Factory.CreateRibbonButton();
            this.BtnUngroupColumns = this.Factory.CreateRibbonButton();
            this.BtnGroupRows = this.Factory.CreateRibbonButton();
            this.BtnUngroupRows = this.Factory.CreateRibbonButton();
            this.SptBtnFormatComments = this.Factory.CreateRibbonSplitButton();
            this.BtnSetFormul = this.Factory.CreateRibbonButton();
            this.BtnFormatComments = this.Factory.CreateRibbonButton();
            this.BtnClearFormateContions = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.BtnProjectManager = this.Factory.CreateRibbonButton();
            this.BtnKP = this.Factory.CreateRibbonButton();
            this.BtnSettings = this.Factory.CreateRibbonButton();
            this.BtnExcelScreenUpdating = this.Factory.CreateRibbonButton();
            this.comboBoxLvlCost = this.Factory.CreateRibbonComboBox();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.BtnAbout = this.Factory.CreateRibbonButton();
            this.RbnTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.group6.SuspendLayout();
            this.group5.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // RbnTab
            // 
            this.RbnTab.Groups.Add(this.group1);
            this.RbnTab.Groups.Add(this.group6);
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
            this.group1.Items.Add(this.SptBtn);
            this.group1.Label = "Создание";
            this.group1.Name = "group1";
            // 
            // BtnCreateProgect
            // 
            this.BtnCreateProgect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnCreateProgect.Image = ((System.Drawing.Image)(resources.GetObject("BtnCreateProgect.Image")));
            this.BtnCreateProgect.Label = "Открыть шаблон";
            this.BtnCreateProgect.Name = "BtnCreateProgect";
            this.BtnCreateProgect.ScreenTip = "Создать новый проект на основе шаблона.";
            this.BtnCreateProgect.ShowImage = true;
            this.BtnCreateProgect.SuperTip = "Укажите в настройках файл шаблона.";
            this.BtnCreateProgect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnCreateProject_Click);
            // 
            // BtnSpectrum
            // 
            this.BtnSpectrum.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnSpectrum.Image = ((System.Drawing.Image)(resources.GetObject("BtnSpectrum.Image")));
            this.BtnSpectrum.Label = "Загрузить базовую оценку";
            this.BtnSpectrum.Name = "BtnSpectrum";
            this.BtnSpectrum.ScreenTip = "Загрузка списка базовой оценки из файла.";
            this.BtnSpectrum.ShowImage = true;
            this.BtnSpectrum.SuperTip = "Укажите столбцы в настройках. Выберите файл. Выберите настройки столбцов.";
            this.BtnSpectrum.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnBaseEstimate_Click);
            // 
            // BtnLoadKP
            // 
            this.BtnLoadKP.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnLoadKP.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadKP.Image")));
            this.BtnLoadKP.Label = "Загрузить КП";
            this.BtnLoadKP.Name = "BtnLoadKP";
            this.BtnLoadKP.ScreenTip = "Сопоставление с базовой оценкой списка КП .";
            this.BtnLoadKP.ShowImage = true;
            this.BtnLoadKP.SuperTip = "Укажите столбцы в настройках КП . Выберите файл. Выберите настройки столбцов.";
            this.BtnLoadKP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadKP_Click);
            // 
            // SptBtn
            // 
            this.SptBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SptBtn.Image = ((System.Drawing.Image)(resources.GetObject("SptBtn.Image")));
            this.SptBtn.Items.Add(this.separator2);
            this.SptBtn.Items.Add(this.BtnLoadLvl12);
            this.SptBtn.Items.Add(this.BtnLoadLvl11);
            this.SptBtn.Items.Add(this.separator1);
            this.SptBtn.Items.Add(this.BtnUpdateLvl12);
            this.SptBtn.Items.Add(this.BtnUpdateLvl11);
            this.SptBtn.Label = "Итоги";
            this.SptBtn.Name = "SptBtn";
            this.SptBtn.SuperTip = "Подготовить сводные данные на листах Урв12, Урв11";
            this.SptBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SptBtn_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            this.separator2.Title = "Загрузить";
            // 
            // BtnLoadLvl12
            // 
            this.BtnLoadLvl12.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadLvl12.Image")));
            this.BtnLoadLvl12.Label = "Урв 12";
            this.BtnLoadLvl12.Name = "BtnLoadLvl12";
            this.BtnLoadLvl12.ShowImage = true;
            this.BtnLoadLvl12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadLvl12_Click);
            // 
            // BtnLoadLvl11
            // 
            this.BtnLoadLvl11.Image = ((System.Drawing.Image)(resources.GetObject("BtnLoadLvl11.Image")));
            this.BtnLoadLvl11.Label = "Урв 11";
            this.BtnLoadLvl11.Name = "BtnLoadLvl11";
            this.BtnLoadLvl11.ShowImage = true;
            this.BtnLoadLvl11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnLoadLvl11_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            this.separator1.Title = "Обновить";
            // 
            // BtnUpdateLvl12
            // 
            this.BtnUpdateLvl12.Image = ((System.Drawing.Image)(resources.GetObject("BtnUpdateLvl12.Image")));
            this.BtnUpdateLvl12.Label = "Урв 12";
            this.BtnUpdateLvl12.Name = "BtnUpdateLvl12";
            this.BtnUpdateLvl12.ShowImage = true;
            this.BtnUpdateLvl12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUpdateLvl12_Click);
            // 
            // BtnUpdateLvl11
            // 
            this.BtnUpdateLvl11.Image = ((System.Drawing.Image)(resources.GetObject("BtnUpdateLvl11.Image")));
            this.BtnUpdateLvl11.Label = "Урв 11";
            this.BtnUpdateLvl11.Name = "BtnUpdateLvl11";
            this.BtnUpdateLvl11.ShowImage = true;
            this.BtnUpdateLvl11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUpdateLvl11_Click);
            // 
            // group6
            // 
            this.group6.Items.Add(this.SptBtnUpdateFormate);
            this.group6.Items.Add(this.SptBtnGroup);
            this.group6.Items.Add(this.SptBtnFormatComments);
            this.group6.Label = "Формат";
            this.group6.Name = "group6";
            // 
            // SptBtnUpdateFormate
            // 
            this.SptBtnUpdateFormate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SptBtnUpdateFormate.Image = ((System.Drawing.Image)(resources.GetObject("SptBtnUpdateFormate.Image")));
            this.SptBtnUpdateFormate.Items.Add(this.BtnFormatNumber);
            this.SptBtnUpdateFormate.Items.Add(this.BtnDataFilter);
            this.SptBtnUpdateFormate.Label = "Обновить формат списка";
            this.SptBtnUpdateFormate.Name = "SptBtnUpdateFormate";
            this.SptBtnUpdateFormate.ScreenTip = "Обновляет формат таблицы на листе Анализ";
            this.SptBtnUpdateFormate.SuperTip = "Обновляет нумерацию, Группирует данные, Устанавливает цвет строк в зависимости от" +
    " уровня списка";
            this.SptBtnUpdateFormate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateFormateList_Click);
            // 
            // BtnFormatNumber
            // 
            this.BtnFormatNumber.Image = ((System.Drawing.Image)(resources.GetObject("BtnFormatNumber.Image")));
            this.BtnFormatNumber.Label = "Формат ячеек";
            this.BtnFormatNumber.Name = "BtnFormatNumber";
            this.BtnFormatNumber.ShowImage = true;
            this.BtnFormatNumber.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFormatNumber_Click);
            // 
            // BtnDataFilter
            // 
            this.BtnDataFilter.Image = ((System.Drawing.Image)(resources.GetObject("BtnDataFilter.Image")));
            this.BtnDataFilter.Label = "Фильтр";
            this.BtnDataFilter.Name = "BtnDataFilter";
            this.BtnDataFilter.ShowImage = true;
            this.BtnDataFilter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnDataFilter_Click);
            // 
            // SptBtnGroup
            // 
            this.SptBtnGroup.Image = ((System.Drawing.Image)(resources.GetObject("SptBtnGroup.Image")));
            this.SptBtnGroup.Items.Add(this.BtnGroupColumns);
            this.SptBtnGroup.Items.Add(this.BtnUngroupColumns);
            this.SptBtnGroup.Items.Add(this.BtnGroupRows);
            this.SptBtnGroup.Items.Add(this.BtnUngroupRows);
            this.SptBtnGroup.Label = "Группировка";
            this.SptBtnGroup.Name = "SptBtnGroup";
            this.SptBtnGroup.SuperTip = "Группирует столбцы и строки на листе Анализ";
            this.SptBtnGroup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGroupColumns_Click);
            // 
            // BtnGroupColumns
            // 
            this.BtnGroupColumns.Image = ((System.Drawing.Image)(resources.GetObject("BtnGroupColumns.Image")));
            this.BtnGroupColumns.Label = "Группировать столбцы";
            this.BtnGroupColumns.Name = "BtnGroupColumns";
            this.BtnGroupColumns.ShowImage = true;
            this.BtnGroupColumns.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGroupColumns_Click);
            // 
            // BtnUngroupColumns
            // 
            this.BtnUngroupColumns.Image = ((System.Drawing.Image)(resources.GetObject("BtnUngroupColumns.Image")));
            this.BtnUngroupColumns.Label = "Разгруппировать столбцы";
            this.BtnUngroupColumns.Name = "BtnUngroupColumns";
            this.BtnUngroupColumns.ShowImage = true;
            this.BtnUngroupColumns.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUngroupColumns_Click);
            // 
            // BtnGroupRows
            // 
            this.BtnGroupRows.Image = ((System.Drawing.Image)(resources.GetObject("BtnGroupRows.Image")));
            this.BtnGroupRows.Label = "Группировать строки";
            this.BtnGroupRows.Name = "BtnGroupRows";
            this.BtnGroupRows.ShowImage = true;
            this.BtnGroupRows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnGroupRows_Click);
            // 
            // BtnUngroupRows
            // 
            this.BtnUngroupRows.Image = ((System.Drawing.Image)(resources.GetObject("BtnUngroupRows.Image")));
            this.BtnUngroupRows.Label = "Разгруппировать строки";
            this.BtnUngroupRows.Name = "BtnUngroupRows";
            this.BtnUngroupRows.ShowImage = true;
            this.BtnUngroupRows.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnUngroupRows_Click);
            // 
            // SptBtnFormatComments
            // 
            this.SptBtnFormatComments.Image = ((System.Drawing.Image)(resources.GetObject("SptBtnFormatComments.Image")));
            this.SptBtnFormatComments.Items.Add(this.BtnSetFormul);
            this.SptBtnFormatComments.Items.Add(this.BtnFormatComments);
            this.SptBtnFormatComments.Items.Add(this.BtnClearFormateContions);
            this.SptBtnFormatComments.Label = "Анализ";
            this.SptBtnFormatComments.Name = "SptBtnFormatComments";
            this.SptBtnFormatComments.ScreenTip = "Форматирование комментариев";
            this.SptBtnFormatComments.SuperTip = "Добавлят правила условного форматирования  на листе Анализ";
            this.SptBtnFormatComments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.UpdateFormateAnalysis_Click);
            // 
            // BtnSetFormul
            // 
            this.BtnSetFormul.Image = ((System.Drawing.Image)(resources.GetObject("BtnSetFormul.Image")));
            this.BtnSetFormul.Label = "Настройки анализа";
            this.BtnSetFormul.Name = "BtnSetFormul";
            this.BtnSetFormul.ShowImage = true;
            this.BtnSetFormul.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSetFormul_Click);
            // 
            // BtnFormatComments
            // 
            this.BtnFormatComments.Image = ((System.Drawing.Image)(resources.GetObject("BtnFormatComments.Image")));
            this.BtnFormatComments.Label = "Настройки формата";
            this.BtnFormatComments.Name = "BtnFormatComments";
            this.BtnFormatComments.ShowImage = true;
            this.BtnFormatComments.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnFormatComments_Click);
            // 
            // BtnClearFormateContions
            // 
            this.BtnClearFormateContions.Image = ((System.Drawing.Image)(resources.GetObject("BtnClearFormateContions.Image")));
            this.BtnClearFormateContions.Label = "Очистить форматирование";
            this.BtnClearFormateContions.Name = "BtnClearFormateContions";
            this.BtnClearFormateContions.ShowImage = true;
            this.BtnClearFormateContions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnClearFormateContions_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.BtnProjectManager);
            this.group5.Items.Add(this.BtnKP);
            this.group5.Items.Add(this.BtnSettings);
            this.group5.Items.Add(this.BtnExcelScreenUpdating);
            this.group5.Items.Add(this.comboBoxLvlCost);
            this.group5.Label = "Настройки";
            this.group5.Name = "group5";
            // 
            // BtnProjectManager
            // 
            this.BtnProjectManager.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnProjectManager.Image = ((System.Drawing.Image)(resources.GetObject("BtnProjectManager.Image")));
            this.BtnProjectManager.Label = "Диспетчер проекта";
            this.BtnProjectManager.Name = "BtnProjectManager";
            this.BtnProjectManager.ScreenTip = "Настройки столбцов проектов";
            this.BtnProjectManager.ShowImage = true;
            this.BtnProjectManager.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ProjectManager_Click);
            // 
            // BtnKP
            // 
            this.BtnKP.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BtnKP.Image = ((System.Drawing.Image)(resources.GetObject("BtnKP.Image")));
            this.BtnKP.Label = "Диспетчер КП";
            this.BtnKP.Name = "BtnKP";
            this.BtnKP.ScreenTip = "Настройки столбцов в файлах КП.";
            this.BtnKP.ShowImage = true;
            this.BtnKP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ManagerKP_Click);
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
            this.BtnExcelScreenUpdating.Image = ((System.Drawing.Image)(resources.GetObject("BtnExcelScreenUpdating.Image")));
            this.BtnExcelScreenUpdating.Label = "Разблокировать";
            this.BtnExcelScreenUpdating.Name = "BtnExcelScreenUpdating";
            this.BtnExcelScreenUpdating.ScreenTip = "Разблокировать обновление окна Excel";
            this.BtnExcelScreenUpdating.ShowImage = true;
            this.BtnExcelScreenUpdating.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnExcelScreenUpdating_Click);
            // 
            // comboBoxLvlCost
            // 
            ribbonDropDownItemImpl1.Label = "Без НДС";
            ribbonDropDownItemImpl2.Label = "С НДС";
            this.comboBoxLvlCost.Items.Add(ribbonDropDownItemImpl1);
            this.comboBoxLvlCost.Items.Add(ribbonDropDownItemImpl2);
            this.comboBoxLvlCost.Label = "Уровень цен";
            this.comboBoxLvlCost.Name = "comboBoxLvlCost";
            this.comboBoxLvlCost.Text = null;
            this.comboBoxLvlCost.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CostLvl_TextChanged);
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
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnCreateProgect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadLvl12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnLoadLvl11;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnKP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSpectrum;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnExcelScreenUpdating;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton SptBtnFormatComments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFormatComments;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnClearFormateContions;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton SptBtnGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnGroupColumns;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnGroupRows;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUngroupColumns;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnUngroupRows;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton SptBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton SptBtnUpdateFormate;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnFormatNumber;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnSetFormul;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BtnDataFilter;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox comboBoxLvlCost;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
