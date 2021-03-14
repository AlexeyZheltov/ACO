namespace ACO.ProjectManager
{
    partial class FormManager
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.BtnAccept = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.TbInfo = new System.Windows.Forms.TextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label6 = new System.Windows.Forms.Label();
            this.TBoxSheetName = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.BtnSetSelectedRangeValues = new System.Windows.Forms.Button();
            this.TBoxFirstColumnRangeValues = new System.Windows.Forms.TextBox();
            this.TBoxLastColumnRangeValues = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.TBoxFirstRowRangeValues = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.BtnRangeOffer = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.TBoxLastColumnOffer = new System.Windows.Forms.TextBox();
            this.TBoxFirstColumnOffer = new System.Windows.Forms.TextBox();
            this.PageColumns = new System.Windows.Forms.TabPage();
            this.BtnCheckCells = new System.Windows.Forms.Button();
            this.BtnUpdateColumns = new System.Windows.Forms.Button();
            this.BtnDeleteColumnMapping = new System.Windows.Forms.GroupBox();
            this.ChkBoxObligatory = new System.Windows.Forms.CheckBox();
            this.ChkBoxCheck = new System.Windows.Forms.CheckBox();
            this.label8 = new System.Windows.Forms.Label();
            this.TextBoxValue = new System.Windows.Forms.TextBox();
            this.BtnDel = new System.Windows.Forms.Button();
            this.BtnAdd = new System.Windows.Forms.Button();
            this.BtnActiveCell = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.TextBoxAddres = new System.Windows.Forms.TextBox();
            this.PageProject = new System.Windows.Forms.TabPage();
            this.BtnOpenFolserSettings = new System.Windows.Forms.Button();
            this.BtnDelete = new System.Windows.Forms.Button();
            this.BtnSelect = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.TboxProjectName = new System.Windows.Forms.TextBox();
            this.BtnAddProject = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.TableProjects = new ACO.ProjectManager.CustomDataGrid();
            this.TableColumns = new ACO.ProjectManager.CustomDataGrid();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.PageColumns.SuspendLayout();
            this.BtnDeleteColumnMapping.SuspendLayout();
            this.PageProject.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnAccept
            // 
            this.BtnAccept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAccept.Location = new System.Drawing.Point(375, 509);
            this.BtnAccept.Name = "BtnAccept";
            this.BtnAccept.Size = new System.Drawing.Size(106, 23);
            this.BtnAccept.TabIndex = 1;
            this.BtnAccept.Text = "Сохранить";
            this.BtnAccept.UseVisualStyleBackColor = true;
            this.BtnAccept.Click += new System.EventHandler(this.BtnAccept_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(487, 509);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(106, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(598, 460);
            this.tabPage3.TabIndex = 5;
            this.tabPage3.Text = "Листы";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.TbInfo);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(598, 460);
            this.tabPage1.TabIndex = 4;
            this.tabPage1.Text = "Конфигурация";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // TbInfo
            // 
            this.TbInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TbInfo.Location = new System.Drawing.Point(4, 41);
            this.TbInfo.Multiline = true;
            this.TbInfo.Name = "TbInfo";
            this.TbInfo.Size = new System.Drawing.Size(589, 413);
            this.TbInfo.TabIndex = 7;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label6);
            this.tabPage2.Controls.Add(this.TBoxSheetName);
            this.tabPage2.Controls.Add(this.groupBox3);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(598, 460);
            this.tabPage2.TabIndex = 3;
            this.tabPage2.Text = "Диапазоны";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(46, 20);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "Лист Анализ";
            // 
            // TBoxSheetName
            // 
            this.TBoxSheetName.Location = new System.Drawing.Point(151, 19);
            this.TBoxSheetName.Name = "TBoxSheetName";
            this.TBoxSheetName.Size = new System.Drawing.Size(123, 20);
            this.TBoxSheetName.TabIndex = 11;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.BtnSetSelectedRangeValues);
            this.groupBox3.Controls.Add(this.TBoxFirstColumnRangeValues);
            this.groupBox3.Controls.Add(this.TBoxLastColumnRangeValues);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.TBoxFirstRowRangeValues);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Location = new System.Drawing.Point(20, 52);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(270, 158);
            this.groupBox3.TabIndex = 8;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Диапазон сумм";
            // 
            // BtnSetSelectedRangeValues
            // 
            this.BtnSetSelectedRangeValues.Location = new System.Drawing.Point(102, 118);
            this.BtnSetSelectedRangeValues.Name = "BtnSetSelectedRangeValues";
            this.BtnSetSelectedRangeValues.Size = new System.Drawing.Size(144, 29);
            this.BtnSetSelectedRangeValues.TabIndex = 6;
            this.BtnSetSelectedRangeValues.Text = "Выделенный диапазон";
            this.BtnSetSelectedRangeValues.UseVisualStyleBackColor = true;
            this.BtnSetSelectedRangeValues.Click += new System.EventHandler(this.BtnSetSelectedRangeValues_Click);
            // 
            // TBoxFirstColumnRangeValues
            // 
            this.TBoxFirstColumnRangeValues.Location = new System.Drawing.Point(166, 28);
            this.TBoxFirstColumnRangeValues.Name = "TBoxFirstColumnRangeValues";
            this.TBoxFirstColumnRangeValues.Size = new System.Drawing.Size(80, 20);
            this.TBoxFirstColumnRangeValues.TabIndex = 1;
            // 
            // TBoxLastColumnRangeValues
            // 
            this.TBoxLastColumnRangeValues.Location = new System.Drawing.Point(166, 54);
            this.TBoxLastColumnRangeValues.Name = "TBoxLastColumnRangeValues";
            this.TBoxLastColumnRangeValues.Size = new System.Drawing.Size(80, 20);
            this.TBoxLastColumnRangeValues.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(26, 87);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(124, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Строка начала данных ";
            // 
            // TBoxFirstRowRangeValues
            // 
            this.TBoxFirstRowRangeValues.Location = new System.Drawing.Point(166, 80);
            this.TBoxFirstRowRangeValues.Name = "TBoxFirstRowRangeValues";
            this.TBoxFirstRowRangeValues.Size = new System.Drawing.Size(80, 20);
            this.TBoxFirstRowRangeValues.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Последний столбец ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 34);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Первый столбец";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.BtnRangeOffer);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.TBoxLastColumnOffer);
            this.groupBox1.Controls.Add(this.TBoxFirstColumnOffer);
            this.groupBox1.Location = new System.Drawing.Point(20, 224);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(270, 145);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Диапазон КП";
            // 
            // BtnRangeOffer
            // 
            this.BtnRangeOffer.Location = new System.Drawing.Point(102, 104);
            this.BtnRangeOffer.Name = "BtnRangeOffer";
            this.BtnRangeOffer.Size = new System.Drawing.Size(144, 29);
            this.BtnRangeOffer.TabIndex = 18;
            this.BtnRangeOffer.Text = "Выделенный диапазон";
            this.BtnRangeOffer.UseVisualStyleBackColor = true;
            this.BtnRangeOffer.Click += new System.EventHandler(this.BtnRangeOffer_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(26, 69);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(110, 13);
            this.label10.TabIndex = 16;
            this.label10.Text = "Последний столбец ";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(26, 42);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(91, 13);
            this.label11.TabIndex = 17;
            this.label11.Text = "Первый столбец";
            // 
            // TBoxLastColumnOffer
            // 
            this.TBoxLastColumnOffer.Location = new System.Drawing.Point(166, 66);
            this.TBoxLastColumnOffer.Name = "TBoxLastColumnOffer";
            this.TBoxLastColumnOffer.Size = new System.Drawing.Size(80, 20);
            this.TBoxLastColumnOffer.TabIndex = 14;
            // 
            // TBoxFirstColumnOffer
            // 
            this.TBoxFirstColumnOffer.Location = new System.Drawing.Point(166, 40);
            this.TBoxFirstColumnOffer.Name = "TBoxFirstColumnOffer";
            this.TBoxFirstColumnOffer.Size = new System.Drawing.Size(80, 20);
            this.TBoxFirstColumnOffer.TabIndex = 15;
            // 
            // PageColumns
            // 
            this.PageColumns.Controls.Add(this.BtnCheckCells);
            this.PageColumns.Controls.Add(this.BtnUpdateColumns);
            this.PageColumns.Controls.Add(this.BtnDeleteColumnMapping);
            this.PageColumns.Controls.Add(this.TableColumns);
            this.PageColumns.Location = new System.Drawing.Point(4, 22);
            this.PageColumns.Name = "PageColumns";
            this.PageColumns.Padding = new System.Windows.Forms.Padding(3);
            this.PageColumns.Size = new System.Drawing.Size(598, 460);
            this.PageColumns.TabIndex = 0;
            this.PageColumns.Text = "Столбцы";
            this.PageColumns.UseVisualStyleBackColor = true;
            this.PageColumns.Click += new System.EventHandler(this.PageColumns_Click);
            // 
            // BtnCheckCells
            // 
            this.BtnCheckCells.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCheckCells.Location = new System.Drawing.Point(478, 10);
            this.BtnCheckCells.Name = "BtnCheckCells";
            this.BtnCheckCells.Size = new System.Drawing.Size(114, 29);
            this.BtnCheckCells.TabIndex = 7;
            this.BtnCheckCells.Text = "Проверка";
            this.BtnCheckCells.UseVisualStyleBackColor = true;
            this.BtnCheckCells.Click += new System.EventHandler(this.BtnCheckCells_Click_1);
            // 
            // BtnUpdateColumns
            // 
            this.BtnUpdateColumns.Location = new System.Drawing.Point(10, 10);
            this.BtnUpdateColumns.Name = "BtnUpdateColumns";
            this.BtnUpdateColumns.Size = new System.Drawing.Size(192, 29);
            this.BtnUpdateColumns.TabIndex = 5;
            this.BtnUpdateColumns.Text = "Добавить выделенный диапазон";
            this.BtnUpdateColumns.UseVisualStyleBackColor = true;
            this.BtnUpdateColumns.Click += new System.EventHandler(this.BtnUpdateColumns_Click);
            // 
            // BtnDeleteColumnMapping
            // 
            this.BtnDeleteColumnMapping.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDeleteColumnMapping.Controls.Add(this.ChkBoxObligatory);
            this.BtnDeleteColumnMapping.Controls.Add(this.ChkBoxCheck);
            this.BtnDeleteColumnMapping.Controls.Add(this.label8);
            this.BtnDeleteColumnMapping.Controls.Add(this.TextBoxValue);
            this.BtnDeleteColumnMapping.Controls.Add(this.BtnDel);
            this.BtnDeleteColumnMapping.Controls.Add(this.BtnAdd);
            this.BtnDeleteColumnMapping.Controls.Add(this.BtnActiveCell);
            this.BtnDeleteColumnMapping.Controls.Add(this.label7);
            this.BtnDeleteColumnMapping.Controls.Add(this.TextBoxAddres);
            this.BtnDeleteColumnMapping.Location = new System.Drawing.Point(11, 45);
            this.BtnDeleteColumnMapping.Name = "BtnDeleteColumnMapping";
            this.BtnDeleteColumnMapping.Size = new System.Drawing.Size(581, 77);
            this.BtnDeleteColumnMapping.TabIndex = 1;
            this.BtnDeleteColumnMapping.TabStop = false;
            this.BtnDeleteColumnMapping.Enter += new System.EventHandler(this.BtnDeleteColumnMapping_Enter);
            // 
            // ChkBoxObligatory
            // 
            this.ChkBoxObligatory.AutoSize = true;
            this.ChkBoxObligatory.Location = new System.Drawing.Point(113, 49);
            this.ChkBoxObligatory.Name = "ChkBoxObligatory";
            this.ChkBoxObligatory.Size = new System.Drawing.Size(101, 17);
            this.ChkBoxObligatory.TabIndex = 6;
            this.ChkBoxObligatory.Text = "Обязательный";
            this.ChkBoxObligatory.UseVisualStyleBackColor = true;
            // 
            // ChkBoxCheck
            // 
            this.ChkBoxCheck.AutoSize = true;
            this.ChkBoxCheck.Location = new System.Drawing.Point(9, 49);
            this.ChkBoxCheck.Name = "ChkBoxCheck";
            this.ChkBoxCheck.Size = new System.Drawing.Size(81, 17);
            this.ChkBoxCheck.TabIndex = 6;
            this.ChkBoxCheck.Text = "Проверять";
            this.ChkBoxCheck.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 23);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(55, 13);
            this.label8.TabIndex = 4;
            this.label8.Text = "Значение";
            // 
            // TextBoxValue
            // 
            this.TextBoxValue.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxValue.Location = new System.Drawing.Point(67, 19);
            this.TextBoxValue.Name = "TextBoxValue";
            this.TextBoxValue.Size = new System.Drawing.Size(290, 20);
            this.TextBoxValue.TabIndex = 3;
            // 
            // BtnDel
            // 
            this.BtnDel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDel.Location = new System.Drawing.Point(369, 43);
            this.BtnDel.Name = "BtnDel";
            this.BtnDel.Size = new System.Drawing.Size(92, 24);
            this.BtnDel.TabIndex = 2;
            this.BtnDel.Text = "Удалить";
            this.BtnDel.UseVisualStyleBackColor = true;
            this.BtnDel.Click += new System.EventHandler(this.BtnDel_Click);
            // 
            // BtnAdd
            // 
            this.BtnAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAdd.Location = new System.Drawing.Point(369, 18);
            this.BtnAdd.Name = "BtnAdd";
            this.BtnAdd.Size = new System.Drawing.Size(92, 24);
            this.BtnAdd.TabIndex = 2;
            this.BtnAdd.Text = "Добавить";
            this.BtnAdd.UseVisualStyleBackColor = true;
            this.BtnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // BtnActiveCell
            // 
            this.BtnActiveCell.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnActiveCell.Location = new System.Drawing.Point(467, 19);
            this.BtnActiveCell.Name = "BtnActiveCell";
            this.BtnActiveCell.Size = new System.Drawing.Size(104, 47);
            this.BtnActiveCell.TabIndex = 2;
            this.BtnActiveCell.Text = "Определить ячейку";
            this.BtnActiveCell.UseVisualStyleBackColor = true;
            this.BtnActiveCell.Click += new System.EventHandler(this.BtnActiveCell_Click);
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(237, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(38, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "Адрес";
            // 
            // TextBoxAddres
            // 
            this.TextBoxAddres.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxAddres.Location = new System.Drawing.Point(281, 46);
            this.TextBoxAddres.Name = "TextBoxAddres";
            this.TextBoxAddres.Size = new System.Drawing.Size(75, 20);
            this.TextBoxAddres.TabIndex = 0;
            // 
            // PageProject
            // 
            this.PageProject.Controls.Add(this.BtnOpenFolserSettings);
            this.PageProject.Controls.Add(this.BtnDelete);
            this.PageProject.Controls.Add(this.BtnSelect);
            this.PageProject.Controls.Add(this.groupBox2);
            this.PageProject.Controls.Add(this.label3);
            this.PageProject.Controls.Add(this.TableProjects);
            this.PageProject.Location = new System.Drawing.Point(4, 22);
            this.PageProject.Name = "PageProject";
            this.PageProject.Padding = new System.Windows.Forms.Padding(3);
            this.PageProject.Size = new System.Drawing.Size(597, 468);
            this.PageProject.TabIndex = 1;
            this.PageProject.Text = "Проект";
            this.PageProject.UseVisualStyleBackColor = true;
            // 
            // BtnOpenFolserSettings
            // 
            this.BtnOpenFolserSettings.Location = new System.Drawing.Point(15, 77);
            this.BtnOpenFolserSettings.Name = "BtnOpenFolserSettings";
            this.BtnOpenFolserSettings.Size = new System.Drawing.Size(137, 24);
            this.BtnOpenFolserSettings.TabIndex = 5;
            this.BtnOpenFolserSettings.Text = "Открыть папку";
            this.BtnOpenFolserSettings.UseVisualStyleBackColor = true;
            this.BtnOpenFolserSettings.Click += new System.EventHandler(this.BtnOpenFolserSettings_Click);
            // 
            // BtnDelete
            // 
            this.BtnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDelete.Location = new System.Drawing.Point(490, 109);
            this.BtnDelete.Name = "BtnDelete";
            this.BtnDelete.Size = new System.Drawing.Size(71, 24);
            this.BtnDelete.TabIndex = 4;
            this.BtnDelete.Text = "Удалить";
            this.BtnDelete.UseVisualStyleBackColor = true;
            this.BtnDelete.Click += new System.EventHandler(this.BtnDelete_Click);
            // 
            // BtnSelect
            // 
            this.BtnSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSelect.Location = new System.Drawing.Point(409, 109);
            this.BtnSelect.Name = "BtnSelect";
            this.BtnSelect.Size = new System.Drawing.Size(71, 24);
            this.BtnSelect.TabIndex = 4;
            this.BtnSelect.Text = "Выбрать";
            this.BtnSelect.UseVisualStyleBackColor = true;
            this.BtnSelect.Click += new System.EventHandler(this.BtnSelect_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.TboxProjectName);
            this.groupBox2.Controls.Add(this.BtnAddProject);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Location = new System.Drawing.Point(15, 11);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(546, 60);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            // 
            // TboxProjectName
            // 
            this.TboxProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TboxProjectName.Location = new System.Drawing.Point(129, 22);
            this.TboxProjectName.Name = "TboxProjectName";
            this.TboxProjectName.Size = new System.Drawing.Size(317, 20);
            this.TboxProjectName.TabIndex = 3;
            // 
            // BtnAddProject
            // 
            this.BtnAddProject.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAddProject.Location = new System.Drawing.Point(461, 17);
            this.BtnAddProject.Name = "BtnAddProject";
            this.BtnAddProject.Size = new System.Drawing.Size(71, 30);
            this.BtnAddProject.TabIndex = 2;
            this.BtnAddProject.Text = "Добавить";
            this.BtnAddProject.UseVisualStyleBackColor = true;
            this.BtnAddProject.Click += new System.EventHandler(this.BtnAddProject_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 26);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(114, 13);
            this.label4.TabIndex = 1;
            this.label4.Text = "Новая конфигурация";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 120);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(124, 13);
            this.label3.TabIndex = 1;
            this.label3.Text = "Конфигурация проекта";
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.PageProject);
            this.tabControl1.Controls.Add(this.PageColumns);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(3, 15);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(605, 494);
            this.tabControl1.TabIndex = 0;
            // 
            // TableProjects
            // 
            this.TableProjects.AllowUserToAddRows = false;
            this.TableProjects.AllowUserToResizeRows = false;
            this.TableProjects.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableProjects.BackgroundColor = System.Drawing.Color.White;
            this.TableProjects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TableProjects.Location = new System.Drawing.Point(6, 139);
            this.TableProjects.MultiSelect = false;
            this.TableProjects.Name = "TableProjects";
            this.TableProjects.ReadOnly = true;
            this.TableProjects.RowHeadersVisible = false;
            this.TableProjects.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.TableProjects.RowsDefaultCellStyle = dataGridViewCellStyle2;
            this.TableProjects.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableProjects.Size = new System.Drawing.Size(566, 326);
            this.TableProjects.TabIndex = 0;
            this.TableProjects.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableProjects_CellContentClick);
            this.TableProjects.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableProjects_CellValueChanged);
            // 
            // TableColumns
            // 
            this.TableColumns.AllowUserToAddRows = false;
            this.TableColumns.AllowUserToResizeColumns = false;
            this.TableColumns.AllowUserToResizeRows = false;
            this.TableColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableColumns.BackgroundColor = System.Drawing.Color.White;
            this.TableColumns.Location = new System.Drawing.Point(7, 137);
            this.TableColumns.MultiSelect = false;
            this.TableColumns.Name = "TableColumns";
            this.TableColumns.ReadOnly = true;
            this.TableColumns.RowHeadersVisible = false;
            this.TableColumns.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.TableColumns.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.TableColumns.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TableColumns.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableColumns.Size = new System.Drawing.Size(585, 315);
            this.TableColumns.TabIndex = 0;
            this.TableColumns.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableColumns_CellContentClick);
            this.TableColumns.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableColumns_CellValueChanged);
            this.TableColumns.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.TableColumns_RowsRemoved);
            this.TableColumns.SelectionChanged += new System.EventHandler(this.TableColumns_SelectionChanged);
            // 
            // FormManager
            // 
            this.AcceptButton = this.BtnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(611, 534);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnAccept);
            this.Controls.Add(this.tabControl1);
            this.Name = "FormManager";
            this.ShowIcon = false;
            this.Text = "Диспетчер";
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.PageColumns.ResumeLayout(false);
            this.BtnDeleteColumnMapping.ResumeLayout(false);
            this.BtnDeleteColumnMapping.PerformLayout();
            this.PageProject.ResumeLayout(false);
            this.PageProject.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button BtnAccept;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TextBox TbInfo;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button BtnSetSelectedRangeValues;
        private System.Windows.Forms.TextBox TBoxFirstColumnRangeValues;
        private System.Windows.Forms.TextBox TBoxLastColumnRangeValues;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox TBoxFirstRowRangeValues;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button BtnRangeOffer;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox TBoxLastColumnOffer;
        private System.Windows.Forms.TextBox TBoxFirstColumnOffer;
        private System.Windows.Forms.TabPage PageColumns;
        private System.Windows.Forms.Button BtnCheckCells;
        private System.Windows.Forms.Button BtnUpdateColumns;
        private System.Windows.Forms.GroupBox BtnDeleteColumnMapping;
        private System.Windows.Forms.CheckBox ChkBoxObligatory;
        private System.Windows.Forms.CheckBox ChkBoxCheck;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox TextBoxValue;
        private System.Windows.Forms.Button BtnDel;
        private System.Windows.Forms.Button BtnAdd;
        private System.Windows.Forms.Button BtnActiveCell;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox TextBoxAddres;
        private CustomDataGrid TableColumns;
        private System.Windows.Forms.TabPage PageProject;
        private System.Windows.Forms.Button BtnOpenFolserSettings;
        private System.Windows.Forms.Button BtnDelete;
        private System.Windows.Forms.Button BtnSelect;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox TboxProjectName;
        private System.Windows.Forms.Button BtnAddProject;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private CustomDataGrid TableProjects;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox TBoxSheetName;
    }
}