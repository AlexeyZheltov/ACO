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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.BtnAccept = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.BtnSetSelectedRangeValues = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.TBoxFirstRowRangeValues = new System.Windows.Forms.TextBox();
            this.TBoxLastColumnRangeValues = new System.Windows.Forms.TextBox();
            this.TBoxFirstColumnRangeValues = new System.Windows.Forms.TextBox();
            this.PageColumns = new System.Windows.Forms.TabPage();
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
            this.TableColumns = new ACO.ProjectManager.CustomDataGrid();
            this.PageProject = new System.Windows.Forms.TabPage();
            this.BtnOpenFolserSettings = new System.Windows.Forms.Button();
            this.BtnDelete = new System.Windows.Forms.Button();
            this.BtnSelect = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.TboxProjectName = new System.Windows.Forms.TextBox();
            this.BtnAddProject = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.TableProjects = new ACO.ProjectManager.CustomDataGrid();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.TbInfo = new System.Windows.Forms.TextBox();
            this.BtnCheckCells = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.label6 = new System.Windows.Forms.Label();
            this.TBoxSheetName = new System.Windows.Forms.TextBox();
            this.tabPage2.SuspendLayout();
            this.PageColumns.SuspendLayout();
            this.BtnDeleteColumnMapping.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).BeginInit();
            this.PageProject.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnAccept
            // 
            this.BtnAccept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAccept.Location = new System.Drawing.Point(385, 501);
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
            this.BtnCancel.Location = new System.Drawing.Point(497, 501);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(106, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.BtnSetSelectedRangeValues);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.TBoxFirstRowRangeValues);
            this.tabPage2.Controls.Add(this.TBoxLastColumnRangeValues);
            this.tabPage2.Controls.Add(this.TBoxFirstColumnRangeValues);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(588, 460);
            this.tabPage2.TabIndex = 3;
            this.tabPage2.Text = "Диапазон сумм";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // BtnSetSelectedRangeValues
            // 
            this.BtnSetSelectedRangeValues.Location = new System.Drawing.Point(135, 169);
            this.BtnSetSelectedRangeValues.Name = "BtnSetSelectedRangeValues";
            this.BtnSetSelectedRangeValues.Size = new System.Drawing.Size(144, 29);
            this.BtnSetSelectedRangeValues.TabIndex = 6;
            this.BtnSetSelectedRangeValues.Text = "Выделенный диапазон";
            this.BtnSetSelectedRangeValues.UseVisualStyleBackColor = true;
            this.BtnSetSelectedRangeValues.Click += new System.EventHandler(this.BtnSetSelectedRangeValues_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 134);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(124, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Строка начала данных ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 108);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Последний столбец ";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 81);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(91, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Первый столбец";
            // 
            // TBoxFirstRowRangeValues
            // 
            this.TBoxFirstRowRangeValues.Location = new System.Drawing.Point(164, 127);
            this.TBoxFirstRowRangeValues.Name = "TBoxFirstRowRangeValues";
            this.TBoxFirstRowRangeValues.Size = new System.Drawing.Size(100, 20);
            this.TBoxFirstRowRangeValues.TabIndex = 1;
            // 
            // TBoxLastColumnRangeValues
            // 
            this.TBoxLastColumnRangeValues.Location = new System.Drawing.Point(164, 101);
            this.TBoxLastColumnRangeValues.Name = "TBoxLastColumnRangeValues";
            this.TBoxLastColumnRangeValues.Size = new System.Drawing.Size(100, 20);
            this.TBoxLastColumnRangeValues.TabIndex = 1;
            // 
            // TBoxFirstColumnRangeValues
            // 
            this.TBoxFirstColumnRangeValues.Location = new System.Drawing.Point(164, 75);
            this.TBoxFirstColumnRangeValues.Name = "TBoxFirstColumnRangeValues";
            this.TBoxFirstColumnRangeValues.Size = new System.Drawing.Size(100, 20);
            this.TBoxFirstColumnRangeValues.TabIndex = 1;
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
            this.PageColumns.Size = new System.Drawing.Size(588, 460);
            this.PageColumns.TabIndex = 0;
            this.PageColumns.Text = "Столбцы";
            this.PageColumns.UseVisualStyleBackColor = true;
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
            this.BtnDeleteColumnMapping.Size = new System.Drawing.Size(569, 77);
            this.BtnDeleteColumnMapping.TabIndex = 1;
            this.BtnDeleteColumnMapping.TabStop = false;
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
            this.TextBoxValue.Size = new System.Drawing.Size(278, 20);
            this.TextBoxValue.TabIndex = 3;
            // 
            // BtnDel
            // 
            this.BtnDel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDel.Location = new System.Drawing.Point(357, 43);
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
            this.BtnAdd.Location = new System.Drawing.Point(357, 18);
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
            this.BtnActiveCell.Location = new System.Drawing.Point(455, 19);
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
            this.label7.Location = new System.Drawing.Point(225, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(38, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "Адрес";
            // 
            // TextBoxAddres
            // 
            this.TextBoxAddres.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxAddres.Location = new System.Drawing.Point(269, 46);
            this.TextBoxAddres.Name = "TextBoxAddres";
            this.TextBoxAddres.Size = new System.Drawing.Size(75, 20);
            this.TextBoxAddres.TabIndex = 0;
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
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.TableColumns.RowsDefaultCellStyle = dataGridViewCellStyle3;
            this.TableColumns.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TableColumns.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableColumns.Size = new System.Drawing.Size(573, 315);
            this.TableColumns.TabIndex = 0;
            this.TableColumns.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableColumns_CellValueChanged);
            this.TableColumns.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.TableColumns_RowsRemoved);
            this.TableColumns.SelectionChanged += new System.EventHandler(this.TableColumns_SelectionChanged);
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
            this.PageProject.Size = new System.Drawing.Size(588, 460);
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
            this.BtnDelete.Location = new System.Drawing.Point(500, 109);
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
            this.BtnSelect.Location = new System.Drawing.Point(419, 109);
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
            this.groupBox2.Size = new System.Drawing.Size(556, 60);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            // 
            // TboxProjectName
            // 
            this.TboxProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TboxProjectName.Location = new System.Drawing.Point(129, 22);
            this.TboxProjectName.Name = "TboxProjectName";
            this.TboxProjectName.Size = new System.Drawing.Size(327, 20);
            this.TboxProjectName.TabIndex = 3;
            // 
            // BtnAddProject
            // 
            this.BtnAddProject.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAddProject.Location = new System.Drawing.Point(471, 17);
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
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.TableProjects.RowsDefaultCellStyle = dataGridViewCellStyle4;
            this.TableProjects.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableProjects.Size = new System.Drawing.Size(576, 318);
            this.TableProjects.TabIndex = 0;
            this.TableProjects.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableProjects_CellContentClick);
            this.TableProjects.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableProjects_CellValueChanged);
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
            this.tabControl1.Location = new System.Drawing.Point(13, 13);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(596, 486);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.TbInfo);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(588, 460);
            this.tabPage1.TabIndex = 4;
            this.tabPage1.Text = "Конфигурация";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // TbInfo
            // 
            this.TbInfo.Location = new System.Drawing.Point(3, 41);
            this.TbInfo.Multiline = true;
            this.TbInfo.Name = "TbInfo";
            this.TbInfo.Size = new System.Drawing.Size(579, 413);
            this.TbInfo.TabIndex = 7;
            // 
            // BtnCheckCells
            // 
            this.BtnCheckCells.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCheckCells.Location = new System.Drawing.Point(466, 8);
            this.BtnCheckCells.Name = "BtnCheckCells";
            this.BtnCheckCells.Size = new System.Drawing.Size(114, 29);
            this.BtnCheckCells.TabIndex = 7;
            this.BtnCheckCells.Text = "Проверка";
            this.BtnCheckCells.UseVisualStyleBackColor = true;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Controls.Add(this.TBoxSheetName);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(588, 460);
            this.tabPage3.TabIndex = 5;
            this.tabPage3.Text = "Листы";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(21, 21);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Лист Анализ";
            // 
            // TBoxSheetName
            // 
            this.TBoxSheetName.Location = new System.Drawing.Point(119, 18);
            this.TBoxSheetName.Name = "TBoxSheetName";
            this.TBoxSheetName.Size = new System.Drawing.Size(100, 20);
            this.TBoxSheetName.TabIndex = 9;
            // 
            // FormManager
            // 
            this.AcceptButton = this.BtnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(621, 526);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnAccept);
            this.Controls.Add(this.tabControl1);
            this.Name = "FormManager";
            this.ShowIcon = false;
            this.Text = "Диспетчер";
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.PageColumns.ResumeLayout(false);
            this.BtnDeleteColumnMapping.ResumeLayout(false);
            this.BtnDeleteColumnMapping.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).EndInit();
            this.PageProject.ResumeLayout(false);
            this.PageProject.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button BtnAccept;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage PageColumns;
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
        private System.Windows.Forms.Button BtnSetSelectedRangeValues;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox TBoxLastColumnRangeValues;
        private System.Windows.Forms.TextBox TBoxFirstColumnRangeValues;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox TBoxFirstRowRangeValues;
        private System.Windows.Forms.Button BtnCheckCells;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TextBox TbInfo;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox TBoxSheetName;
    }
}