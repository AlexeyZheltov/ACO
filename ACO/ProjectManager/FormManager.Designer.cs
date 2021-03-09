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
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.PageProject = new System.Windows.Forms.TabPage();
            this.BtnDelete = new System.Windows.Forms.Button();
            this.BtnSelect = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.TboxProjectName = new System.Windows.Forms.TextBox();
            this.BtnAddProject = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.TableProjects = new ACO.ProjectManager.CustomDataGrid();
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
            this.TableColumns = new ACO.ProjectManager.CustomDataGrid();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.customDataGrid1 = new ACO.ProjectManager.CustomDataGrid();
            this.BtnAccept = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.tabControl1.SuspendLayout();
            this.PageProject.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).BeginInit();
            this.PageColumns.SuspendLayout();
            this.BtnDeleteColumnMapping.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).BeginInit();
            this.tabPage1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.customDataGrid1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.PageProject);
            this.tabControl1.Controls.Add(this.PageColumns);
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(13, 13);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(648, 503);
            this.tabControl1.TabIndex = 0;
            // 
            // PageProject
            // 
            this.PageProject.Controls.Add(this.BtnDelete);
            this.PageProject.Controls.Add(this.BtnSelect);
            this.PageProject.Controls.Add(this.groupBox2);
            this.PageProject.Controls.Add(this.label3);
            this.PageProject.Controls.Add(this.TableProjects);
            this.PageProject.Location = new System.Drawing.Point(4, 22);
            this.PageProject.Name = "PageProject";
            this.PageProject.Padding = new System.Windows.Forms.Padding(3);
            this.PageProject.Size = new System.Drawing.Size(640, 477);
            this.PageProject.TabIndex = 1;
            this.PageProject.Text = "Проект";
            this.PageProject.UseVisualStyleBackColor = true;
            // 
            // BtnDelete
            // 
            this.BtnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDelete.Location = new System.Drawing.Point(546, 397);
            this.BtnDelete.Name = "BtnDelete";
            this.BtnDelete.Size = new System.Drawing.Size(71, 24);
            this.BtnDelete.TabIndex = 4;
            this.BtnDelete.Text = "Удалить";
            this.BtnDelete.UseVisualStyleBackColor = true;
            // 
            // BtnSelect
            // 
            this.BtnSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSelect.Location = new System.Drawing.Point(465, 397);
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
            this.groupBox2.Location = new System.Drawing.Point(15, 34);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(602, 60);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            // 
            // TboxProjectName
            // 
            this.TboxProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TboxProjectName.Location = new System.Drawing.Point(18, 25);
            this.TboxProjectName.Name = "TboxProjectName";
            this.TboxProjectName.Size = new System.Drawing.Size(481, 20);
            this.TboxProjectName.TabIndex = 3;
            // 
            // BtnAddProject
            // 
            this.BtnAddProject.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAddProject.Location = new System.Drawing.Point(517, 20);
            this.BtnAddProject.Name = "BtnAddProject";
            this.BtnAddProject.Size = new System.Drawing.Size(71, 30);
            this.BtnAddProject.TabIndex = 2;
            this.BtnAddProject.Text = "Добавить";
            this.BtnAddProject.UseVisualStyleBackColor = true;
            this.BtnAddProject.Click += new System.EventHandler(this.BtnAddProject_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 14);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(124, 13);
            this.label3.TabIndex = 1;
            this.label3.Text = "Конфигурация проекта";
            // 
            // TableProjects
            // 
            this.TableProjects.AllowUserToAddRows = false;
            this.TableProjects.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableProjects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TableProjects.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.TableProjects.Location = new System.Drawing.Point(15, 110);
            this.TableProjects.MultiSelect = false;
            this.TableProjects.Name = "TableProjects";
            this.TableProjects.RowHeadersVisible = false;
            this.TableProjects.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableProjects.Size = new System.Drawing.Size(602, 273);
            this.TableProjects.TabIndex = 0;
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
            this.PageColumns.Size = new System.Drawing.Size(640, 477);
            this.PageColumns.TabIndex = 0;
            this.PageColumns.Text = "Столбцы 1";
            this.PageColumns.UseVisualStyleBackColor = true;
            // 
            // BtnCheckCells
            // 
            this.BtnCheckCells.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCheckCells.Location = new System.Drawing.Point(408, 10);
            this.BtnCheckCells.Name = "BtnCheckCells";
            this.BtnCheckCells.Size = new System.Drawing.Size(104, 29);
            this.BtnCheckCells.TabIndex = 5;
            this.BtnCheckCells.Text = "Проверить";
            this.BtnCheckCells.UseVisualStyleBackColor = true;
            this.BtnCheckCells.Click += new System.EventHandler(this.BtnCheckCells_Click);
            // 
            // BtnUpdateColumns
            // 
            this.BtnUpdateColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnUpdateColumns.Location = new System.Drawing.Point(518, 10);
            this.BtnUpdateColumns.Name = "BtnUpdateColumns";
            this.BtnUpdateColumns.Size = new System.Drawing.Size(104, 29);
            this.BtnUpdateColumns.TabIndex = 5;
            this.BtnUpdateColumns.Text = "Обновить";
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
            this.BtnDeleteColumnMapping.Size = new System.Drawing.Size(621, 77);
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
            this.TextBoxValue.Location = new System.Drawing.Point(67, 23);
            this.TextBoxValue.Name = "TextBoxValue";
            this.TextBoxValue.Size = new System.Drawing.Size(330, 20);
            this.TextBoxValue.TabIndex = 3;
            // 
            // BtnDel
            // 
            this.BtnDel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDel.Location = new System.Drawing.Point(409, 43);
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
            this.BtnAdd.Location = new System.Drawing.Point(409, 18);
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
            this.BtnActiveCell.Location = new System.Drawing.Point(507, 19);
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
            this.label7.Location = new System.Drawing.Point(277, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(38, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "Адрес";
            // 
            // TextBoxAddres
            // 
            this.TextBoxAddres.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxAddres.Location = new System.Drawing.Point(321, 46);
            this.TextBoxAddres.Name = "TextBoxAddres";
            this.TextBoxAddres.Size = new System.Drawing.Size(75, 20);
            this.TextBoxAddres.TabIndex = 0;
            // 
            // TableColumns
            // 
            this.TableColumns.AllowUserToAddRows = false;
            this.TableColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableColumns.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.TableColumns.Location = new System.Drawing.Point(7, 137);
            this.TableColumns.MultiSelect = false;
            this.TableColumns.Name = "TableColumns";
            this.TableColumns.RowHeadersVisible = false;
            this.TableColumns.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableColumns.Size = new System.Drawing.Size(625, 332);
            this.TableColumns.TabIndex = 0;
            this.TableColumns.SelectionChanged += new System.EventHandler(this.TableColumns_SelectionChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.customDataGrid1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(640, 477);
            this.tabPage1.TabIndex = 2;
            this.tabPage1.Text = "Столбцы КП";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.checkBox2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Location = new System.Drawing.Point(12, 26);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(621, 77);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(113, 49);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(101, 17);
            this.checkBox1.TabIndex = 6;
            this.checkBox1.Text = "Обязательный";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(9, 49);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(81, 17);
            this.checkBox2.TabIndex = 6;
            this.checkBox2.Text = "Проверять";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Значение";
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(67, 23);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(330, 20);
            this.textBox1.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(409, 43);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(92, 24);
            this.button1.TabIndex = 2;
            this.button1.Text = "Удалить";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.Location = new System.Drawing.Point(409, 18);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(92, 24);
            this.button2.TabIndex = 2;
            this.button2.Text = "Добавить";
            this.button2.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(507, 19);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(104, 47);
            this.button3.TabIndex = 2;
            this.button3.Text = "Определить ячейку";
            this.button3.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(277, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Адрес";
            // 
            // textBox2
            // 
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox2.Location = new System.Drawing.Point(321, 46);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(75, 20);
            this.textBox2.TabIndex = 0;
            // 
            // customDataGrid1
            // 
            this.customDataGrid1.AllowUserToAddRows = false;
            this.customDataGrid1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.customDataGrid1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.customDataGrid1.Location = new System.Drawing.Point(8, 118);
            this.customDataGrid1.MultiSelect = false;
            this.customDataGrid1.Name = "customDataGrid1";
            this.customDataGrid1.RowHeadersVisible = false;
            this.customDataGrid1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.customDataGrid1.Size = new System.Drawing.Size(625, 332);
            this.customDataGrid1.TabIndex = 2;
            // 
            // BtnAccept
            // 
            this.BtnAccept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAccept.Location = new System.Drawing.Point(437, 518);
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
            this.BtnCancel.Location = new System.Drawing.Point(549, 518);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(106, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(640, 477);
            this.tabPage2.TabIndex = 3;
            this.tabPage2.Text = "tabPage2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // FormManager
            // 
            this.AcceptButton = this.BtnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(673, 543);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnAccept);
            this.Controls.Add(this.tabControl1);
            this.Name = "FormManager";
            this.Text = "Диспетчер";
            this.tabControl1.ResumeLayout(false);
            this.PageProject.ResumeLayout(false);
            this.PageProject.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).EndInit();
            this.PageColumns.ResumeLayout(false);
            this.BtnDeleteColumnMapping.ResumeLayout(false);
            this.BtnDeleteColumnMapping.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).EndInit();
            this.tabPage1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.customDataGrid1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage PageColumns;
        private System.Windows.Forms.TabPage PageProject;
        private System.Windows.Forms.Button BtnAccept;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.GroupBox BtnDeleteColumnMapping;
        private CustomDataGrid TableColumns;
        private System.Windows.Forms.Button BtnAddProject;
        private System.Windows.Forms.Label label3;
        private CustomDataGrid TableProjects;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox TboxProjectName;
        private System.Windows.Forms.Button BtnDelete;
        private System.Windows.Forms.Button BtnSelect;
        private System.Windows.Forms.Button BtnUpdateColumns;
        private System.Windows.Forms.Button BtnAdd;
        private System.Windows.Forms.Button BtnActiveCell;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox TextBoxAddres;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox TextBoxValue;
        private System.Windows.Forms.CheckBox ChkBoxObligatory;
        private System.Windows.Forms.CheckBox ChkBoxCheck;
        private System.Windows.Forms.Button BtnCheckCells;
        private System.Windows.Forms.Button BtnDel;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox2;
        private CustomDataGrid customDataGrid1;
        private System.Windows.Forms.TabPage tabPage2;
    }
}