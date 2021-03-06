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
            this.BtnAccept = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.PageProject.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).BeginInit();
            this.PageColumns.SuspendLayout();
            this.BtnDeleteColumnMapping.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.PageProject);
            this.tabControl1.Controls.Add(this.PageColumns);
            this.tabControl1.Location = new System.Drawing.Point(13, 13);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(638, 502);
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
            this.PageProject.Size = new System.Drawing.Size(630, 476);
            this.PageProject.TabIndex = 1;
            this.PageProject.Text = "Проект";
            this.PageProject.UseVisualStyleBackColor = true;
            // 
            // BtnDelete
            // 
            this.BtnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDelete.Location = new System.Drawing.Point(536, 396);
            this.BtnDelete.Name = "BtnDelete";
            this.BtnDelete.Size = new System.Drawing.Size(71, 24);
            this.BtnDelete.TabIndex = 4;
            this.BtnDelete.Text = "Удалить";
            this.BtnDelete.UseVisualStyleBackColor = true;
            // 
            // BtnSelect
            // 
            this.BtnSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSelect.Location = new System.Drawing.Point(455, 396);
            this.BtnSelect.Name = "BtnSelect";
            this.BtnSelect.Size = new System.Drawing.Size(71, 24);
            this.BtnSelect.TabIndex = 4;
            this.BtnSelect.Text = "Выбрать";
            this.BtnSelect.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.TboxProjectName);
            this.groupBox2.Controls.Add(this.BtnAddProject);
            this.groupBox2.Location = new System.Drawing.Point(15, 34);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(592, 60);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            // 
            // TboxProjectName
            // 
            this.TboxProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TboxProjectName.Location = new System.Drawing.Point(18, 25);
            this.TboxProjectName.Name = "TboxProjectName";
            this.TboxProjectName.Size = new System.Drawing.Size(471, 20);
            this.TboxProjectName.TabIndex = 3;
            // 
            // BtnAddProject
            // 
            this.BtnAddProject.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAddProject.Location = new System.Drawing.Point(507, 20);
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
            this.TableProjects.Size = new System.Drawing.Size(592, 272);
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
            this.PageColumns.Size = new System.Drawing.Size(630, 476);
            this.PageColumns.TabIndex = 0;
            this.PageColumns.Text = "Столбцы";
            this.PageColumns.UseVisualStyleBackColor = true;
            // 
            // BtnCheckCells
            // 
            this.BtnCheckCells.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCheckCells.Location = new System.Drawing.Point(398, 10);
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
            this.BtnUpdateColumns.Location = new System.Drawing.Point(508, 10);
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
            this.BtnDeleteColumnMapping.Size = new System.Drawing.Size(611, 77);
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
            this.TextBoxValue.Size = new System.Drawing.Size(320, 20);
            this.TextBoxValue.TabIndex = 3;
            // 
            // BtnDel
            // 
            this.BtnDel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDel.Location = new System.Drawing.Point(399, 43);
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
            this.BtnAdd.Location = new System.Drawing.Point(399, 18);
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
            this.BtnActiveCell.Location = new System.Drawing.Point(497, 19);
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
            this.label7.Location = new System.Drawing.Point(267, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(38, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "Адрес";
            // 
            // TextBoxAddres
            // 
            this.TextBoxAddres.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.TextBoxAddres.Location = new System.Drawing.Point(311, 46);
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
            this.TableColumns.Size = new System.Drawing.Size(615, 331);
            this.TableColumns.TabIndex = 0;
            this.TableColumns.SelectionChanged += new System.EventHandler(this.TableColumns_SelectionChanged);
            // 
            // BtnAccept
            // 
            this.BtnAccept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAccept.Location = new System.Drawing.Point(427, 517);
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
            this.BtnCancel.Location = new System.Drawing.Point(539, 517);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(106, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            // 
            // FormManager
            // 
            this.AcceptButton = this.BtnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(663, 542);
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
    }
}