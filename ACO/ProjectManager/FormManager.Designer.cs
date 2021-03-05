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
            this.BtnUpdateColumns = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.numericUpDown2 = new System.Windows.Forms.NumericUpDown();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label8 = new System.Windows.Forms.Label();
            this.TextBoxValue = new System.Windows.Forms.TextBox();
            this.BtnAdd = new System.Windows.Forms.Button();
            this.BtnActiveCell = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.TextBoxColumn = new System.Windows.Forms.TextBox();
            this.TextBoxRow = new System.Windows.Forms.TextBox();
            this.TextBoxAddres = new System.Windows.Forms.TextBox();
            this.TextBoxCellName = new System.Windows.Forms.TextBox();
            this.TableColumns = new ACO.ProjectManager.CustomDataGrid();
            this.BtnAccept = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.ChkBoxCheck = new System.Windows.Forms.CheckBox();
            this.ChkBoxObligatory = new System.Windows.Forms.CheckBox();
            this.BtnCheckCells = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.PageProject.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).BeginInit();
            this.PageColumns.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.groupBox1.SuspendLayout();
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
            this.tabControl1.Size = new System.Drawing.Size(741, 502);
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
            this.PageProject.Size = new System.Drawing.Size(624, 486);
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
            this.TableProjects.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableProjects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
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
            this.PageColumns.Controls.Add(this.label2);
            this.PageColumns.Controls.Add(this.label1);
            this.PageColumns.Controls.Add(this.numericUpDown2);
            this.PageColumns.Controls.Add(this.numericUpDown1);
            this.PageColumns.Controls.Add(this.groupBox1);
            this.PageColumns.Controls.Add(this.TableColumns);
            this.PageColumns.Location = new System.Drawing.Point(4, 22);
            this.PageColumns.Name = "PageColumns";
            this.PageColumns.Padding = new System.Windows.Forms.Padding(3);
            this.PageColumns.Size = new System.Drawing.Size(733, 476);
            this.PageColumns.TabIndex = 0;
            this.PageColumns.Text = "Столбцы";
            this.PageColumns.UseVisualStyleBackColor = true;
            // 
            // BtnUpdateColumns
            // 
            this.BtnUpdateColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnUpdateColumns.Location = new System.Drawing.Point(611, 10);
            this.BtnUpdateColumns.Name = "BtnUpdateColumns";
            this.BtnUpdateColumns.Size = new System.Drawing.Size(104, 34);
            this.BtnUpdateColumns.TabIndex = 5;
            this.BtnUpdateColumns.Text = "Обновить";
            this.BtnUpdateColumns.UseVisualStyleBackColor = true;
            this.BtnUpdateColumns.Click += new System.EventHandler(this.BtnUpdateColumns_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(223, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(19, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "до";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(127, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Строки заголовков. От:";
            // 
            // numericUpDown2
            // 
            this.numericUpDown2.Location = new System.Drawing.Point(248, 16);
            this.numericUpDown2.Name = "numericUpDown2";
            this.numericUpDown2.Size = new System.Drawing.Size(68, 20);
            this.numericUpDown2.TabIndex = 2;
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(144, 16);
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(68, 20);
            this.numericUpDown1.TabIndex = 2;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.ChkBoxObligatory);
            this.groupBox1.Controls.Add(this.ChkBoxCheck);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.TextBoxValue);
            this.groupBox1.Controls.Add(this.BtnAdd);
            this.groupBox1.Controls.Add(this.BtnActiveCell);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.TextBoxColumn);
            this.groupBox1.Controls.Add(this.TextBoxRow);
            this.groupBox1.Controls.Add(this.TextBoxAddres);
            this.groupBox1.Controls.Add(this.TextBoxCellName);
            this.groupBox1.Location = new System.Drawing.Point(11, 45);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(714, 109);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Добавить стобец";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 52);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(55, 13);
            this.label8.TabIndex = 4;
            this.label8.Text = "Значение";
            // 
            // TextBoxValue
            // 
            this.TextBoxValue.Location = new System.Drawing.Point(70, 48);
            this.TextBoxValue.Multiline = true;
            this.TextBoxValue.Name = "TextBoxValue";
            this.TextBoxValue.Size = new System.Drawing.Size(202, 42);
            this.TextBoxValue.TabIndex = 3;
            // 
            // BtnAdd
            // 
            this.BtnAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAdd.Location = new System.Drawing.Point(600, 65);
            this.BtnAdd.Name = "BtnAdd";
            this.BtnAdd.Size = new System.Drawing.Size(104, 34);
            this.BtnAdd.TabIndex = 2;
            this.BtnAdd.Text = "Добавить";
            this.BtnAdd.UseVisualStyleBackColor = true;
            this.BtnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // BtnActiveCell
            // 
            this.BtnActiveCell.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnActiveCell.Location = new System.Drawing.Point(600, 15);
            this.BtnActiveCell.Name = "BtnActiveCell";
            this.BtnActiveCell.Size = new System.Drawing.Size(104, 44);
            this.BtnActiveCell.TabIndex = 2;
            this.BtnActiveCell.Text = "Определить ячейку";
            this.BtnActiveCell.UseVisualStyleBackColor = true;
            this.BtnActiveCell.Click += new System.EventHandler(this.BtnActiveCell_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(300, 78);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(49, 13);
            this.label6.TabIndex = 1;
            this.label6.Text = "Столбец";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(302, 52);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(43, 13);
            this.label5.TabIndex = 1;
            this.label5.Text = "Строка";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(310, 27);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(38, 13);
            this.label7.TabIndex = 1;
            this.label7.Text = "Адрес";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 27);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 13);
            this.label4.TabIndex = 1;
            this.label4.Text = "Ячейка";
            // 
            // TextBoxColumn
            // 
            this.TextBoxColumn.Location = new System.Drawing.Point(354, 74);
            this.TextBoxColumn.Name = "TextBoxColumn";
            this.TextBoxColumn.Size = new System.Drawing.Size(75, 20);
            this.TextBoxColumn.TabIndex = 0;
            // 
            // TextBoxRow
            // 
            this.TextBoxRow.Location = new System.Drawing.Point(354, 48);
            this.TextBoxRow.Name = "TextBoxRow";
            this.TextBoxRow.Size = new System.Drawing.Size(75, 20);
            this.TextBoxRow.TabIndex = 0;
            // 
            // TextBoxAddres
            // 
            this.TextBoxAddres.Location = new System.Drawing.Point(354, 23);
            this.TextBoxAddres.Name = "TextBoxAddres";
            this.TextBoxAddres.Size = new System.Drawing.Size(75, 20);
            this.TextBoxAddres.TabIndex = 0;
            // 
            // TextBoxCellName
            // 
            this.TextBoxCellName.Location = new System.Drawing.Point(70, 23);
            this.TextBoxCellName.Name = "TextBoxCellName";
            this.TextBoxCellName.Size = new System.Drawing.Size(203, 20);
            this.TextBoxCellName.TabIndex = 0;
            // 
            // TableColumns
            // 
            this.TableColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableColumns.Location = new System.Drawing.Point(7, 160);
            this.TableColumns.MultiSelect = false;
            this.TableColumns.Name = "TableColumns";
            this.TableColumns.RowHeadersVisible = false;
            this.TableColumns.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableColumns.Size = new System.Drawing.Size(718, 308);
            this.TableColumns.TabIndex = 0;
            this.TableColumns.SelectionChanged += new System.EventHandler(this.TableColumns_SelectionChanged);
            // 
            // BtnAccept
            // 
            this.BtnAccept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAccept.Location = new System.Drawing.Point(530, 517);
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
            this.BtnCancel.Location = new System.Drawing.Point(642, 517);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(106, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            // 
            // ChkBoxCheck
            // 
            this.ChkBoxCheck.AutoSize = true;
            this.ChkBoxCheck.Location = new System.Drawing.Point(450, 23);
            this.ChkBoxCheck.Name = "ChkBoxCheck";
            this.ChkBoxCheck.Size = new System.Drawing.Size(81, 17);
            this.ChkBoxCheck.TabIndex = 6;
            this.ChkBoxCheck.Text = "Проверять";
            this.ChkBoxCheck.UseVisualStyleBackColor = true;
            // 
            // ChkBoxObligatory
            // 
            this.ChkBoxObligatory.AutoSize = true;
            this.ChkBoxObligatory.Location = new System.Drawing.Point(450, 50);
            this.ChkBoxObligatory.Name = "ChkBoxObligatory";
            this.ChkBoxObligatory.Size = new System.Drawing.Size(101, 17);
            this.ChkBoxObligatory.TabIndex = 6;
            this.ChkBoxObligatory.Text = "Обязательный";
            this.ChkBoxObligatory.UseVisualStyleBackColor = true;
            // 
            // BtnCheckCells
            // 
            this.BtnCheckCells.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCheckCells.Location = new System.Drawing.Point(461, 10);
            this.BtnCheckCells.Name = "BtnCheckCells";
            this.BtnCheckCells.Size = new System.Drawing.Size(104, 34);
            this.BtnCheckCells.TabIndex = 5;
            this.BtnCheckCells.Text = "Проверить";
            this.BtnCheckCells.UseVisualStyleBackColor = true;
            this.BtnCheckCells.Click += new System.EventHandler(this.BtnCheckCells_Click);
            // 
            // FormManager
            // 
            this.AcceptButton = this.BtnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(766, 542);
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
            this.PageColumns.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage PageColumns;
        private System.Windows.Forms.TabPage PageProject;
        private System.Windows.Forms.Button BtnAccept;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numericUpDown2;
        private CustomDataGrid TableColumns;
        private System.Windows.Forms.Button BtnAddProject;
        private System.Windows.Forms.Label label3;
        private CustomDataGrid TableProjects;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox TboxProjectName;
        private System.Windows.Forms.Button BtnDelete;
        private System.Windows.Forms.Button BtnSelect;
        private System.Windows.Forms.Button BtnUpdateColumns;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox TextBoxColumn;
        private System.Windows.Forms.TextBox TextBoxRow;
        private System.Windows.Forms.TextBox TextBoxCellName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button BtnAdd;
        private System.Windows.Forms.Button BtnActiveCell;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox TextBoxAddres;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox TextBoxValue;
        private System.Windows.Forms.CheckBox ChkBoxObligatory;
        private System.Windows.Forms.CheckBox ChkBoxCheck;
        private System.Windows.Forms.Button BtnCheckCells;
    }
}