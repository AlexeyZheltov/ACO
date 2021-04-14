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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormManager));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.BtnAccept = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.BtnSetCurrentSheet = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.TBoxSheetName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.TBoxFirstRowRangeValues = new System.Windows.Forms.TextBox();
            this.PageColumns = new System.Windows.Forms.TabPage();
            this.TableColumns = new ACO.ProjectManager.CustomDataGrid();
            this.PageProject = new System.Windows.Forms.TabPage();
            this.BtnOpenFolder = new System.Windows.Forms.Button();
            this.BtnDelete = new System.Windows.Forms.Button();
            this.BtnSelect = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.TboxProjectName = new System.Windows.Forms.TextBox();
            this.BtnAddProject = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.TableProjects = new ACO.ProjectManager.CustomDataGrid();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2.SuspendLayout();
            this.PageColumns.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).BeginInit();
            this.PageProject.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnAccept
            // 
            this.BtnAccept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAccept.Location = new System.Drawing.Point(272, 487);
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
            this.BtnCancel.Location = new System.Drawing.Point(384, 487);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(106, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.BtnSetCurrentSheet);
            this.tabPage2.Controls.Add(this.label6);
            this.tabPage2.Controls.Add(this.TBoxSheetName);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.TBoxFirstRowRangeValues);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(494, 446);
            this.tabPage2.TabIndex = 3;
            this.tabPage2.Text = "Диапазон";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // BtnSetCurrentSheet
            // 
            this.BtnSetCurrentSheet.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSetCurrentSheet.Location = new System.Drawing.Point(286, 13);
            this.BtnSetCurrentSheet.Name = "BtnSetCurrentSheet";
            this.BtnSetCurrentSheet.Size = new System.Drawing.Size(94, 30);
            this.BtnSetCurrentSheet.TabIndex = 13;
            this.BtnSetCurrentSheet.Text = "Текущий лист";
            this.BtnSetCurrentSheet.UseVisualStyleBackColor = true;
            this.BtnSetCurrentSheet.Click += new System.EventHandler(this.BtnSetCurrentSheet_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(46, 22);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 13);
            this.label6.TabIndex = 12;
            this.label6.Text = "Лист Анализ";
            // 
            // TBoxSheetName
            // 
            this.TBoxSheetName.Location = new System.Drawing.Point(186, 19);
            this.TBoxSheetName.Name = "TBoxSheetName";
            this.TBoxSheetName.Size = new System.Drawing.Size(80, 20);
            this.TBoxSheetName.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(46, 51);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(124, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Строка начала данных ";
            // 
            // TBoxFirstRowRangeValues
            // 
            this.TBoxFirstRowRangeValues.Location = new System.Drawing.Point(186, 48);
            this.TBoxFirstRowRangeValues.Name = "TBoxFirstRowRangeValues";
            this.TBoxFirstRowRangeValues.Size = new System.Drawing.Size(80, 20);
            this.TBoxFirstRowRangeValues.TabIndex = 1;
            // 
            // PageColumns
            // 
            this.PageColumns.Controls.Add(this.TableColumns);
            this.PageColumns.Location = new System.Drawing.Point(4, 22);
            this.PageColumns.Name = "PageColumns";
            this.PageColumns.Padding = new System.Windows.Forms.Padding(3);
            this.PageColumns.Size = new System.Drawing.Size(494, 446);
            this.PageColumns.TabIndex = 0;
            this.PageColumns.Text = "Столбцы";
            this.PageColumns.UseVisualStyleBackColor = true;
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
            this.TableColumns.Location = new System.Drawing.Point(6, 6);
            this.TableColumns.MultiSelect = false;
            this.TableColumns.Name = "TableColumns";
            this.TableColumns.ReadOnly = true;
            this.TableColumns.RowHeadersVisible = false;
            this.TableColumns.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.TableColumns.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.TableColumns.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.TableColumns.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableColumns.Size = new System.Drawing.Size(482, 433);
            this.TableColumns.TabIndex = 0;
            this.TableColumns.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableColumns_CellValueChanged);
            this.TableColumns.RowsRemoved += new System.Windows.Forms.DataGridViewRowsRemovedEventHandler(this.TableColumns_RowsRemoved);
            // 
            // PageProject
            // 
            this.PageProject.Controls.Add(this.BtnOpenFolder);
            this.PageProject.Controls.Add(this.BtnDelete);
            this.PageProject.Controls.Add(this.BtnSelect);
            this.PageProject.Controls.Add(this.groupBox2);
            this.PageProject.Controls.Add(this.label3);
            this.PageProject.Controls.Add(this.TableProjects);
            this.PageProject.Location = new System.Drawing.Point(4, 22);
            this.PageProject.Name = "PageProject";
            this.PageProject.Padding = new System.Windows.Forms.Padding(3);
            this.PageProject.Size = new System.Drawing.Size(494, 446);
            this.PageProject.TabIndex = 1;
            this.PageProject.Text = "Проект";
            this.PageProject.UseVisualStyleBackColor = true;
            // 
            // BtnOpenFolder
            // 
            this.BtnOpenFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnOpenFolder.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("BtnOpenFolder.BackgroundImage")));
            this.BtnOpenFolder.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.BtnOpenFolder.Location = new System.Drawing.Point(454, 77);
            this.BtnOpenFolder.Name = "BtnOpenFolder";
            this.BtnOpenFolder.Size = new System.Drawing.Size(27, 27);
            this.BtnOpenFolder.TabIndex = 10;
            this.BtnOpenFolder.UseVisualStyleBackColor = true;
            this.BtnOpenFolder.Click += new System.EventHandler(this.BtnOpenFolder_Click);
            // 
            // BtnDelete
            // 
            this.BtnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnDelete.Location = new System.Drawing.Point(377, 78);
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
            this.BtnSelect.Location = new System.Drawing.Point(296, 78);
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
            this.groupBox2.Size = new System.Drawing.Size(473, 60);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            // 
            // TboxProjectName
            // 
            this.TboxProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TboxProjectName.Location = new System.Drawing.Point(129, 22);
            this.TboxProjectName.Name = "TboxProjectName";
            this.TboxProjectName.Size = new System.Drawing.Size(244, 20);
            this.TboxProjectName.TabIndex = 3;
            // 
            // BtnAddProject
            // 
            this.BtnAddProject.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAddProject.Location = new System.Drawing.Point(388, 17);
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
            this.label3.Location = new System.Drawing.Point(23, 89);
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
            this.TableProjects.Location = new System.Drawing.Point(6, 110);
            this.TableProjects.MultiSelect = false;
            this.TableProjects.Name = "TableProjects";
            this.TableProjects.ReadOnly = true;
            this.TableProjects.RowHeadersVisible = false;
            this.TableProjects.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.TableProjects.RowsDefaultCellStyle = dataGridViewCellStyle2;
            this.TableProjects.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableProjects.Size = new System.Drawing.Size(482, 333);
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
            this.tabControl1.Location = new System.Drawing.Point(3, 15);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(502, 472);
            this.tabControl1.TabIndex = 0;
            // 
            // FormManager
            // 
            this.AcceptButton = this.BtnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(508, 512);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnAccept);
            this.Controls.Add(this.tabControl1);
            this.Name = "FormManager";
            this.ShowIcon = false;
            this.Text = "Диспетчер проектов";
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.PageColumns.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).EndInit();
            this.PageProject.ResumeLayout(false);
            this.PageProject.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.TableProjects)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button BtnAccept;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox TBoxFirstRowRangeValues;
        private System.Windows.Forms.TabPage PageColumns;
        private CustomDataGrid TableColumns;
        private System.Windows.Forms.TabPage PageProject;
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
        private System.Windows.Forms.Button BtnOpenFolder;
        private System.Windows.Forms.Button BtnSetCurrentSheet;
    }
}