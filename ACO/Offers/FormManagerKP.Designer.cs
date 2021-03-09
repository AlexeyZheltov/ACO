namespace ACO.Offers
{
    partial class FormManagerKP
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
            this.BtnAddColumns = new System.Windows.Forms.Button();
            this.BtnSave = new System.Windows.Forms.Button();
            this.ListKP = new System.Windows.Forms.ListView();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.BtnCreate = new System.Windows.Forms.Button();
            this.projectNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.customerDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.projectNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dateDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.ColHeader = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.TableColumns = new ACO.ProjectManager.CustomDataGrid();
            this.customDataGrid1 = new ACO.ProjectManager.CustomDataGrid();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.customDataGrid1)).BeginInit();
            this.SuspendLayout();
            // 
            // BtnAddColumns
            // 
            this.BtnAddColumns.Location = new System.Drawing.Point(406, 12);
            this.BtnAddColumns.Name = "BtnAddColumns";
            this.BtnAddColumns.Size = new System.Drawing.Size(75, 23);
            this.BtnAddColumns.TabIndex = 2;
            this.BtnAddColumns.Text = "Обновить";
            this.BtnAddColumns.UseVisualStyleBackColor = true;
            this.BtnAddColumns.Click += new System.EventHandler(this.BtnAddColumns_Click);
            // 
            // BtnSave
            // 
            this.BtnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSave.Location = new System.Drawing.Point(736, 460);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(75, 23);
            this.BtnSave.TabIndex = 2;
            this.BtnSave.Text = "Сохранить";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // ListKP
            // 
            this.ListKP.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.ColName,
            this.ColHeader});
            this.ListKP.HideSelection = false;
            this.ListKP.Location = new System.Drawing.Point(13, 38);
            this.ListKP.Name = "ListKP";
            this.ListKP.Size = new System.Drawing.Size(468, 115);
            this.ListKP.TabIndex = 3;
            this.ListKP.UseCompatibleStateImageBehavior = false;
            this.ListKP.View = System.Windows.Forms.View.Details;
            this.ListKP.SelectedIndexChanged += new System.EventHandler(this.ListKP_SelectedIndexChanged);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(13, 12);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(296, 20);
            this.textBox1.TabIndex = 4;
            // 
            // BtnCreate
            // 
            this.BtnCreate.Location = new System.Drawing.Point(315, 11);
            this.BtnCreate.Name = "BtnCreate";
            this.BtnCreate.Size = new System.Drawing.Size(85, 23);
            this.BtnCreate.TabIndex = 5;
            this.BtnCreate.Text = "Добавить";
            this.BtnCreate.UseVisualStyleBackColor = true;
            this.BtnCreate.Click += new System.EventHandler(this.BtnCreate_Click);
            // 
            // projectNameDataGridViewTextBoxColumn
            // 
            this.projectNameDataGridViewTextBoxColumn.DataPropertyName = "ProjectName";
            this.projectNameDataGridViewTextBoxColumn.HeaderText = "ProjectName";
            this.projectNameDataGridViewTextBoxColumn.Name = "projectNameDataGridViewTextBoxColumn";
            // 
            // customerDataGridViewTextBoxColumn
            // 
            this.customerDataGridViewTextBoxColumn.DataPropertyName = "Customer";
            this.customerDataGridViewTextBoxColumn.HeaderText = "Customer";
            this.customerDataGridViewTextBoxColumn.Name = "customerDataGridViewTextBoxColumn";
            // 
            // projectNumberDataGridViewTextBoxColumn
            // 
            this.projectNumberDataGridViewTextBoxColumn.DataPropertyName = "ProjectNumber";
            this.projectNumberDataGridViewTextBoxColumn.HeaderText = "ProjectNumber";
            this.projectNumberDataGridViewTextBoxColumn.Name = "projectNumberDataGridViewTextBoxColumn";
            // 
            // dateDataGridViewTextBoxColumn
            // 
            this.dateDataGridViewTextBoxColumn.DataPropertyName = "Date";
            this.dateDataGridViewTextBoxColumn.HeaderText = "Date";
            this.dateDataGridViewTextBoxColumn.Name = "dateDataGridViewTextBoxColumn";
            // 
            // ColName
            // 
            this.ColName.Text = "Название";
            this.ColName.Width = 186;
            // 
            // ColHeader
            // 
            this.ColHeader.Text = "Путь";
            this.ColHeader.Width = 274;
            // 
            // TableColumns
            // 
            this.TableColumns.AllowUserToAddRows = false;
            this.TableColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableColumns.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TableColumns.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.TableColumns.Location = new System.Drawing.Point(13, 159);
            this.TableColumns.MultiSelect = false;
            this.TableColumns.Name = "TableColumns";
            this.TableColumns.RowHeadersVisible = false;
            this.TableColumns.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.TableColumns.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableColumns.Size = new System.Drawing.Size(798, 295);
            this.TableColumns.TabIndex = 1;
            this.TableColumns.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.TableColumns_CellValueChanged);
            // 
            // customDataGrid1
            // 
            this.customDataGrid1.AllowUserToAddRows = false;
            this.customDataGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.customDataGrid1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.customDataGrid1.Location = new System.Drawing.Point(487, 11);
            this.customDataGrid1.MultiSelect = false;
            this.customDataGrid1.Name = "customDataGrid1";
            this.customDataGrid1.RowHeadersVisible = false;
            this.customDataGrid1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.customDataGrid1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.customDataGrid1.Size = new System.Drawing.Size(324, 88);
            this.customDataGrid1.TabIndex = 0;
            // 
            // FormManagerKP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(821, 490);
            this.Controls.Add(this.BtnCreate);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.ListKP);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.BtnAddColumns);
            this.Controls.Add(this.TableColumns);
            this.Controls.Add(this.customDataGrid1);
            this.Name = "FormManagerKP";
            this.Text = "FormManagerKP";
            this.Load += new System.EventHandler(this.FormManagerKP_Load);
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.customDataGrid1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ProjectManager.CustomDataGrid customDataGrid1;
        private System.Windows.Forms.DataGridViewTextBoxColumn projectNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn customerDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn projectNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn dateDataGridViewTextBoxColumn;
        private ProjectManager.CustomDataGrid TableColumns;
        private System.Windows.Forms.Button BtnAddColumns;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.ListView ListKP;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button BtnCreate;
        private System.Windows.Forms.ColumnHeader ColName;
        private System.Windows.Forms.ColumnHeader ColHeader;
    }
}