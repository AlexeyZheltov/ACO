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
            this.components = new System.ComponentModel.Container();
            this.customDataGrid1 = new ACO.ProjectManager.CustomDataGrid();
            this.projectNameDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.customerDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.projectNumberDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dateDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.offerBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.TableColumns = new ACO.ProjectManager.CustomDataGrid();
            this.BtnAddColumns = new System.Windows.Forms.Button();
            this.BtnSave = new System.Windows.Forms.Button();
            this.ListKP = new System.Windows.Forms.ListView();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.BtnCreate = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.customDataGrid1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.offerBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).BeginInit();
            this.SuspendLayout();
            // 
            // customDataGrid1
            // 
            this.customDataGrid1.AllowUserToAddRows = false;
            this.customDataGrid1.AutoGenerateColumns = false;
            this.customDataGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.customDataGrid1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.projectNameDataGridViewTextBoxColumn,
            this.customerDataGridViewTextBoxColumn,
            this.projectNumberDataGridViewTextBoxColumn,
            this.dateDataGridViewTextBoxColumn});
            this.customDataGrid1.DataSource = this.offerBindingSource;
            this.customDataGrid1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.customDataGrid1.Location = new System.Drawing.Point(12, 43);
            this.customDataGrid1.MultiSelect = false;
            this.customDataGrid1.Name = "customDataGrid1";
            this.customDataGrid1.RowHeadersVisible = false;
            this.customDataGrid1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.customDataGrid1.Size = new System.Drawing.Size(406, 52);
            this.customDataGrid1.TabIndex = 0;
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
            // offerBindingSource
            // 
            this.offerBindingSource.DataSource = typeof(ACO.Offer);
            // 
            // TableColumns
            // 
            this.TableColumns.AllowUserToAddRows = false;
            this.TableColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TableColumns.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.TableColumns.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.TableColumns.Location = new System.Drawing.Point(13, 130);
            this.TableColumns.MultiSelect = false;
            this.TableColumns.Name = "TableColumns";
            this.TableColumns.RowHeadersVisible = false;
            this.TableColumns.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.TableColumns.Size = new System.Drawing.Size(800, 193);
            this.TableColumns.TabIndex = 1;
            // 
            // BtnAddColumns
            // 
            this.BtnAddColumns.Location = new System.Drawing.Point(13, 101);
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
            this.BtnSave.Location = new System.Drawing.Point(738, 329);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(75, 23);
            this.BtnSave.TabIndex = 2;
            this.BtnSave.Text = "Сохранить";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // ListKP
            // 
            this.ListKP.HideSelection = false;
            this.ListKP.Location = new System.Drawing.Point(425, 12);
            this.ListKP.Name = "ListKP";
            this.ListKP.Size = new System.Drawing.Size(388, 112);
            this.ListKP.TabIndex = 3;
            this.ListKP.UseCompatibleStateImageBehavior = false;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(13, 12);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(240, 20);
            this.textBox1.TabIndex = 4;
            // 
            // BtnCreate
            // 
            this.BtnCreate.Location = new System.Drawing.Point(260, 11);
            this.BtnCreate.Name = "BtnCreate";
            this.BtnCreate.Size = new System.Drawing.Size(110, 23);
            this.BtnCreate.TabIndex = 5;
            this.BtnCreate.Text = "Добавить";
            this.BtnCreate.UseVisualStyleBackColor = true;
            this.BtnCreate.Click += new System.EventHandler(this.BtnCreate_Click);
            // 
            // FormManagerKP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(823, 359);
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
            ((System.ComponentModel.ISupportInitialize)(this.customDataGrid1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.offerBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.TableColumns)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ProjectManager.CustomDataGrid customDataGrid1;
        private System.Windows.Forms.DataGridViewTextBoxColumn projectNameDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn customerDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn projectNumberDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn dateDataGridViewTextBoxColumn;
        private System.Windows.Forms.BindingSource offerBindingSource;
        private ProjectManager.CustomDataGrid TableColumns;
        private System.Windows.Forms.Button BtnAddColumns;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.ListView ListKP;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button BtnCreate;
    }
}