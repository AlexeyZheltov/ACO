
namespace ACO
{
    partial class FrmColorCommentsFomat
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmColorCommentsFomat));
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.colorDialog = new System.Windows.Forms.ColorDialog();
            this.BtnInteriorColor = new System.Windows.Forms.Button();
            this.BtnAccept = new System.Windows.Forms.Button();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.BtnForeColor = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ChkBoxBold = new System.Windows.Forms.CheckBox();
            this.BtnAdd = new System.Windows.Forms.Button();
            this.RulesDataGrid = new ACO.ProjectManager.CustomDataGrid();
            this.ColumnID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.column1 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.Column4 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.RulesDataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.richTextBox1.DetectUrls = false;
            this.richTextBox1.HideSelection = false;
            this.richTextBox1.Location = new System.Drawing.Point(13, 161);
            this.richTextBox1.Multiline = false;
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
            this.richTextBox1.Size = new System.Drawing.Size(91, 30);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = " Формат   0%";
            // 
            // BtnInteriorColor
            // 
            this.BtnInteriorColor.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.BtnInteriorColor.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BtnInteriorColor.Location = new System.Drawing.Point(13, 34);
            this.BtnInteriorColor.Margin = new System.Windows.Forms.Padding(1);
            this.BtnInteriorColor.Name = "BtnInteriorColor";
            this.BtnInteriorColor.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.BtnInteriorColor.Size = new System.Drawing.Size(91, 30);
            this.BtnInteriorColor.TabIndex = 1;
            this.BtnInteriorColor.Text = "Цвет заливки";
            this.BtnInteriorColor.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnInteriorColor.UseVisualStyleBackColor = true;
            this.BtnInteriorColor.Click += new System.EventHandler(this.BtnInteriorColor_Click);
            // 
            // BtnAccept
            // 
            this.BtnAccept.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnAccept.Location = new System.Drawing.Point(445, 251);
            this.BtnAccept.Name = "BtnAccept";
            this.BtnAccept.Size = new System.Drawing.Size(93, 26);
            this.BtnAccept.TabIndex = 3;
            this.BtnAccept.Text = "Принять";
            this.BtnAccept.UseVisualStyleBackColor = true;
            this.BtnAccept.Click += new System.EventHandler(this.BtnAccept_Click);
            // 
            // BtnCancel
            // 
            this.BtnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(543, 251);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(93, 26);
            this.BtnCancel.TabIndex = 3;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.ChkBoxBold);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.richTextBox1);
            this.groupBox1.Controls.Add(this.BtnForeColor);
            this.groupBox1.Controls.Add(this.BtnInteriorColor);
            this.groupBox1.Location = new System.Drawing.Point(523, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(117, 206);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Формат";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(35, 144);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(44, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Ячейка";
            // 
            // BtnForeColor
            // 
            this.BtnForeColor.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.BtnForeColor.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.BtnForeColor.Location = new System.Drawing.Point(13, 72);
            this.BtnForeColor.Margin = new System.Windows.Forms.Padding(1);
            this.BtnForeColor.Name = "BtnForeColor";
            this.BtnForeColor.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.BtnForeColor.Size = new System.Drawing.Size(91, 30);
            this.BtnForeColor.TabIndex = 1;
            this.BtnForeColor.Text = "Цвет шрифта";
            this.BtnForeColor.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.BtnForeColor.UseVisualStyleBackColor = true;
            this.BtnForeColor.Click += new System.EventHandler(this.BtnForeColor_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.RulesDataGrid);
            this.groupBox2.Location = new System.Drawing.Point(13, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(498, 231);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Правила ";
            // 
            // ChkBoxBold
            // 
            this.ChkBoxBold.AutoSize = true;
            this.ChkBoxBold.Location = new System.Drawing.Point(13, 110);
            this.ChkBoxBold.Name = "ChkBoxBold";
            this.ChkBoxBold.Size = new System.Drawing.Size(91, 17);
            this.ChkBoxBold.TabIndex = 4;
            this.ChkBoxBold.Text = "Полужирный";
            this.ChkBoxBold.UseVisualStyleBackColor = true;
            this.ChkBoxBold.CheckedChanged += new System.EventHandler(this.ChkBoxBold_CheckedChanged);
            // 
            // BtnAdd
            // 
            this.BtnAdd.Location = new System.Drawing.Point(19, 253);
            this.BtnAdd.Name = "BtnAdd";
            this.BtnAdd.Size = new System.Drawing.Size(75, 23);
            this.BtnAdd.TabIndex = 6;
            this.BtnAdd.Text = "Добавить";
            this.BtnAdd.UseVisualStyleBackColor = true;
            this.BtnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // RulesDataGrid
            // 
            this.RulesDataGrid.AllowUserToAddRows = false;
            this.RulesDataGrid.AllowUserToResizeRows = false;
            this.RulesDataGrid.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.RulesDataGrid.BackgroundColor = System.Drawing.Color.White;
            this.RulesDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.RulesDataGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnID,
            this.column1,
            this.Column4,
            this.Column2,
            this.Column3});
            this.RulesDataGrid.Location = new System.Drawing.Point(6, 19);
            this.RulesDataGrid.MultiSelect = false;
            this.RulesDataGrid.Name = "RulesDataGrid";
            this.RulesDataGrid.RowHeadersVisible = false;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.RulesDataGrid.RowsDefaultCellStyle = dataGridViewCellStyle1;
            this.RulesDataGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.RulesDataGrid.Size = new System.Drawing.Size(486, 206);
            this.RulesDataGrid.TabIndex = 0;
            this.RulesDataGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.CustomDataGrid_CellClick);
            this.RulesDataGrid.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.CustomDataGrid_CellValueChanged);
            // 
            // ColumnID
            // 
            this.ColumnID.HeaderText = "Номер";
            this.ColumnID.Name = "ColumnID";
            this.ColumnID.ReadOnly = true;
            this.ColumnID.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.ColumnID.Visible = false;
            // 
            // column1
            // 
            this.column1.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.column1.HeaderText = "Столбец";
            this.column1.Items.AddRange(new object[] {
            "",
            "Комментарии к описанию работ",
            "Отклонение по объемам",
            "Комментарии к объемам работ",
            "Отклонение по стоимости",
            "Комментарии к стоимости работ",
            "Отклонение МАТ",
            "Отклонение РАБ",
            "Комментарии к строкам",
            "Выделение"});
            this.column1.MinimumWidth = 50;
            this.column1.Name = "column1";
            this.column1.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.column1.Width = 220;
            // 
            // Column4
            // 
            this.Column4.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.ComboBox;
            this.Column4.HeaderText = "Условие";
            this.Column4.Items.AddRange(new object[] {
            "",
            "Равно",
            "Не равно",
            "Больше",
            "Больше равно",
            "Меньше",
            "Меньше равно",
            "Между",
            "Содержит"});
            this.Column4.Name = "Column4";
            this.Column4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // Column2
            // 
            this.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Column2.HeaderText = "Значение1";
            this.Column2.Name = "Column2";
            this.Column2.Width = 86;
            // 
            // Column3
            // 
            this.Column3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Column3.HeaderText = "Значение2";
            this.Column3.Name = "Column3";
            this.Column3.Width = 86;
            // 
            // FrmColorCommentsFomat
            // 
            this.AcceptButton = this.BtnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(648, 284);
            this.Controls.Add(this.BtnAdd);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.BtnAccept);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmColorCommentsFomat";
            this.Text = "Форматирование комментариев";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.RulesDataGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.ColorDialog colorDialog;
        private System.Windows.Forms.Button BtnInteriorColor;
        private System.Windows.Forms.Button BtnAccept;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button BtnForeColor;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private ProjectManager.CustomDataGrid RulesDataGrid;
        private System.Windows.Forms.CheckBox ChkBoxBold;
        private System.Windows.Forms.Button BtnAdd;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnID;
        private System.Windows.Forms.DataGridViewComboBoxColumn column1;
        private System.Windows.Forms.DataGridViewComboBoxColumn Column4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
    }
}