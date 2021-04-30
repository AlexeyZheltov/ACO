
namespace ACO.Settings
{
    partial class FormSettingFormuls
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormSettingFormuls));
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnSave = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.TBoxBottom = new System.Windows.Forms.TextBox();
            this.TBoxTop = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.Rbtn0 = new System.Windows.Forms.RadioButton();
            this.Rbtn2 = new System.Windows.Forms.RadioButton();
            this.Rbtn1 = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnCancel
            // 
            this.BtnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(155, 215);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(77, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnSave
            // 
            this.BtnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSave.Location = new System.Drawing.Point(67, 215);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(82, 23);
            this.BtnSave.TabIndex = 1;
            this.BtnSave.Text = "Сохранить";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.TBoxBottom);
            this.groupBox1.Controls.Add(this.TBoxTop);
            this.groupBox1.Location = new System.Drawing.Point(8, 121);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(232, 88);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Границы оценок в формулах";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Нижняя, %";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Верхняя, %";
            // 
            // TBoxBottom
            // 
            this.TBoxBottom.Location = new System.Drawing.Point(117, 57);
            this.TBoxBottom.Name = "TBoxBottom";
            this.TBoxBottom.Size = new System.Drawing.Size(92, 20);
            this.TBoxBottom.TabIndex = 0;
            // 
            // TBoxTop
            // 
            this.TBoxTop.Location = new System.Drawing.Point(117, 31);
            this.TBoxTop.Name = "TBoxTop";
            this.TBoxTop.Size = new System.Drawing.Size(92, 20);
            this.TBoxTop.TabIndex = 0;
            this.TBoxTop.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TBoxTop_KeyPress);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.Rbtn0);
            this.groupBox2.Controls.Add(this.Rbtn2);
            this.groupBox2.Controls.Add(this.Rbtn1);
            this.groupBox2.Location = new System.Drawing.Point(7, 12);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(232, 100);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Вид анализа";
            // 
            // Rbtn0
            // 
            this.Rbtn0.AutoSize = true;
            this.Rbtn0.Location = new System.Drawing.Point(25, 27);
            this.Rbtn0.Name = "Rbtn0";
            this.Rbtn0.Size = new System.Drawing.Size(150, 17);
            this.Rbtn0.TabIndex = 1;
            this.Rbtn0.TabStop = true;
            this.Rbtn0.Text = "Отклонение от базового";
            this.Rbtn0.UseVisualStyleBackColor = true;
            // 
            // Rbtn2
            // 
            this.Rbtn2.AutoSize = true;
            this.Rbtn2.Location = new System.Drawing.Point(25, 73);
            this.Rbtn2.Name = "Rbtn2";
            this.Rbtn2.Size = new System.Drawing.Size(70, 17);
            this.Rbtn2.TabIndex = 2;
            this.Rbtn2.TabStop = true;
            this.Rbtn2.Text = "Медиана";
            this.Rbtn2.UseVisualStyleBackColor = true;
            // 
            // Rbtn1
            // 
            this.Rbtn1.AutoSize = true;
            this.Rbtn1.Location = new System.Drawing.Point(25, 50);
            this.Rbtn1.Name = "Rbtn1";
            this.Rbtn1.Size = new System.Drawing.Size(68, 17);
            this.Rbtn1.TabIndex = 3;
            this.Rbtn1.TabStop = true;
            this.Rbtn1.Text = "Среднее";
            this.Rbtn1.UseVisualStyleBackColor = true;
            // 
            // FormSettingFormuls
            // 
            this.AcceptButton = this.BtnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(244, 243);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.BtnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(260, 278);
            this.MinimumSize = new System.Drawing.Size(260, 278);
            this.Name = "FormSettingFormuls";
            this.Text = "Настройки анализа";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Closing);
            this.Load += new System.EventHandler(this.FormSettingFormuls_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox TBoxTop;
        private System.Windows.Forms.TextBox TBoxBottom;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton Rbtn0;
        private System.Windows.Forms.RadioButton Rbtn2;
        private System.Windows.Forms.RadioButton Rbtn1;
    }
}