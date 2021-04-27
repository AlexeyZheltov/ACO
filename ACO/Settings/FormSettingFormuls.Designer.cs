
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnSave = new System.Windows.Forms.Button();
            this.Rbtn1 = new System.Windows.Forms.RadioButton();
            this.Rbtn2 = new System.Windows.Forms.RadioButton();
            this.Rbtn3 = new System.Windows.Forms.RadioButton();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.Rbtn3);
            this.panel1.Controls.Add(this.Rbtn2);
            this.panel1.Controls.Add(this.Rbtn1);
            this.panel1.Location = new System.Drawing.Point(5, 10);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(235, 97);
            this.panel1.TabIndex = 0;
            // 
            // BtnCancel
            // 
            this.BtnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(155, 112);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(77, 23);
            this.BtnCancel.TabIndex = 1;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            // 
            // BtnSave
            // 
            this.BtnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSave.Location = new System.Drawing.Point(67, 112);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(82, 23);
            this.BtnSave.TabIndex = 1;
            this.BtnSave.Text = "Сохранить";
            this.BtnSave.UseVisualStyleBackColor = true;
            // 
            // Rbtn1
            // 
            this.Rbtn1.AutoSize = true;
            this.Rbtn1.Location = new System.Drawing.Point(29, 16);
            this.Rbtn1.Name = "Rbtn1";
            this.Rbtn1.Size = new System.Drawing.Size(68, 17);
            this.Rbtn1.TabIndex = 0;
            this.Rbtn1.TabStop = true;
            this.Rbtn1.Text = "Среднее";
            this.Rbtn1.UseVisualStyleBackColor = true;
            // 
            // Rbtn2
            // 
            this.Rbtn2.AutoSize = true;
            this.Rbtn2.Location = new System.Drawing.Point(29, 39);
            this.Rbtn2.Name = "Rbtn2";
            this.Rbtn2.Size = new System.Drawing.Size(70, 17);
            this.Rbtn2.TabIndex = 0;
            this.Rbtn2.TabStop = true;
            this.Rbtn2.Text = "Медиана";
            this.Rbtn2.UseVisualStyleBackColor = true;
            // 
            // Rbtn3
            // 
            this.Rbtn3.AutoSize = true;
            this.Rbtn3.Location = new System.Drawing.Point(29, 62);
            this.Rbtn3.Name = "Rbtn3";
            this.Rbtn3.Size = new System.Drawing.Size(75, 17);
            this.Rbtn3.TabIndex = 0;
            this.Rbtn3.TabStop = true;
            this.Rbtn3.Text = "Формулы";
            this.Rbtn3.UseVisualStyleBackColor = true;
            // 
            // FormSettingFormuls
            // 
            this.AcceptButton = this.BtnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(244, 140);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.BtnCancel);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormSettingFormuls";
            this.Text = "Настройки анализа";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button BtnSave;
        private System.Windows.Forms.RadioButton Rbtn3;
        private System.Windows.Forms.RadioButton Rbtn2;
        private System.Windows.Forms.RadioButton Rbtn1;
    }
}