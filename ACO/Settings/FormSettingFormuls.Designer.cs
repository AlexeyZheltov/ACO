
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
            this.RbtnBaseCost0 = new System.Windows.Forms.RadioButton();
            this.RbtnCostMedian2 = new System.Windows.Forms.RadioButton();
            this.RbtnAvgCost1 = new System.Windows.Forms.RadioButton();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.RbtnBaseCount0 = new System.Windows.Forms.RadioButton();
            this.RbtnAvgCount1 = new System.Windows.Forms.RadioButton();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // BtnCancel
            // 
            this.BtnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(289, 229);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(77, 23);
            this.BtnCancel.TabIndex = 7;
            this.BtnCancel.Text = "Отмена";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnSave
            // 
            this.BtnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnSave.Location = new System.Drawing.Point(201, 229);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Size = new System.Drawing.Size(82, 23);
            this.BtnSave.TabIndex = 6;
            this.BtnSave.Text = "Сохранить";
            this.BtnSave.UseVisualStyleBackColor = true;
            this.BtnSave.Click += new System.EventHandler(this.BtnSave_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.TBoxBottom);
            this.groupBox1.Controls.Add(this.TBoxTop);
            this.groupBox1.Location = new System.Drawing.Point(8, 135);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(177, 88);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Границы оценок в формулах";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(61, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Нижняя, %";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Верхняя, %";
            // 
            // TBoxBottom
            // 
            this.TBoxBottom.Location = new System.Drawing.Point(83, 57);
            this.TBoxBottom.Name = "TBoxBottom";
            this.TBoxBottom.Size = new System.Drawing.Size(82, 20);
            this.TBoxBottom.TabIndex = 5;
            this.TBoxBottom.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TBoxTop_KeyPress);
            // 
            // TBoxTop
            // 
            this.TBoxTop.Location = new System.Drawing.Point(83, 31);
            this.TBoxTop.Name = "TBoxTop";
            this.TBoxTop.Size = new System.Drawing.Size(82, 20);
            this.TBoxTop.TabIndex = 4;
            this.TBoxTop.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TBoxTop_KeyPress);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.groupBox2.Controls.Add(this.RbtnBaseCost0);
            this.groupBox2.Controls.Add(this.RbtnCostMedian2);
            this.groupBox2.Controls.Add(this.RbtnAvgCost1);
            this.groupBox2.Location = new System.Drawing.Point(7, 25);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(178, 100);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Стоимость";
            // 
            // RbtnBaseCost0
            // 
            this.RbtnBaseCost0.AutoSize = true;
            this.RbtnBaseCost0.Location = new System.Drawing.Point(16, 27);
            this.RbtnBaseCost0.Name = "RbtnBaseCost0";
            this.RbtnBaseCost0.Size = new System.Drawing.Size(150, 17);
            this.RbtnBaseCost0.TabIndex = 1;
            this.RbtnBaseCost0.TabStop = true;
            this.RbtnBaseCost0.Text = "Отклонение от базового";
            this.RbtnBaseCost0.UseVisualStyleBackColor = true;
            // 
            // RbtnCostMedian2
            // 
            this.RbtnCostMedian2.AutoSize = true;
            this.RbtnCostMedian2.Location = new System.Drawing.Point(16, 73);
            this.RbtnCostMedian2.Name = "RbtnCostMedian2";
            this.RbtnCostMedian2.Size = new System.Drawing.Size(70, 17);
            this.RbtnCostMedian2.TabIndex = 3;
            this.RbtnCostMedian2.TabStop = true;
            this.RbtnCostMedian2.Text = "Медиана";
            this.RbtnCostMedian2.UseVisualStyleBackColor = true;
            // 
            // RbtnAvgCost1
            // 
            this.RbtnAvgCost1.AutoSize = true;
            this.RbtnAvgCost1.Location = new System.Drawing.Point(16, 50);
            this.RbtnAvgCost1.Name = "RbtnAvgCost1";
            this.RbtnAvgCost1.Size = new System.Drawing.Size(68, 17);
            this.RbtnAvgCost1.TabIndex = 2;
            this.RbtnAvgCost1.TabStop = true;
            this.RbtnAvgCost1.Text = "Среднее";
            this.RbtnAvgCost1.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.groupBox3.Controls.Add(this.RbtnBaseCount0);
            this.groupBox3.Controls.Add(this.RbtnAvgCount1);
            this.groupBox3.Location = new System.Drawing.Point(191, 25);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(178, 100);
            this.groupBox3.TabIndex = 4;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Количество";
            // 
            // RbtnBaseCount0
            // 
            this.RbtnBaseCount0.AutoSize = true;
            this.RbtnBaseCount0.Location = new System.Drawing.Point(13, 27);
            this.RbtnBaseCount0.Name = "RbtnBaseCount0";
            this.RbtnBaseCount0.Size = new System.Drawing.Size(150, 17);
            this.RbtnBaseCount0.TabIndex = 1;
            this.RbtnBaseCount0.TabStop = true;
            this.RbtnBaseCount0.Text = "Отклонение от базового";
            this.RbtnBaseCount0.UseVisualStyleBackColor = true;
            // 
            // RbtnAvgCount1
            // 
            this.RbtnAvgCount1.AutoSize = true;
            this.RbtnAvgCount1.Location = new System.Drawing.Point(13, 50);
            this.RbtnAvgCount1.Name = "RbtnAvgCount1";
            this.RbtnAvgCount1.Size = new System.Drawing.Size(68, 17);
            this.RbtnAvgCount1.TabIndex = 2;
            this.RbtnAvgCount1.TabStop = true;
            this.RbtnAvgCount1.Text = "Среднее";
            this.RbtnAvgCount1.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 6);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Вид анализа";
            // 
            // FormSettingFormuls
            // 
            this.AcceptButton = this.BtnSave;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(378, 257);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.BtnSave);
            this.Controls.Add(this.BtnCancel);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormSettingFormuls";
            this.Text = "Настройки анализа";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Closing);
            this.Load += new System.EventHandler(this.FormSettingFormuls_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

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
        private System.Windows.Forms.RadioButton RbtnBaseCost0;
        private System.Windows.Forms.RadioButton RbtnCostMedian2;
        private System.Windows.Forms.RadioButton RbtnAvgCost1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton RbtnBaseCount0;
        private System.Windows.Forms.RadioButton RbtnAvgCount1;
        private System.Windows.Forms.Label label3;
    }
}