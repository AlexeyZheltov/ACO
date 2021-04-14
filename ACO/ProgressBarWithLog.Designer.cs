
namespace ACO
{
    partial class ProgressBarWithLog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProgressBarWithLog));
            this.LogTextBox = new System.Windows.Forms.TextBox();
            this.SubProgressBar = new System.Windows.Forms.ProgressBar();
            this.SubLabel = new System.Windows.Forms.Label();
            this.MainLabel = new System.Windows.Forms.Label();
            this.MainProgressBar = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.LogGroupBox = new System.Windows.Forms.GroupBox();
            this.OpenLogFolderButton = new System.Windows.Forms.Button();
            this.AbortButton = new System.Windows.Forms.Button();
            this.LogGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // LogTextBox
            // 
            this.LogTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.LogTextBox.Location = new System.Drawing.Point(3, 16);
            this.LogTextBox.Multiline = true;
            this.LogTextBox.Name = "LogTextBox";
            this.LogTextBox.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.LogTextBox.Size = new System.Drawing.Size(770, 140);
            this.LogTextBox.TabIndex = 0;
            // 
            // SubProgressBar
            // 
            this.SubProgressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.SubProgressBar.Location = new System.Drawing.Point(12, 72);
            this.SubProgressBar.Name = "SubProgressBar";
            this.SubProgressBar.Size = new System.Drawing.Size(776, 10);
            this.SubProgressBar.TabIndex = 1;
            // 
            // SubLabel
            // 
            this.SubLabel.AutoSize = true;
            this.SubLabel.Location = new System.Drawing.Point(12, 52);
            this.SubLabel.Name = "SubLabel";
            this.SubLabel.Size = new System.Drawing.Size(108, 13);
            this.SubLabel.TabIndex = 2;
            this.SubLabel.Text = "Выполнение задачи";
            // 
            // MainLabel
            // 
            this.MainLabel.AutoSize = true;
            this.MainLabel.Location = new System.Drawing.Point(12, 12);
            this.MainLabel.Name = "MainLabel";
            this.MainLabel.Size = new System.Drawing.Size(92, 13);
            this.MainLabel.TabIndex = 4;
            this.MainLabel.Text = "Общий прогресс";
            // 
            // MainProgressBar
            // 
            this.MainProgressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MainProgressBar.Location = new System.Drawing.Point(12, 32);
            this.MainProgressBar.Name = "MainProgressBar";
            this.MainProgressBar.Size = new System.Drawing.Size(776, 10);
            this.MainProgressBar.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Location = new System.Drawing.Point(12, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(776, 1);
            this.label3.TabIndex = 5;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Location = new System.Drawing.Point(12, 85);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(776, 1);
            this.label4.TabIndex = 6;
            // 
            // LogGroupBox
            // 
            this.LogGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.LogGroupBox.Controls.Add(this.LogTextBox);
            this.LogGroupBox.Location = new System.Drawing.Point(12, 91);
            this.LogGroupBox.Name = "LogGroupBox";
            this.LogGroupBox.Size = new System.Drawing.Size(776, 159);
            this.LogGroupBox.TabIndex = 7;
            this.LogGroupBox.TabStop = false;
            this.LogGroupBox.Text = "Процесс";
            // 
            // OpenLogFolderButton
            // 
            this.OpenLogFolderButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.OpenLogFolderButton.Location = new System.Drawing.Point(537, 258);
            this.OpenLogFolderButton.Name = "OpenLogFolderButton";
            this.OpenLogFolderButton.Size = new System.Drawing.Size(136, 23);
            this.OpenLogFolderButton.TabIndex = 8;
            this.OpenLogFolderButton.Text = "Открыть папку логов";
            this.OpenLogFolderButton.UseVisualStyleBackColor = true;
            this.OpenLogFolderButton.Visible = false;
            this.OpenLogFolderButton.Click += new System.EventHandler(this.OpenLogFolderButton_Click);
            // 
            // AbortButton
            // 
            this.AbortButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.AbortButton.Location = new System.Drawing.Point(710, 258);
            this.AbortButton.Name = "AbortButton";
            this.AbortButton.Size = new System.Drawing.Size(75, 23);
            this.AbortButton.TabIndex = 9;
            this.AbortButton.Text = "Прервать";
            this.AbortButton.UseVisualStyleBackColor = true;
            this.AbortButton.Click += new System.EventHandler(this.AbortButton_Click);
            // 
            // ProgressBarWithLog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 289);
            this.Controls.Add(this.AbortButton);
            this.Controls.Add(this.OpenLogFolderButton);
            this.Controls.Add(this.LogGroupBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.MainLabel);
            this.Controls.Add(this.MainProgressBar);
            this.Controls.Add(this.SubLabel);
            this.Controls.Add(this.SubProgressBar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ProgressBarWithLog";
            this.Text = "СПЕКТРУМ прогресс выполнения";
            this.LogGroupBox.ResumeLayout(false);
            this.LogGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox LogTextBox;
        private System.Windows.Forms.ProgressBar SubProgressBar;
        private System.Windows.Forms.Label SubLabel;
        private System.Windows.Forms.Label MainLabel;
        private System.Windows.Forms.ProgressBar MainProgressBar;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox LogGroupBox;
        private System.Windows.Forms.Button OpenLogFolderButton;
        private System.Windows.Forms.Button AbortButton;
    }
}