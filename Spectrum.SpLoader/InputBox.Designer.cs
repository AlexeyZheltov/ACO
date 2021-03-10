
namespace Spectrum.SpLoader
{
    partial class InputBox
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
            this.CaptionLabel = new System.Windows.Forms.Label();
            this.ValueTextBox = new System.Windows.Forms.TextBox();
            this.OkButton = new System.Windows.Forms.Button();
            this.Cancel_Button = new System.Windows.Forms.Button();
            this.errorProvider = new System.Windows.Forms.ErrorProvider(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).BeginInit();
            this.SuspendLayout();
            // 
            // CaptionLabel
            // 
            this.CaptionLabel.AutoSize = true;
            this.CaptionLabel.Location = new System.Drawing.Point(13, 13);
            this.CaptionLabel.Name = "CaptionLabel";
            this.CaptionLabel.Size = new System.Drawing.Size(43, 13);
            this.CaptionLabel.TabIndex = 0;
            this.CaptionLabel.Text = "Caption";
            // 
            // ValueTextBox
            // 
            this.ValueTextBox.Location = new System.Drawing.Point(16, 30);
            this.ValueTextBox.Name = "ValueTextBox";
            this.ValueTextBox.Size = new System.Drawing.Size(378, 20);
            this.ValueTextBox.TabIndex = 1;
            // 
            // OkButton
            // 
            this.OkButton.Location = new System.Drawing.Point(238, 57);
            this.OkButton.Name = "OkButton";
            this.OkButton.Size = new System.Drawing.Size(75, 23);
            this.OkButton.TabIndex = 2;
            this.OkButton.Text = "Ok";
            this.OkButton.UseVisualStyleBackColor = true;
            this.OkButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // Cancel_Button
            // 
            this.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Cancel_Button.Location = new System.Drawing.Point(319, 57);
            this.Cancel_Button.Name = "Cancel_Button";
            this.Cancel_Button.Size = new System.Drawing.Size(75, 23);
            this.Cancel_Button.TabIndex = 3;
            this.Cancel_Button.Text = "Cancel";
            this.Cancel_Button.UseVisualStyleBackColor = true;
            // 
            // errorProvider
            // 
            this.errorProvider.BlinkStyle = System.Windows.Forms.ErrorBlinkStyle.NeverBlink;
            this.errorProvider.ContainerControl = this;
            // 
            // InputBox
            // 
            this.AcceptButton = this.OkButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.Cancel_Button;
            this.ClientSize = new System.Drawing.Size(417, 88);
            this.Controls.Add(this.Cancel_Button);
            this.Controls.Add(this.OkButton);
            this.Controls.Add(this.ValueTextBox);
            this.Controls.Add(this.CaptionLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "InputBox";
            this.Text = "InputBox";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.InputBox_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label CaptionLabel;
        private System.Windows.Forms.TextBox ValueTextBox;
        private System.Windows.Forms.Button OkButton;
        private System.Windows.Forms.Button Cancel_Button;
        private System.Windows.Forms.ErrorProvider errorProvider;
    }
}