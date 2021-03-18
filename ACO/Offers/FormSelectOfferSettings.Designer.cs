
namespace ACO.Offers
{
    partial class FormSelectOfferSettings
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
            this.listBoxOffers = new System.Windows.Forms.ListBox();
            this.BtnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listBoxOffers
            // 
            this.listBoxOffers.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.listBoxOffers.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listBoxOffers.FormattingEnabled = true;
            this.listBoxOffers.Location = new System.Drawing.Point(2, 17);
            this.listBoxOffers.Name = "listBoxOffers";
            this.listBoxOffers.Size = new System.Drawing.Size(245, 171);
            this.listBoxOffers.TabIndex = 0;
            this.listBoxOffers.SelectedIndexChanged += new System.EventHandler(this.listBoxOffers_SelectedIndexChanged);
            // 
            // BtnOK
            // 
            this.BtnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.BtnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.BtnOK.Location = new System.Drawing.Point(84, 190);
            this.BtnOK.Name = "BtnOK";
            this.BtnOK.Size = new System.Drawing.Size(84, 23);
            this.BtnOK.TabIndex = 1;
            this.BtnOK.Text = "ОК";
            this.BtnOK.UseVisualStyleBackColor = true;
            this.BtnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // FormSelectOfferSettings
            // 
            this.AcceptButton = this.BtnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(248, 215);
            this.Controls.Add(this.BtnOK);
            this.Controls.Add(this.listBoxOffers);
            this.KeyPreview = true;
            this.Name = "FormSelectOfferSettings";
            this.ShowIcon = false;
            this.Text = "Выбрать настроки КП";
            this.Load += new System.EventHandler(this.FormSelectOfferSettings_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox listBoxOffers;
        private System.Windows.Forms.Button BtnOK;
    }
}