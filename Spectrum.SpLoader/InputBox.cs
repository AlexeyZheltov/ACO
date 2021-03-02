using System;
using System.Windows.Forms;

namespace Spectrum.SpLoader
{
    public partial class InputBox : Form
    {
        public string Value { get; set; }

        public string Title
        {
            get => Text;
            set => Text = value;
        }

        public string Caption
        {
            get => CaptionLabel.Text;
            set => CaptionLabel.Text = value;
        }

        public InputBox() => InitializeComponent();

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(ValueTextBox.Text))
                errorProvider.SetError(ValueTextBox, "Поле не может быть пустым");
            else
            {
                Value = ValueTextBox.Text;
                DialogResult = DialogResult.OK;
            }
        }

        private void InputBox_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
                DialogResult = DialogResult.Cancel;

            errorProvider.Dispose();
        }
    }
}
