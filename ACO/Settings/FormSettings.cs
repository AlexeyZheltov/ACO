using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ACO.Settings
{
    public partial class FormSettings : Form
    {
        readonly Properties.Settings settings = Properties.Settings.Default;
        public FormSettings()
        {
            InitializeComponent();
        }

        private void FormSettings_Load(object sender, EventArgs e)
        {
            TBoxTamplate.Text = settings.TamplateProgectPath;
            ValidatePath(TBoxTamplate);
        }

        private void BtnSetTamplatePath_Click(object sender, EventArgs e)
        {
            TBoxTamplate.Text = GetFiles();
            ValidatePath(TBoxTamplate);
        }
        private string GetFiles()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel|*.xl*|All files|*.*";
            openFileDialog.Title = "Выберите файл";
            openFileDialog.Multiselect = false;
            string file = default;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog.FileName;
            }
            return file;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            settings.TamplateProgectPath = TBoxTamplate.Text;
            settings.Save();
        }

        private void FormSettings_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) DialogResult = DialogResult.Cancel;
        }
        private void ValidatePath(TextBox textBox)
        {
            if (textBox.Text == "") errorProvider.SetError(textBox, "Файл не выбран");
            else if (!File.Exists(textBox.Text)) errorProvider.SetError(textBox, "Указанный файл не существует");
            else errorProvider.SetError(textBox, "");
        }

        private void BtnClose_Click(object sender, EventArgs e)
        {

        }
    }
}
