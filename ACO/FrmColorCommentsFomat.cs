using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ACO
{
    public partial class FrmColorCommentsFomat : Form
    {
        public FrmColorCommentsFomat()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            colorDialog.Color = richTextBox1.BackColor;
           if ( colorDialog.ShowDialog()== DialogResult.OK)
            {
             richTextBox1.BackColor = colorDialog.Color;
            }
        }
         
        private void button2_Click(object sender, EventArgs e)
        {
            fontDialog.Font = richTextBox1.Font;
            if (fontDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.Font = fontDialog.Font;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            colorDialog.Color = richTextBox1.ForeColor;
            if (colorDialog.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.ForeColor = colorDialog.Color;
            }
        }
        
    }
}
