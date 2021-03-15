using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ACO.Offers
{
    public partial class FormSelectOfferSettings : Form
    {
        List<OfferSettings> _offerSettings;

        public FormSelectOfferSettings()
        {
            InitializeComponent();
            _offerSettings = new List<OfferSettings>();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void FormSelectOfferSettings_Load(object sender, EventArgs e)
        {

        }

        private void BtnOK_Click(object sender, EventArgs e)
        {

        }
    }
}
