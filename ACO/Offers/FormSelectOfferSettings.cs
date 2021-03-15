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
        public string OfferSettingsName { get; set; }
        List<OfferSettings> _offerSettings;
          OfferSettings Settings { get; set; }

        public FormSelectOfferSettings()
        {
            InitializeComponent();
            ///_offerSettings = new List<OfferSettings>();
            _offerSettings = new OfferManager().Mappings;
        }


        private void FormSelectOfferSettings_Load(object sender, EventArgs e)
        {
            BindingSource source = new BindingSource();
            for (int i = 0; i < _offerSettings.Count; i++)
            {
                source.Add(_offerSettings[i].Name);
            }
            listBoxOffers.DataSource = source;            
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
                if (Settings != null)
            {
                DialogResult = DialogResult.OK;
            }
        }

        private void listBoxOffers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxOffers.SelectedItem != null)
            {
                string name = listBoxOffers.SelectedItem.ToString();
                Settings = _offerSettings.Find(x => x.Name == name);              
            }
        }
    }
}
