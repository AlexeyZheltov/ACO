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
        List<OfferSettings> _Mappings;
        // public  OfferSettings Mapping { get; set; }
        // string nameOfferSetting = "";

        public FormSelectOfferSettings()
        {
            InitializeComponent();
            _Mappings = new OfferManager().Mappings;
        }


        private void FormSelectOfferSettings_Load(object sender, EventArgs e)
        {
            // BindingSource source = new BindingSource();
            for (int i = 0; i < _Mappings.Count; i++)
            {
                listBoxOffers.Items.Add(_Mappings[i].Name);
            }
            //   source.Add(_offerSettings[i].Name);
            // listBoxOffers.DataSource = source;            
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            //    if (Mapping != null)
            //{
            //    DialogResult = DialogResult.OK;
            //}
        }

        private void listBoxOffers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxOffers.SelectedItem != null)
            {
                OfferSettingsName = listBoxOffers.SelectedItem.ToString();
               
                // Mapping = _offerSettings.Find(x => x.Name == name);              
            }
        }
    }
}
