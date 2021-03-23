using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ACO.Offers
{
    public partial class FormSelectOfferSettings : Form
    {
        public string OfferSettingsName { get; set; }
        List<OfferSettings> _Mappings;

        public FormSelectOfferSettings()
        {
            InitializeComponent();
            _Mappings = new OfferManager().Mappings;
        }


        private void FormSelectOfferSettings_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < _Mappings.Count; i++)
            {
                listBoxOffers.Items.Add(_Mappings[i].Name);
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
        }

        private void ListBoxOffers_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxOffers.SelectedItem != null)
            {
                OfferSettingsName = listBoxOffers.SelectedItem.ToString();
            }
        }

        private void ListBoxOffers_DoubleClick(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(OfferSettingsName))
            {
                DialogResult = DialogResult.OK;
                Close();
            }
        }
    }
}
