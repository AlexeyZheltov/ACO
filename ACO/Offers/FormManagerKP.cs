using System;
using System.Collections.Generic;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ACO.ProjectManager;

namespace ACO.Offers
{
    public partial class FormManagerKP : Form
    {
        private List<ColumnMapping> _mappingColumnsKP;

        public FormManagerKP()
        {
            InitializeComponent();
        }

   

        private void BtnAddColumns_Click(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.Selection;
            if ((rng?.Cells?.Count ?? 0) == 0) return;
            foreach (Excel.Range cell in rng.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Text))
                {
                    ColumnMapping mapping = new ColumnMapping(cell);
                    ColumnMapping findMapping = _mappingColumnsKP.Find(m => m.Address == mapping.Address);
                    if (findMapping != null)
                    {
                        _mappingColumnsKP.Remove(findMapping);
                    }
                    _mappingColumnsKP.Add(mapping);
                }
            }
            UpdateTable();
        }

        private void UpdateTable()
        {
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _mappingColumnsKP.Count; i++)
            {
                Source.Add(_mappingColumnsKP[i]);
            };
            TableColumns.DataSource = Source;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {

        }

        private void FormManagerKP_Load(object sender, EventArgs e)
        {
            LoadOffersMapping();
            //ListKP.Items.Add(new ListViewItem())
           
        }

        private void LoadOffersMapping()
        {
            OfferManager manager = new OfferManager();
            List<OfferMapping>  OffersMapping = manager.OffersMapping;

        }

        private void BtnCreate_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            OfferMapping.Create(name);
        }
    }
}
