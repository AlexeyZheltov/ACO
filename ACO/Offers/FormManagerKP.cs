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
        List<OfferMapping> _offersMapping;
        OfferMapping _CurrentMapping;
        OfferManager _manager;
        public FormManagerKP()
        {
            InitializeComponent();
            TableColumns.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            TableColumns.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            TableColumns.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            TableColumns.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            TableColumns.Columns[4].Visible = false;
            TableColumns.Columns[5].Visible = false;
            TableColumns.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            _manager = new OfferManager();
            _offersMapping = _manager.OffersMapping;
            ListKP.FullRowSelect = true;
            ListKP.MultiSelect = false;
        }



        private void BtnAddColumns_Click(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Range rng = app.Selection;

            if ((rng?.Cells?.Count ?? 0) == 0 || rng == null) return;
            if (_mappingColumnsKP == null) _mappingColumnsKP = new List<ColumnMapping>();
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
            if (_mappingColumnsKP is null) return;
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _mappingColumnsKP.Count; i++)
            {
                Source.Add(_mappingColumnsKP[i]);
            };
            TableColumns.DataSource = Source;
        }

       private void Save()
        {
            if (_mappingColumnsKP != null)
            {
                _CurrentMapping.Columns = _mappingColumnsKP;
                _CurrentMapping.Save();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            Save();
            Close();
        }

        private void FormManagerKP_Load(object sender, EventArgs e)
        {
            LoadOffersMapping();
            UpdateTable();
        }

        private void LoadOffersMapping()
        {
            ListKP.Items.Clear();
            if (_offersMapping != null)
            {
                foreach (OfferMapping offer in _offersMapping)
                {
                    ListViewItem itm = new ListViewItem(offer.Name);
                    itm.SubItems.Add(offer.FileName);
                    ListKP.Items.Add(itm);
                }
                ListKP.Columns[0].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
                ListKP.Columns[1].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
            }
        }

        private void BtnCreate_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            OfferMapping.Create(name);
            LoadOffersMapping();
        }

        private void ListKP_SelectedIndexChanged(object sender, EventArgs e)
        {
            BindingSource source = new BindingSource();
            if (_offersMapping == null) return;
            _CurrentMapping = _offersMapping.First();
            if (ListKP.SelectedItems.Count > 0)
            {
                string name = ListKP.SelectedItems[0].SubItems[0].Text;
                if (!string.IsNullOrEmpty(name))
                {
                    OfferMapping mapping = _offersMapping.Find(X => X.Name == name);
                    if (mapping != null)
                    {
                        _CurrentMapping = mapping;
                        source.DataSource = mapping.Columns;
                    }
                }
            }
            TableColumns.DataSource = source;
        }

        private void TableColumns_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            string address = TableColumns.Rows[row].Cells[3].Value?.ToString() ?? "";
            ColumnMapping mapping = _CurrentMapping.Columns.Find(f => f.Address == address);
            if (mapping is null) return;
            object value = null;
            switch (col)
            {
                case 1:
                    value = TableColumns.Rows[row].Cells[3].Value;
                    mapping.Check = (bool)value;
                    break;
                case 2:
                    value = TableColumns.Rows[row].Cells[3].Value;
                    mapping.Obligatory = (bool)value;
                    break;
                case 3:
                    mapping.Address = address;
                    break;
                    //default:
                    //    break;
            }
            Save();
        }
    }
}
