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
        private Excel.Application _app = Globals.ThisAddIn.Application;
        private List<OfferColumnMapping> _mappingColumnsOffer;
        List<OfferSettings> _offersMapping;
        OfferSettings _CurrentMapping;
        OfferManager _manager;
        ProjectManager.ProjectManager _projectManager;
        public FormManagerKP()
        {
            InitializeComponent();

          //  TableColumns.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            _projectManager = new ProjectManager.ProjectManager();
            _manager = new OfferManager();
            _mappingColumnsOffer = new List<OfferColumnMapping>();

            ListKP.FullRowSelect = true;
            ListKP.MultiSelect = false;
            ListKP.View = View.List;           
        }
        private void FormManagerKP_Load(object sender, EventArgs e)
        {           
            LoadData();
        }

        /// <summary>
        ///  Таблица с ячейками 
        /// </summary>
        private void SetTableColumns()
        {
            if (TableColumns.Columns.Count < 5) return;

            TableColumns.Columns[2].Width = 70;
            TableColumns.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            TableColumns.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            TableColumns.Columns[0].HeaderText = "Ссылка \"Анализ\"";
            TableColumns.Columns[1].HeaderText = "Наименование столбцa";
            TableColumns.Columns[2].HeaderText = "Адрес";
            TableColumns.Columns[3].Visible = false;
            TableColumns.Columns[4].Visible = false;
        }

        private void LoadOffersMapping()
        {
            ListKP.Items.Clear();
            if (_offersMapping != null)
            {
                if (_CurrentMapping == null && (_offersMapping?.Count ?? 0) > 0) _CurrentMapping = _offersMapping?.First();
                _mappingColumnsOffer = _CurrentMapping?.Columns;
                foreach (OfferSettings offer in _offersMapping)
                {
                    ListKP.Items.Add(offer.Name);
                }
                
            }
        }

        private void LoadData()
       {
            _offersMapping = _manager.GetMappings();
            _CurrentMapping = _offersMapping.First();

            LoadOffersMapping();
            UpdateTable();
            
            if (_CurrentMapping is null) return;
            TBoxSheetName.Text =  _CurrentMapping.SheetName ;
            TBoxFirstRowRangeValues.Text = _CurrentMapping.RowStart.ToString();
            TBoxFirstColumnRangeValues.Text = _CurrentMapping.RangeValuesStart.ToString();
            TBoxLastColumnRangeValues.Text = _CurrentMapping.RangeValuesEnd.ToString();
        }

        private void BtnAddColumns_Click(object sender, EventArgs e)
        {

            Excel.Range rng = _app.Selection;

            if ((rng?.Cells?.Count ?? 0) == 0 || rng == null) return;
            if (_mappingColumnsOffer == null) _mappingColumnsOffer = new List<OfferColumnMapping>();
            foreach (Excel.Range cell in rng.Cells)
            {
                string cellText = cell.Text;
                cellText = cellText.Replace("\n", "");
                if (!string.IsNullOrEmpty(cellText))
                {
                    OfferColumnMapping findMapping = null;
                    OfferColumnMapping mappingFromCell = new OfferColumnMapping(cell);

                    /// Искать соответствующий столбец в настройках проекта
                    ColumnMapping findProjectMapping =
                        _projectManager.ActiveProject.Columns.Find(c => c.Value == cellText);
                    if (findProjectMapping != null)
                    {
                        mappingFromCell.Link = cellText;

                        findMapping = _mappingColumnsOffer.Find(m => m.Link == cellText);
                        if (findMapping != null)
                        {
                            _mappingColumnsOffer.Remove(findMapping);
                        }
                    }

                    findMapping = _mappingColumnsOffer.Find(
                                   m => m.Address == mappingFromCell.Address);
                    if (findMapping != null)
                    {
                        _mappingColumnsOffer.Remove(findMapping);
                    }
                    _mappingColumnsOffer.Add(mappingFromCell);
                }
            }
            UpdateTable();
        }


        private void UpdateTable()
        {

            if ((_mappingColumnsOffer?.Count ?? 0) == 0) return;
            if ((_projectManager.ActiveProject?.Columns.Count ?? 0) > 0)
            {
                foreach (ColumnMapping col in _projectManager.ActiveProject.Columns)
                {
                    OfferColumnMapping columnMapping = _mappingColumnsOffer.Find(x => x.Link == col.Value);
                    if (col is null)
                    {
                        _mappingColumnsOffer.Add(new OfferColumnMapping() { Link = col.Value });
                    }
                }
            }
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _mappingColumnsOffer.Count; i++)
            {
                Source.Add(_mappingColumnsOffer[i]);
            };
            TableColumns.DataSource = Source;
            SetTableColumns();
        }


        private void Save()
        {
            if (_CurrentMapping == null) return;
            _CurrentMapping.SheetName = TBoxSheetName.Text;
            _CurrentMapping.RowStart = int.TryParse(TBoxFirstRowRangeValues.Text, out int rs) ? rs : 0;
            _CurrentMapping.RangeValuesStart = int.TryParse(TBoxFirstColumnRangeValues.Text, out int fr) ? fr : 0;
            _CurrentMapping.RangeValuesEnd = int.TryParse(TBoxLastColumnRangeValues.Text, out int lr) ? lr : 0;

            if (_mappingColumnsOffer != null)
            {
                _CurrentMapping.Columns = _mappingColumnsOffer;
            }
                _CurrentMapping.Save();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            Save();
            Close();
        }
      
      
        private void BtnCreate_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            OfferSettings.Create(name);
            _offersMapping = _manager.GetMappings();
            _CurrentMapping = _offersMapping.Find(m => m.Name == name);
            LoadOffersMapping();
            UpdateTable();
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
                    OfferSettings mapping = _offersMapping.Find(X => X.Name == name);
                    if (mapping != null)
                    {
                        _CurrentMapping = mapping;
                        source.DataSource = mapping.Columns;
                    }
                }
            }
            TableColumns.DataSource = source;
        }

        /// <summary>
        ///  Удалить выделенную строку
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (TableColumns.SelectedRows.Count > 0 && _mappingColumnsOffer != null)
            {
                DataGridViewRow row = TableColumns.SelectedRows[0];
                string address = row.Cells[2].Value?.ToString();
                OfferColumnMapping mapping = _mappingColumnsOffer.Find(x => x.Address == address);
                TableColumns.Rows.Remove(row);
                if (mapping != null) _mappingColumnsOffer.Remove(mapping);
            }
        }

        private void TableColumns_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            if (!(row >= 0 && col >= 0)) return;
            string address = TableColumns.Rows[row].Cells[2].Value?.ToString() ?? "";
            OfferColumnMapping mapping = _CurrentMapping.Columns.Find(f => f.Address == address);
            if (mapping is null) return;
            object value = null;
            switch (col)
            {
                case 0:
                    value = TableColumns.Rows[row].Cells[0].Value;
                    mapping.Link = value.ToString() ;
                    break;
                case 1:
                    value = TableColumns.Rows[row].Cells[1].Value;
                    mapping.Value = value.ToString();                   
                    break;                
                    //default:
                    //    break;
            }
            
        }
     

        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            string folder = OfferManager.GetFolderSettingsKP();
            System.Diagnostics.Process.Start(folder);
        }

        private void BtnSetSelectedRangeValues_Click(object sender, EventArgs e)
        {
            Excel.Range rng = _app.Selection;
            if (rng is null) return;
            TBoxSheetName.Text = rng.Parent.name;
            TBoxFirstColumnRangeValues.Text = rng.Column.ToString();
            int lastCol = rng.Column + rng.Columns.Count - 1;
            TBoxLastColumnRangeValues.Text = lastCol.ToString();
            int rowStart = rng.Row + rng.Rows.Count;
            TBoxFirstRowRangeValues.Text = rowStart.ToString();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

     
    }
}
