using System;
using System.Collections.Generic;
using System.ComponentModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
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

            _projectManager = new ProjectManager.ProjectManager();
            _manager = new OfferManager();
            _mappingColumnsOffer = new List<OfferColumnMapping>();

            ListKP.FullRowSelect = true;
            ListKP.MultiSelect = false;
            ListKP.View = View.List;
            LoadData();
        }

        private void LoadTableOffers()
        {
            ListKP.Items.Clear();
            if (_offersMapping != null)
            {
                foreach (OfferSettings offer in _offersMapping)
                {
                    ListKP.Items.Add(offer.Name);
                }
            }
        }

        private void GetOfferMappings()
        {
            _offersMapping = _manager.Mappings;
            if ((_offersMapping?.Count ?? 0) > 0)
            {
                _CurrentMapping = _offersMapping.First();
            }
        }

        private void LoadData()
        {
            GetOfferMappings();
            LoadTableOffers();
            SetTableMapping();
        }

        private void ListKP_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ListKP.SelectedItems.Count > 0)
            {
                string nameSettings = ListKP.SelectedItems[0].Text;
                _CurrentMapping = _manager.Mappings.Find(x => x.Name == nameSettings);
                SetTableMapping();
            }
        }
        private void BtnCreate_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            OfferSettings.Create(name);
            _manager.UpdateMappings();
            _offersMapping = _manager.Mappings;
            _CurrentMapping = _offersMapping.Find(m => m.Name == name);
            LoadTableOffers();
            SetTableMapping();
        }
        private void BtnCopoySettings_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            if (string.IsNullOrWhiteSpace(name)) { return; }
            OfferSettings.Copy(_CurrentMapping, name);
            _manager.UpdateMappings();
            _offersMapping = _manager.Mappings;
            _CurrentMapping = _offersMapping.Find(m => m.Name == name);
            LoadTableOffers();
            SetTableMapping();
        }

        private void SetTableMapping()
        {
            if ((_projectManager.ActiveProject?.Columns.Count ?? 0) == 0) { return; }
            _mappingColumnsOffer.Clear();

            foreach (ColumnMapping col in _projectManager.ActiveProject.Columns)
            {
                OfferColumnMapping columnMapping = _CurrentMapping.Columns.Find(x => x.Name == col.Name);

                OfferColumnMapping cm = new OfferColumnMapping();
                if (columnMapping != null)
                {
                    cm.ColumnSymbol = columnMapping.ColumnSymbol;
                }
                cm.Name = col.Name;
                _mappingColumnsOffer.Add(cm);
            }
            UpdateTableSource();

            if (_CurrentMapping != null)
            {
                TBoxSheetName.Text = _CurrentMapping.SheetName;
                TBoxFirstRowRangeValues.Text = _CurrentMapping.RowStart.ToString();
            }
            SetTableColumns();
        }

        private void UpdateTableSource()
        {
            if ((_mappingColumnsOffer?.Count ?? 0) == 0) return;
            BindingSource Source = new BindingSource();
            TableColumns.Rows.Clear();
            for (int i = 0; i < _mappingColumnsOffer.Count; i++)
            {
                Source.Add(_mappingColumnsOffer[i]);
            };
            TableColumns.DataSource = Source;
        }
        /// <summary>
        ///  Таблица с ячейками 
        /// </summary>
        private void SetTableColumns()
        {
            if (TableColumns.Columns.Count < 2) return;

            TableColumns.Columns[1].Width = 70;
            TableColumns.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            TableColumns.Columns[0].HeaderText = "Cтолбец";
            TableColumns.Columns[1].HeaderText = "Адрес";
            TableColumns.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        }

        private void Save()
        {
            if (_CurrentMapping == null) return;
            _CurrentMapping.SheetName = TBoxSheetName.Text;
            _CurrentMapping.RowStart = int.TryParse(TBoxFirstRowRangeValues.Text, out int rs) ? rs : 0;

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
                string name = row.Cells[0].Value?.ToString();
                OfferColumnMapping mapping = _mappingColumnsOffer.Find(x => x.Name == name);
                TableColumns.Rows.Remove(row);
                if (mapping != null) _mappingColumnsOffer.Remove(mapping);
            }
        }

        private void TableColumns_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!(e.RowIndex >= 0 && e.RowIndex >= 0)) return;
            string name = TableColumns.Rows[e.RowIndex].Cells[0].Value?.ToString() ?? "";
            OfferColumnMapping mapping = _mappingColumnsOffer.Find(f => f.Name == name);

            if (e.ColumnIndex == 1)
            {
                object value = TableColumns.Rows[e.RowIndex].Cells[1].Value;
                mapping.ColumnSymbol = value?.ToString() ?? "";
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
            int rowStart = rng.Row + rng.Rows.Count;
            TBoxFirstRowRangeValues.Text = rowStart.ToString();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

      
    }
}
