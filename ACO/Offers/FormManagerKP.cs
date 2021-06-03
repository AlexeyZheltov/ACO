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
        private readonly Excel.Application _app = Globals.ThisAddIn.Application;
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

            //  ListKP.FullRowSelect = true;
            // ListKP.MultiSelect = false;
            TableColumns.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

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

            VewActiveOfferSettings(_CurrentMapping);
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

        /// <summary>
        ///  Показать активные настройки КП в заголовке формы.
        /// </summary>
        /// <param name="project"></param>
        private void VewActiveOfferSettings(OfferSettings offerSettings)
        {
            this.Text = $"Диспетчер КП [{offerSettings.Name}]";
        }

        private void TableColumns_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridViewTextBoxEditingControl tb = (DataGridViewTextBoxEditingControl)e.Control;
            tb.KeyPress += new KeyPressEventHandler(dataGridViewTextBox_KeyPress);
            e.Control.KeyPress += new KeyPressEventHandler(dataGridViewTextBox_KeyPress);

        }

        static readonly char[] _allowLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
        private void dataGridViewTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char keyChar = e.KeyChar;
            if (Char.IsControl(keyChar))
                return;

            keyChar = Char.ToUpper(keyChar);

            if ((sender as TextBox).TextLength == 3 || !_allowLetters.Contains(keyChar))
            {
                e.Handled = true;
                return;
            }
            e.KeyChar = keyChar;
        }


        #region Cut\copy Datagrid
        private void ListKP_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                DataObject d = TableColumns.GetClipboardContent();
                Clipboard.SetDataObject(d);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                string s = Clipboard.GetText();
                string[] lines = s.Split('\n');
                int row = TableColumns.CurrentCell.RowIndex;
                int col = TableColumns.CurrentCell.ColumnIndex;
                string[] cells = lines[0].Split('\t');
                int cellsSelected = cells.Length;
                for (int i = 0; i < cellsSelected; i++)
                {
                    TableColumns[col, row].Value = cells[i];
                    col++;
                }
            }
            else if (e.Control && e.KeyCode == Keys.X)
            {
                CopyToClipboard(); //Copy to clipboard
                                   //Clear selected cells
                ClearSelection();
            }
            else if (e.KeyCode == Keys.Delete)
            {
                ClearSelection();
            }

            }


        private void TableColumns_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {            
            if (TableColumns.SelectedCells.Count > 0)
                TableColumns.ContextMenuStrip = contextMenuStrip1;
        }

        private void копироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyToClipboard();
        }

        private void вырезатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CopyToClipboard(); //Copy to clipboard
            ClearSelection();
        }
        private void ClearSelection()
        {
            foreach (DataGridViewCell dgvCell in TableColumns.SelectedCells)
                dgvCell.Value = string.Empty;
        }

        private void вставитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
             //Perform paste Operation
            PasteClipboardValue();
        }
                    
        private void CopyToClipboard()
        {
            //Copy to clipboard
            DataObject dataObj = TableColumns.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }


        private void PasteClipboardValue()
        {
            //Show Error if no cell is selected
            if (TableColumns.SelectedCells.Count == 0)
            {
                MessageBox.Show("Выделите ячейку", "Вставка",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //Get the starting Cell
            DataGridViewCell startCell = GetStartCell(TableColumns);
            //Get the clipboard value in a dictionary
            Dictionary<int, Dictionary<int, string>> cbValue =
                    ClipBoardValues(Clipboard.GetText());

            int iRowIndex = startCell.RowIndex;
            foreach (int rowKey in cbValue.Keys)
            {
                int iColIndex = startCell.ColumnIndex;
                foreach (int cellKey in cbValue[rowKey].Keys)
                {
                    //Check if the index is within the limit
                    if (iColIndex <= TableColumns.Columns.Count - 1
                    && iRowIndex <= TableColumns.Rows.Count - 1)
                    {
                        DataGridViewCell cell = TableColumns[iColIndex, iRowIndex];

                        //Copy to selected cells if 'chkPasteToSelectedCells' is checked
                          cell.Value = cbValue[rowKey][cellKey];
                    }
                    iColIndex++;
                }
                iRowIndex++;
            }
        }

        private DataGridViewCell GetStartCell(DataGridView dgView)
        {
            //get the smallest row,column index
            if (dgView.SelectedCells.Count == 0)
                return null;

            int rowIndex = dgView.Rows.Count - 1;
            int colIndex = dgView.Columns.Count - 1;

            foreach (DataGridViewCell dgvCell in dgView.SelectedCells)
            {
                if (dgvCell.RowIndex < rowIndex)
                    rowIndex = dgvCell.RowIndex;
                if (dgvCell.ColumnIndex < colIndex)
                    colIndex = dgvCell.ColumnIndex;
            }
            return dgView[colIndex, rowIndex];
        }

        private Dictionary<int, Dictionary<int, string>> ClipBoardValues(string clipboardValue)
        {
            Dictionary<int, Dictionary<int, string>>
            copyValues = new Dictionary<int, Dictionary<int, string>>();

            String[] lines = clipboardValue.Split('\n');

            for (int i = 0; i <= lines.Length - 1; i++)
            {
                copyValues[i] = new Dictionary<int, string>();
                String[] lineContent = lines[i].Split('\t');

                //if an empty cell value copied, then set the dictionary with an empty string
                //else Set value to dictionary
                if (lineContent.Length == 0)
                    copyValues[i][0] = string.Empty;
                else
                {
                    for (int j = 0; j <= lineContent.Length - 1; j++)
                        copyValues[i][j] = lineContent[j];
                }
            }
            return copyValues;
        }
        #endregion
        
    }
}
