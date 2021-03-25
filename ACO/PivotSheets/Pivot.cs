using ACO.ExcelHelpers;
using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.PivotSheets
{
    /// <summary>
    ///  Загрузка и обновление сводных таблиц
    /// </summary>
    class Pivot
    {
        Excel.Application _app = Globals.ThisAddIn.Application;
        Excel.Worksheet _SheetUrv12;
        Excel.Worksheet _SheetUrv11;
        Excel.Worksheet _AnalisysSheet;
        ProjectManager.ProjectManager _projectManager;
        Project _project;

        public Pivot()
        {
            // _app = Globals.ThisAddIn.Application;
            Excel.Workbook wb = _app.ActiveWorkbook;
            //string sheetName = "Урв12";
            _SheetUrv12 = ExcelHelper.GetSheet(wb, "Урв12");
            _SheetUrv11 = ExcelHelper.GetSheet(wb, "Урв11");
            _projectManager = new ProjectManager.ProjectManager();
            _project = _projectManager.ActiveProject;
            string analisysSheetName = _project.AnalysisSheetName;
            _AnalisysSheet = ExcelHelper.GetSheet(wb, analisysSheetName);
        }

        /// Добавить список
        /// Добавить столбцы КП
        /// Проставить формулы
        public void LoadUrv12(IProgressBarWithLogUI pb)
        {
            string letterName = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Name]).ColumnSymbol;
            string letterNumber = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            string letterLevel = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            string letterCost = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol;

            int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
            //foreach (Excel.Range row in _AnalisysSheet.UsedRange.Rows)
            // Dictionary<string, OfferAddress> addresses = GetAdderss();
            List<OfferAddress> addresses = GetAdderss();
            int rowPaste = 14;
            int colPaste = 6;
            for (int row = _project.RowStart; row <= lastRow; row++)
            {
                string number = _AnalisysSheet.Range[$"${letterNumber}{row}"].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;

                string name = _AnalisysSheet.Range[$"${letterName}{row}"].Value?.ToString() ?? "";
                string level = _AnalisysSheet.Range[$"${letterLevel}{row}"].Value?.ToString() ?? "";
                string cost = _AnalisysSheet.Range[$"${letterCost}{row}"].Value?.ToString() ?? "";
                int levelNum = int.TryParse(level, out int ln) ? ln : 0;

                if (levelNum > 0 && levelNum < 6)
                {
                    _SheetUrv12.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    _SheetUrv12.Cells[rowPaste, 2].Value = number;
                    _SheetUrv12.Cells[rowPaste, 3].Value = name;

                    foreach (OfferAddress address in addresses)
                    {
                        _SheetUrv12.Cells[rowPaste, colPaste].Value = _AnalisysSheet.Cells[row, address.ColTotalCost].Value?.ToString() ?? "";
                        _SheetUrv12.Cells[rowPaste, colPaste + 1].Value = _AnalisysSheet.Cells[row, address.ColPercentMaterial].Value?.ToString() ?? "";
                        _SheetUrv12.Cells[rowPaste, colPaste + 2].Value = _AnalisysSheet.Cells[row, address.ColPercentWorks].Value?.ToString() ?? "";
                        _SheetUrv12.Cells[rowPaste, colPaste + 3].Value = _AnalisysSheet.Cells[row, address.ColPercentTotal].Value?.ToString() ?? "";
                        _SheetUrv12.Cells[rowPaste, colPaste + 4].Value = _AnalisysSheet.Cells[row, address.ColComments].Value?.ToString() ?? "";
                        colPaste += 5;
                    }
                    colPaste = 6;
                    // PrintOffers(rowPaste, row);
                    rowPaste++;
                }
            }
        }

        private List<OfferAddress> GetAdderss()
        {
            List<OfferAddress> addresses = new List<OfferAddress>();
            int lastCol = _AnalisysSheet.Cells[1, _AnalisysSheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            int columnStart = 0;
            int columnTotal = 0;
            string name = "";
            //double cost =0;

            for (int col = 1; col <= lastCol; col++)
            {
                string val = _AnalisysSheet.Cells[1, col].Value?.ToString() ?? "";
                if (val == "offer_start")
                {
                    columnStart = col;
                    name = _AnalisysSheet.Cells[6, col].Value?.ToString() ?? "";
                }
                if (val == Project.ColumnsNames[StaticColumns.CostTotal])
                {
                    columnTotal = col;

                    //string costText = _AnalisysSheet.Cells[1, col].Value?.ToString() ?? "";
                    //cost = double.TryParse(costText, out double c) ? c : 0;
                }
                if (val == "offer_end")
                {
                    OfferAddress address = new OfferAddress();
                    address.ColStartOffer = columnStart;
                    address.ColStartOfferComments = col;
                    address.ColTotalCost = columnTotal;
                    address.ColPercentTotal = col + 4;
                    address.ColPercentMaterial = col + 6;
                    address.ColPercentWorks = col + 7;
                    address.ColComments = col + 8;
                    addresses.Add(address);
                    //if (!string.IsNullOrEmpty(name))
                    //{
                    //}
                }
            }
            return addresses;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="rowPaste"></param>
        /// <param name="row"></param>
        private void PrintOffers(int rowPaste, int row)
        {
            //_SheetUrv11

        }

        /// <summary>
        /// 
        /// </summary>
        class OfferAddress
        {
            public int ColPercentMaterial { get; set; }
            public int ColPercentWorks { get; set; }
            public int ColPercentTotal { get; set; }
            public int ColTotalCost { get; set; }
            public int ColComments { get; set; }
            public int ColStartOffer { get; set; }
            public int ColStartOfferComments { get; set; }
            // public string p { get; set; }
            //public string letterParticipantName { get; set; }
            //public string letterPercentMaterial { get; set; }
            //public string letterPercentWorks { get; set; }
            //public string letterPercentTotal { get; set; }
            //public string letterTotalCost { get; set; }
        }
        class OfferComments
        {
            public string ParticipantName { get; set; }
            public string PercentMaterial { get; set; }
            public string PercentWorks { get; set; }
            public string PercentTotal { get; set; }
            public double TotalCost { get; set; }
        }

        public void LoadUrv11(IProgressBarWithLogUI pb)
        {

        }
    }
}
