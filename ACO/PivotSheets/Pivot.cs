using ACO.ExcelHelpers;
using ACO.ProjectBook;
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
            Excel.Workbook wb = _app.ActiveWorkbook;
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
            List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;// GetAdderss();
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
                    Excel.Range numberCell = _SheetUrv12.Cells[rowPaste, 2];
                    numberCell.NumberFormat = "@";
                    numberCell.Value = number;
                    _SheetUrv12.Cells[rowPaste, 3].Value = name;
                    _SheetUrv12.Cells[rowPaste, 4].Value = cost;

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
                    rowPaste++;
                }
            }
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range rng = ws.Cells[13, 1];
            rng.EntireRow.Delete();           
        }

        //private List<OfferAddress> GetAdderss()
        //{
        //    List<OfferAddress> addresses = new List<OfferAddress>();
        //    int lastCol = _AnalisysSheet.Cells[1, _AnalisysSheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
        //    int columnStart = 0;
        //    int columnTotal = 0;
        //   // string name = "";

        //    for (int col = 1; col <= lastCol; col++)
        //    {
        //        string val = _AnalisysSheet.Cells[1, col].Value?.ToString() ?? "";
        //        if (val == "offer_start")
        //        {
        //            columnStart = col;
        //        //    name = _AnalisysSheet.Cells[6, col].Value?.ToString() ?? "";
        //        }
        //        if (val == Project.ColumnsNames[StaticColumns.CostTotal])
        //        {
        //            columnTotal = col;
        //        }
        //        if (val == "offer_end")
        //        {
        //            OfferAddress address = new OfferAddress
        //            {
        //                ColStartOffer = columnStart,
        //                ColStartOfferComments = col,
        //                ColTotalCost = columnTotal,
        //                ColPercentTotal = col + 4,
        //                ColPercentMaterial = col + 6,
        //                ColPercentWorks = col + 7,
        //                ColComments = col + 8
        //            };
        //            addresses.Add(address);
        //        }
        //    }
        //    return addresses;
        //}




        public void LoadUrv11(IProgressBarWithLogUI pb)
        {
            List<OfferAddress> addresses = GetAdderssLvl12();
            int lastRow = _SheetUrv12.UsedRange.Row + _SheetUrv12.UsedRange.Rows.Count - 1;
            int rowPaste = 14;
            int colPaste = 6;
            for (int row = 13; row <= lastRow; row++)
            {
                string number = _SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                number = number.Trim(new char[] { ' ', '.' });
                if (string.IsNullOrEmpty(number)) break;

                string name = _SheetUrv12.Cells[row, 3].Value?.ToString() ?? "";
                string cost = _SheetUrv12.Range[row, 4].Value?.ToString() ?? "";
                //string level = number.Split(".").Length;
                int levelNum = number.Split('.').Length;  //int levelNum = int.TryParse(level, out int ln) ? ln : 0;
                if (levelNum > 0 && levelNum < 3)
                {
                    _SheetUrv11.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    Excel.Range numberCell = _SheetUrv11.Cells[rowPaste, 2];
                    numberCell.NumberFormat = "@";
                    numberCell.Value = number;
                   
                    _SheetUrv11.Cells[rowPaste, 3].Value = name;
                    _SheetUrv11.Cells[rowPaste, 7].Value = cost;

                    foreach (OfferAddress address in addresses)
                    {
                        _SheetUrv11.Cells[rowPaste, colPaste].Value = _SheetUrv12.Cells[row, address.ColTotalCost].Value?.ToString() ?? "";
                        _SheetUrv11.Cells[rowPaste, colPaste + 3].Value = _SheetUrv12.Cells[row, address.ColPercentTotal].Value?.ToString() ?? "";
                        colPaste += 5;
                    }
                    colPaste = 6;
                    rowPaste++;
                }
            }
        }

        /// <summary>
        ///  Номера столбцов заполненных кп
        /// </summary>
        /// <returns></returns>
        private List<OfferAddress> GetAdderssLvl12()
        {
            List<OfferAddress> addresses = new List<OfferAddress>();
            int lastCol = _SheetUrv12.Cells[13, _SheetUrv12.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

            for (int col = 9; col <= lastCol; col += 3)
            {
                string val = _SheetUrv12.Cells[13, col].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(val))
                {
                    OfferAddress address = new OfferAddress
                    {
                        ColTotalCost = col,
                        ColPercentTotal = col + 1,
                        ColComments = col + 2
                    };
                    addresses.Add(address);
                }
            }
            return addresses;
        }

        /// <summary>
        ///  Обновление значений урв 12
        /// </summary>
        /// <param name="pb"></param>
        internal void UpdateUrv12(IProgressBarWithLogUI pb)
        {
            string letterNumber = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            string letterLevel = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;

            int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
            List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;
            int rowPaste = 14;
            int colPaste = 6;
            for (int row = _project.RowStart; row <= lastRow; row++)
            {
                string number = _AnalisysSheet.Range[$"${letterNumber}{row}"].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;
                if (_SheetUrv12.Cells[rowPaste, 2].Value == number)
                {
                    string level = _AnalisysSheet.Range[$"${letterLevel}{row}"].Value?.ToString() ?? "";
                    int levelNum = int.TryParse(level, out int ln) ? ln : 0;

                    if (levelNum > 0 && levelNum < 6)
                    {
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
                        rowPaste++;
                    }
                }
            }
        }

    }
}
