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
        //private int const rowStart= 
        Excel.Application _app = Globals.ThisAddIn.Application;
        Excel.Worksheet _SheetUrv12;
        Excel.Worksheet _SheetUrv11;
        Excel.Worksheet _SheetPalette;

        Excel.Worksheet _AnalisysSheet;
        ProjectManager.ProjectManager _projectManager;
        Project _project;

        public Pivot()
        {
            Excel.Workbook wb = _app.ActiveWorkbook;
            _SheetUrv12 = ExcelHelper.GetSheet(wb, "Урв12");
            _SheetUrv11 = ExcelHelper.GetSheet(wb, "Урв11");
            _SheetPalette = ExcelHelper.GetSheet(wb, "Палитра");
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
            pb.SetMainBarVolum(4);
            pb.MainBarTick("Очистка");
            ClearDataRng12();
            string letterName = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Name]).ColumnSymbol;
            string letterNumber = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            string letterLevel = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            string letterCost = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol;

            int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
            pb.MainBarTick("Определение столбцов КП");
            List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;
            int rowPaste = 14;
            int colPaste = 6;
            int lastCol = colPaste + 5 * addresses.Count - 1;

            int count = lastRow - _project.RowStart + 1;
            if (count < 1) throw new AddInException($"Строки отсутствуют лист: {_project.AnalysisSheetName}");
            pb.SetSubBarVolume(count);
            pb.MainBarTick("Заполнение строк");
            ProjectWorkbook projectWorkbook = new ProjectWorkbook();


            for (int row = _project.RowStart; row <= lastRow; row++)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();

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

                    // Формат строки по уровню
                    Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);
                    if (pallets.TryGetValue(level, out Excel.Range pallet))
                    {
                        pallet.Copy();

                        _SheetUrv12.Range[_SheetUrv12.Cells[rowPaste, 2], _SheetUrv12.Cells[rowPaste, lastCol]].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    }
                    // Вывод и форматирование значений
                    foreach (OfferAddress address in addresses)
                    {
                        _SheetUrv12.Cells[rowPaste, colPaste].Value = _AnalisysSheet.Cells[row, address.ColTotalCost].Value?.ToString() ?? "";
                        _SheetUrv12.Cells[rowPaste, colPaste + 1].Value = _AnalisysSheet.Cells[row, address.ColPercentMaterial].Value?.ToString() ?? "";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 1], level);
                        _SheetUrv12.Cells[rowPaste, colPaste + 2].Value = _AnalisysSheet.Cells[row, address.ColPercentWorks].Value?.ToString() ?? "";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 2], level);
                        _SheetUrv12.Cells[rowPaste, colPaste + 3].Value = _AnalisysSheet.Cells[row, address.ColPercentTotal].Value?.ToString() ?? "";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 3], level);
                        _SheetUrv12.Cells[rowPaste, colPaste + 4].Value = _AnalisysSheet.Cells[row, address.ColComments].Value?.ToString() ?? "";
                        colPaste += 5;
                    }
                    colPaste = 6;
                    rowPaste++;
                }
            }
            pb.MainBarTick("Удаление стр №13");           
            Excel.Range rng = _SheetUrv12.Cells[13, 1];
            rng.EntireRow.Delete();
        }

    
        private int GetLastRowUrv12()
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            return rowBottomTotal - 2;
        }
        private int GetLastRowUrv11()
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv11, "СУММА ПРЕДЛОЖЕННЫХ ОПТИМИЗАЦИЙ (с НДС)").Row;
            return rowBottomTotal - 2;
        }

        public void LoadUrv11(IProgressBarWithLogUI pb)
        {
            pb.SetMainBarVolum(4);
            pb.MainBarTick("Очистка");
            ClearDataRng11();
          
            int lastRow = GetLastRowUrv12();
            pb.MainBarTick("Определение столбцов КП");
            List<OfferAddress> addresses = GetAdderssLvl12();
            int offersCount = addresses.Count; // GetOffersCount();
            int rowPaste = 14;
            int colPaste = 9;
            int lastCol = colPaste + 3 * offersCount - 1;

            int count = lastRow - 12;
            if (count < 1) throw new AddInException($"Строки отсутствуют лист: {_SheetUrv12.Name}");
            pb.SetSubBarVolume(count);
            pb.MainBarTick("Печать строк");
            for (int row = 13; row <= lastRow; row++)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();
                string number = _SheetUrv12.Cells[row, 2].Value?.ToString() ?? ""; //Range[$"${letterNumber}{row}"].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;
                ProjectWorkbook projectWorkbook = new ProjectWorkbook();
                string name = _SheetUrv12.Cells[row, 3].Value?.ToString() ?? "";
                string cost = _SheetUrv12.Cells[row, 4].Value?.ToString() ?? "";
                number = number.Trim(new char[] { ' ', '.' });
                int levelNum = number.Split('.').Length;

                if (levelNum > 0 && levelNum < 3)
                {
                    _SheetUrv11.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    Excel.Range numberCell = _SheetUrv11.Cells[rowPaste, 2];
                    numberCell.NumberFormat = "@";
                    numberCell.Value = number;
                    _SheetUrv11.Cells[rowPaste, 3].Value = name;
                    _SheetUrv11.Cells[rowPaste, 7].Value = cost;

                    // Формат строки по уровню
                    Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);
                    if (pallets.TryGetValue(levelNum.ToString(), out Excel.Range pallet))
                    {
                        pallet.Copy();
                        _SheetUrv11.Range[_SheetUrv11.Cells[rowPaste, 2], _SheetUrv11.Cells[rowPaste, lastCol]].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    }
                    // Вывод и форматирование значений
                    // for (int columnOffer = 5; columnOffer<=1;columnOffer++)
                    foreach (OfferAddress address in addresses)
                    {
                        _SheetUrv11.Cells[rowPaste, colPaste].Value = _SheetUrv12.Cells[row, address.ColTotalCost].Value?.ToString() ?? "";
                        _SheetUrv11.Cells[rowPaste, colPaste + 1].Value = _SheetUrv12.Cells[row, address.ColPercentTotal].Value?.ToString() ?? "";
                        projectWorkbook.ColorCell(_SheetUrv11.Cells[rowPaste, colPaste + 1], levelNum.ToString());
                        _SheetUrv11.Cells[rowPaste, colPaste + 2].Value = _SheetUrv12.Cells[row, address.ColComments].Value?.ToString() ?? "";
                        colPaste += 3;
                    }
                    colPaste = 9;
                    rowPaste++;
                }
            }
            pb.MainBarTick("Удаление стр №13");
           // Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range rng = _SheetUrv11.Cells[13, 1];
            rng.EntireRow.Delete();
        }
        private int GetOffersCount()
        {
            int count = 0;
            int lastCol = _SheetUrv12.UsedRange.Column + _SheetUrv12.UsedRange.Columns.Count - 1;
            for (int col = 6; col <= lastCol; col += 5)
            {
                string text = _SheetUrv12.Cells[13, col].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(text)) break;
                count++;
            }
            return count;
        }

        /// <summary>
        ///  Номера столбцов заполненных кп
        /// </summary>
        /// <returns></returns>
        private List<OfferAddress> GetAdderssLvl12()
        {
            List<OfferAddress> addresses = new List<OfferAddress>();
            int lastCol = GetLastColumnUrv12(_SheetUrv12, 13); //_SheetUrv12.Cells[13, _SheetUrv12.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

            for (int col = 6; col <= lastCol; col += 5)
            {
                string val = _SheetUrv12.Cells[13, col].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(val))
                {
                    OfferAddress address = new OfferAddress
                    {
                        ColTotalCost = col,
                        ColPercentTotal = col + 2,
                        ColComments = col + 3
                    };
                    addresses.Add(address);
                }
            }
            return addresses;
        }

        private void ClearDataRng12()
        {
            int lastRow = GetLastRowUrv12();
            if (lastRow <= 14) return;
            int lastColumn = _SheetUrv12.UsedRange.Column + _SheetUrv12.UsedRange.Columns.Count - 1;
            Excel.Range dataRng = _SheetUrv12.Range[_SheetUrv12.Cells[14, 2], _SheetUrv12.Cells[lastRow, lastColumn]];
            dataRng.EntireRow.Delete();
            dataRng = _SheetUrv12.Range[_SheetUrv12.Cells[13, 2], _SheetUrv12.Cells[13, lastColumn]];
            dataRng.ClearContents();
            //TODO формат dataRng.
            return;
        }
        private void ClearDataRng11()
        {
            int lastRow = GetLastRowUrv11();
            if (lastRow <= 14) return;
            int lastColumn = _SheetUrv11.UsedRange.Column + _SheetUrv11.UsedRange.Columns.Count - 1;
            Excel.Range dataRng = _SheetUrv11.Range[_SheetUrv11.Cells[14, 2], _SheetUrv11.Cells[lastRow, lastColumn]];
            dataRng.EntireRow.Delete();
            dataRng = _SheetUrv11.Range[_SheetUrv11.Cells[13, 2], _SheetUrv11.Cells[13, lastColumn]];
            dataRng.ClearContents();
            //TODO формат dataRng.
            return;
        }

        private int GetLastColumnUrv12(Excel.Worksheet sh, int row)
        {
            int lastCol = sh.Cells[row, sh.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            return lastCol;
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

        internal void UpdateUrv11(IProgressBarWithLogUI pb)
        {
            int lastRow = _SheetUrv11.UsedRange.Row + _SheetUrv11.UsedRange.Rows.Count - 1;
            int rowStart = 13;
            int rowPaste = 14;

            //    int colPaste = 6;
            for (int row = _project.RowStart; row <= lastRow; row++)
            {
                string number = _SheetUrv11.Cells[row, 2].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;
                //        if (_SheetUrv12.Cells[rowPaste, 2].Value == number)
                //        {
                //            string level = _AnalisysSheet.Range[$"${letterLevel}{row}"].Value?.ToString() ?? "";
                //            int levelNum = int.TryParse(level, out int ln) ? ln : 0;

                //            if (levelNum > 0 && levelNum < 6)
                //            {
                //                foreach (OfferAddress address in addresses)
                //                {
                //                    _SheetUrv12.Cells[rowPaste, colPaste].Value = _AnalisysSheet.Cells[row, address.ColTotalCost].Value?.ToString() ?? "";
                //                    _SheetUrv12.Cells[rowPaste, colPaste + 1].Value = _AnalisysSheet.Cells[row, address.ColPercentMaterial].Value?.ToString() ?? "";
                //                    _SheetUrv12.Cells[rowPaste, colPaste + 2].Value = _AnalisysSheet.Cells[row, address.ColPercentWorks].Value?.ToString() ?? "";
                //                    _SheetUrv12.Cells[rowPaste, colPaste + 3].Value = _AnalisysSheet.Cells[row, address.ColPercentTotal].Value?.ToString() ?? "";
                //                    _SheetUrv12.Cells[rowPaste, colPaste + 4].Value = _AnalisysSheet.Cells[row, address.ColComments].Value?.ToString() ?? "";
                //                    colPaste += 5;
                //                }
                //                colPaste = 6;
                //                rowPaste++;
                //            }
                //        }

            }
        }
    }
}
