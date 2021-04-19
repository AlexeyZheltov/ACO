﻿using ACO.ExcelHelpers;
using ACO.ProjectBook;
using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        readonly Excel.Application _app = Globals.ThisAddIn.Application;
        readonly Excel.Worksheet _SheetUrv12;
        readonly Excel.Worksheet _SheetUrv11;
        readonly Excel.Worksheet _SheetPalette;
        readonly Excel.Worksheet _AnalisysSheet;
        readonly ProjectManager.ProjectManager _projectManager;
        readonly Project _project;
        readonly IProgressBarWithLogUI pb;

        public Pivot(IProgressBarWithLogUI pb)
        {
            this.pb = pb;
            Excel.Workbook wb = _app.ActiveWorkbook;
            _SheetUrv12 = ExcelHelper.GetSheet(wb, "Урв12");
            _SheetUrv11 = ExcelHelper.GetSheet(wb, "Урв11");
            _SheetPalette = ExcelHelper.GetSheet(wb, "Палитра");
            _projectManager = new ProjectManager.ProjectManager();
            _project = _projectManager.ActiveProject;
            string analisysSheetName = _project.AnalysisSheetName;
            _AnalisysSheet = ExcelHelper.GetSheet(wb, analisysSheetName);
        }

        private void PasteTitleOffer12(int colPaste)
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int rowCostWorks = ExcelHelper.FindCell(_SheetUrv12, "СТОИМОСТЬ НЕКОТОРЫХ РАБОТ").Row;
            Excel.Range rngTitle = _SheetUrv12.Range["F10:J12"];
            rngTitle.Copy();
            _SheetUrv12.Cells[10, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
            rngTitle = _SheetUrv12.Range[$"F{rowBottomTotal - 1}:J{rowCostWorks}"];
            rngTitle.Copy();
            _SheetUrv12.Cells[rowBottomTotal - 1, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);

            //Формат строк в столбце
            int lastrow = _SheetUrv12.UsedRange.Row + _SheetUrv12.UsedRange.Rows.Count - 2;
            Excel.Range formatCell = _SheetUrv12.Range[$"F{rowCostWorks + 1}"];
            Excel.Range rng = _SheetUrv12.Range[_SheetUrv12.Cells[rowCostWorks + 1, colPaste], _SheetUrv12.Cells[lastrow, colPaste]];
            formatCell.Copy();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            formatCell = _SheetUrv12.Range[$"G{rowCostWorks + 1}"];
            rng = _SheetUrv12.Range[_SheetUrv12.Cells[rowCostWorks + 1, colPaste], _SheetUrv12.Cells[lastrow, colPaste + 4]];
            formatCell.Copy();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        }


        private void PrintTitlesOffers12(List<OfferAddress> addresses)
        {
            int i = 0;
            foreach (OfferAddress offer in addresses)
            {
                int colPaste = 6 + 5 * i;
                Excel.Range cell = _SheetUrv12.Cells[10, colPaste];
                if (cell.Value is null)
                {
                    PasteTitleOffer12(colPaste);
                    string headerName = cell.Value?.ToString() ?? "";
                    cell.Value = headerName.Replace("УЧАСТНИК 1", offer.Name);
                }
                i++;
            }
        }

        /// Добавить список
        /// Добавить столбцы КП
        /// Проставить формулы
        public void LoadUrv12()
        {
            pb.SetMainBarVolum(6);
            pb.MainBarTick("Очистка");
            ClearDataRng12();
            string letterName = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Name]).ColumnSymbol;
            string letterNumber = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            string letterLevel = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            string letterCost = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol;

            int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
            pb.MainBarTick("Определение столбцов КП");
            List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;

            pb.Writeline("Копирование заголовков");
            PrintTitlesOffers12(addresses);
            int rowPaste = 14;
            int colPaste = 6;
            int lastCol = colPaste + 5 * addresses.Count - 1;

            int count = lastRow - _project.RowStart + 1;
            if (count < 1) throw new AddInException($"Строки отсутствуют лист: {_project.AnalysisSheetName}");
            pb.SetSubBarVolume(count);
            pb.MainBarTick("Заполнение строк");
            ProjectWorkbook projectWorkbook = new ProjectWorkbook();

            Excel.Range dataRange = projectWorkbook.GetAnalysisRange();

            for (int row = _project.RowStart; row <= lastRow; row++)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();
                Excel.Range cellNumber = _AnalisysSheet.Range[$"${letterNumber}{row}"];
                int columnCellNumber = cellNumber.Column;
                string number = cellNumber.Value?.ToString() ?? "";

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
                    string letterOutName = ExcelHelper.GetColumnLetter(numberCell);
                    int colTotalCost = ExcelHelper.GetColumn(letterCost, _AnalisysSheet);
                    int column = colTotalCost - columnCellNumber + 1;
                    _SheetUrv12.Cells[rowPaste, 4].Formula =
                         $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {column}, FALSE)";

                    // Формат строки по уровню
                    Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);
                    if (pallets.TryGetValue(level, out Excel.Range pallet))
                    {
                        ExcelHelper.SetCellFormat(_SheetUrv12.Range[_SheetUrv12.Cells[rowPaste, 2], _SheetUrv12.Cells[rowPaste, lastCol]], pallet);
                    }
                    // Вывод и форматирование значений
                    foreach (OfferAddress address in addresses)
                    {
                        int col = address.ColTotalCost - columnCellNumber + 1;
                        //РУБ. РФ
                        _SheetUrv12.Cells[rowPaste, colPaste].Formula =
                           $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                        //% отклонения материалы
                        col = address.ColPercentMaterials - columnCellNumber + 1;
                        _SheetUrv12.Cells[rowPaste, colPaste + 1].Formula =
                          $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                        // _SheetUrv12.Cells[rowPaste, colPaste + 1].NumberFormat = "0%";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 1], level);

                        //% отклонения работы
                        col = address.ColPercentWorks - columnCellNumber + 1;
                        _SheetUrv12.Cells[rowPaste, colPaste + 2].Formula =
                          $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                        // _SheetUrv12.Cells[rowPaste, colPaste + 2].NumberFormat = "0%";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 2], level);

                        // % отклонения всего
                        string letterOutTotalDiff = ExcelHelper.GetColumnLetter(_SheetUrv12.Cells[rowPaste, colPaste]);
                        _SheetUrv12.Cells[rowPaste, colPaste + 3].Formula = $"=${letterOutTotalDiff}{rowPaste}/$D{rowPaste}-1";
                        // _SheetUrv12.Cells[rowPaste, colPaste + 3].NumberFormat = "0%";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 3], level);

                        //КОММЕНТАРИИ К СТОИМОСТИ
                        col = address.ColComments - columnCellNumber + 1;
                        _SheetUrv12.Cells[rowPaste, colPaste + 4].Formula =
                            $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";

                        colPaste += 5;
                    }
                    colPaste = 6;
                    rowPaste++;
                }
            }
            pb.MainBarTick("Формулы итогов");
            TotalFormuls12();
            pb.MainBarTick("Формат ячеек");
            SetNumberFormat12(addresses.Count);
            pb.MainBarTick("Общие комментарии");
            new OfferInfo(projectWorkbook).SetInfo();
            
            pb.MainBarTick("Удаление строки №13");
            Excel.Range rng = _SheetUrv12.Cells[13, 1];
            rng.EntireRow.Delete();
            _SheetUrv12.Activate();
        }


        /// <summary>
        ///  Формат 
        /// </summary>
        /// <param name="addresses"></param>
        private void SetNumberFormat12(int addressesCount)
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            // int colPaste6 = 6;
            Excel.Range rng = _SheetUrv12.Range[_SheetUrv12.Cells[13, 4], _SheetUrv12.Cells[rowBottomTotal - 2, 4]];
            rng.NumberFormat = "# ##0,00";

            int lastCol = 6 + 5 * addressesCount - 1;
            for (int col = 6; col <= lastCol; col += 5)
            {
                rng = _SheetUrv12.Range[_SheetUrv12.Cells[13, col], _SheetUrv12.Cells[rowBottomTotal - 2, col]];
                rng.NumberFormat = "# ##0,00";
                rng = _SheetUrv12.Range[_SheetUrv12.Cells[13, col + 1], _SheetUrv12.Cells[rowBottomTotal - 2, col + 1]];
                rng.NumberFormat = "0%";
                rng = _SheetUrv12.Range[_SheetUrv12.Cells[13, col + 2], _SheetUrv12.Cells[rowBottomTotal - 2, col + 2]];
                rng.NumberFormat = "0%";
                rng = _SheetUrv12.Range[_SheetUrv12.Cells[13, col + 3], _SheetUrv12.Cells[rowBottomTotal - 2, col + 3]];
                rng.NumberFormat = "0%";
            }
        }


        /// <summary>
        ///  Обновление значений урв 12
        /// </summary>
        /// <param name="pb"></param>
        internal void UpdateUrv12()
        {
            pb.SetMainBarVolum(1);
            pb.MainBarTick("Обновление формул \"Урв 12\"");
            int rowStart = 13;
            int colPaste = 1;

            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            Excel.Range dataRange = projectWorkbook.GetAnalysisRange();
            List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;
            pb.Writeline("Копирование заголовков");
            PrintTitlesOffers12(addresses);

            string letterNumber = projectWorkbook.GetLetter(StaticColumns.Number);
            Excel.Range cellNumber = _AnalisysSheet.Range[$"${letterNumber}{_project.RowStart}"];
            int columnCellNumber = cellNumber.Column;
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal - 2;

            pb.SetSubBarVolume(addresses.Count);
            //// Вывод и форматирование значений
            foreach (OfferAddress address in addresses)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();
                colPaste += 5;

             
                string textCost = _SheetUrv12.Cells[rowStart, colPaste].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(textCost)) continue;  // Пропустить заполненные КП

                string formulaSumm = "";
                for (int row = rowStart; row <= lastRow; row++)
                {
                    string number = _SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number)) continue;
                    number = number.Trim(new char[] { ' ', '.' });
                    int levelNum = number.Split('.').Length;
                    if (levelNum == 1)
                    {
                        formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={_SheetUrv12.Cells[row, colPaste].Address}" :
                                                                                $"+{_SheetUrv12.Cells[row, colPaste].Address}";
                    }
                    // Формат строки по уровню
                    Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);
                    string keyLvl = levelNum.ToString();
                    if (pallets.TryGetValue(keyLvl, out Excel.Range pallet))
                    {
                        ExcelHelper.SetCellFormat(_SheetUrv12.Range[_SheetUrv12.Cells[row, colPaste], _SheetUrv12.Cells[row, colPaste + 4]], pallet);
                    }
                }
                int col = address.ColTotalCost - columnCellNumber + 1;
                //РУБ. РФ
                _SheetUrv12.Cells[rowStart, colPaste].Formula =
                   $"= VLOOKUP($B{rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                //% отклонения материалы
                col = address.ColPercentMaterials - columnCellNumber + 1;
                _SheetUrv12.Cells[rowStart, colPaste + 1].Formula =
                  $"= VLOOKUP($B{rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                //% отклонения работы
                col = address.ColPercentWorks - columnCellNumber + 1;
                _SheetUrv12.Cells[rowStart, colPaste + 2].Formula =
                  $"= VLOOKUP($B{rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                // % отклонения всего
                string letterOutTotalDiff = ExcelHelper.GetColumnLetter(_SheetUrv12.Cells[rowStart, colPaste]);
                _SheetUrv12.Cells[rowStart, colPaste + 3].Formula = $"=${letterOutTotalDiff}{rowStart}/$D{rowStart}-1";
                //КОММЕНТАРИИ К СТОИМОСТИ
                col = address.ColComments - columnCellNumber + 1;
                _SheetUrv12.Cells[rowStart, colPaste + 4].Formula =
                    $"= VLOOKUP($B{rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";

                Excel.Range rng = _SheetUrv12.Range[_SheetUrv12.Cells[rowStart, colPaste], _SheetUrv12.Cells[rowStart, colPaste + 4]];
                Excel.Range destination = _SheetUrv12.Range[_SheetUrv12.Cells[rowStart, colPaste], _SheetUrv12.Cells[lastRow, colPaste + 4]];
                rng.AutoFill(destination, Excel.XlAutoFillType.xlFillValues);
                destination.Columns[2].NumberFormat = "0%";
                destination.Columns[3].NumberFormat = "0%";
                destination.Columns[4].NumberFormat = "0%";

                if (!string.IsNullOrEmpty(formulaSumm))
                {
                    _SheetUrv12.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
                    _SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Formula =
                                        $"={_SheetUrv12.Cells[rowBottomTotal, colPaste].Address}*0.2";
                    _SheetUrv12.Cells[rowBottomTotal + 2, colPaste].Formula =
                                        $"={_SheetUrv12.Cells[rowBottomTotal, colPaste].Address}+" +
                                        $"{_SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Address}";
                }
                //TODO подсчитать кол-во.
                PrintTotalComments(address);

            }
            //TODO загрузить наиболее дорогии позиции 
        }

        private void PrintTotalComments(OfferAddress address)
        {
            int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
            int countChangedName = 0;
            for (int row = _project.RowStart; row <= lastRow; row++)
            {
                string changedName = _AnalisysSheet.Cells[row, address.ColStartOfferComments].Value?.ToString() ?? "";
                if (changedName == "ЛОЖЬ") { countChangedName++; }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void TotalFormuls12()
        {
            List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal - 2;
            string formulaSumm = "";
            for (int row = 13; row <= lastRow; row++)
            {
                string number = _SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;
                number = number.Trim(new char[] { ' ', '.' });
                int levelNum = number.Split('.').Length;

                if (levelNum == 1)
                {
                    formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={_SheetUrv12.Cells[row, 4].Address}" :
                                                                            $"+{_SheetUrv12.Cells[row, 4].Address}";
                }
            }
            if (!string.IsNullOrEmpty(formulaSumm))
            {

                _SheetUrv12.Cells[rowBottomTotal, 4].Formula = formulaSumm;
                _SheetUrv12.Cells[rowBottomTotal + 1, 4].Formula =
                                    $"={_SheetUrv12.Cells[rowBottomTotal, 4].Address}*0.2";
                _SheetUrv12.Cells[rowBottomTotal + 2, 4].Formula =
                                    $"={_SheetUrv12.Cells[rowBottomTotal, 4].Address}+" +
                                    $"{_SheetUrv12.Cells[rowBottomTotal + 1, 4].Address}";
            }

            int colPaste = 6;
            foreach (OfferAddress address in addresses)
            {
                formulaSumm = "";
                //string formulaSumm
                for (int row = 13; row <= lastRow; row++)
                {
                    string number = _SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number)) continue;
                    number = number.Trim(new char[] { ' ', '.' });
                    int levelNum = number.Split('.').Length;
                    if (levelNum == 1)
                    {
                        formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={_SheetUrv12.Cells[row, colPaste].Address}" :
                                                                                $"+{_SheetUrv12.Cells[row, colPaste].Address}";
                    }
                }
                if (!string.IsNullOrEmpty(formulaSumm))
                {
                    _SheetUrv12.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
                    _SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Formula =
                                        $"={_SheetUrv12.Cells[rowBottomTotal, colPaste].Address}*0.2";
                    _SheetUrv12.Cells[rowBottomTotal + 2, colPaste].Formula =
                                        $"={_SheetUrv12.Cells[rowBottomTotal, colPaste].Address}+" +
                                        $"{_SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Address}";
                    colPaste += 5;
                }

               // string formula = GetFormuls();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void TotalFormuls11()
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv11, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal - 2;
            int colPaste = 7;
            string formulaSumm = "";
            for (int row = 13; row <= lastRow; row++)
            {

                string number = _SheetUrv11.Cells[row, 2].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;
                number = number.Trim(new char[] { ' ', '.' });
                int levelNum = number.Split('.').Length;

                if (levelNum == 1)
                {
                    formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={_SheetUrv11.Cells[row, colPaste].Address}" :
                                                                         $"+{_SheetUrv11.Cells[row, colPaste].Address}";
                }
            }
            if (!string.IsNullOrEmpty(formulaSumm))
            {
                _SheetUrv11.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
            }
            _SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Formula =
                 $"={_SheetUrv11.Cells[rowBottomTotal, colPaste].Address}*0.2";
            _SheetUrv11.Cells[rowBottomTotal + 2, colPaste].Formula =
                                $"={_SheetUrv11.Cells[rowBottomTotal, colPaste].Address}+" +
                                $"{_SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Address}";

          //  _SheetUrv11.Cells[rowBottomTotal + 5, colPaste+1].Formula =
          
          // _SheetUrv11.Cells[rowBottomTotal + 5, colPaste ].Formula =
            //                    $"={_SheetUrv11.Cells[rowBottomTotal+4, colPaste].Address}+{_SheetUrv11.Cells[rowBottomTotal+2, colPaste].Address}";

            List<OfferAddress> addresses = GetAdderssLvl12();

            colPaste = 9;
            foreach (OfferAddress address in addresses)
            {
                formulaSumm = "";
                for (int row = 13; row <= lastRow; row++)
                {
                    string number = _SheetUrv11.Cells[row, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number)) continue;
                    number = number.Trim(new char[] { ' ', '.' });
                    int levelNum = number.Split('.').Length;

                    if (levelNum == 1)
                    {
                        formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={_SheetUrv11.Cells[row, colPaste].Address}" :
                                                                                $"+{_SheetUrv11.Cells[row, colPaste].Address}";
                    }
                }
                if (!string.IsNullOrEmpty(formulaSumm))
                {
                    _SheetUrv11.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
                }
                _SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Formula =
                                    $"={_SheetUrv11.Cells[rowBottomTotal, colPaste].Address}*0.2";
                _SheetUrv11.Cells[rowBottomTotal + 2, colPaste].Formula =
                                    $"={_SheetUrv11.Cells[rowBottomTotal, colPaste].Address}+" +
                                    $"{_SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Address}";
                // _SheetUrv11.Cells[rowBottomTotal + 5, colPaste + 1].Formula =
                //                $"={_SheetUrv11.Cells[rowBottomTotal, colPaste].Address}/$G{rowBottomTotal}-1";
                //_SheetUrv11.Cells[rowBottomTotal + 5, colPaste + 1].NumberFormat = "0%";
                //_SheetUrv11.Cells[rowBottomTotal + 5, colPaste].Formula =
                //                    $"={_SheetUrv11.Cells[rowBottomTotal + 4, colPaste].Address}+{_SheetUrv11.Cells[rowBottomTotal + 2, colPaste].Address}";


                colPaste += 3;
            }
        }

        private int GetLastRowUrv12()
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            return rowBottomTotal - 2;
        }
        private int GetLastRowUrv11()
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv11, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            return rowBottomTotal - 2;
        }


        private void PasteTitleOffer11(int colPaste)
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv11, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            Excel.Range rngTitle = _SheetUrv11.Range["I10:K12"];
            rngTitle.Copy();
            _SheetUrv11.Cells[10, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
            rngTitle = _SheetUrv11.Range[$"I{rowBottomTotal - 1}:K{rowBottomTotal + 6}"];
            rngTitle.Copy();
            _SheetUrv11.Cells[rowBottomTotal - 1, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
        }

        private void PrintTitlesOffers11(List<OfferAddress> addresses)
        {
            int i = 0;
            foreach (OfferAddress offer in addresses)
            {
                int colPaste = 9 + 3 * i;
                Excel.Range cell = _SheetUrv11.Cells[10, colPaste];
                if (cell.Value is null)
                {
                    PasteTitleOffer11(colPaste);
                    string headerName = cell.Value?.ToString() ?? "";
                    cell.Value = offer.Name;
                }
                i++;
            }
        }

        public void LoadUrv11()
        {
            pb.SetMainBarVolum(5);
            pb.MainBarTick("Очистка");
            ClearDataRng11();
            int lastRow = GetLastRowUrv12();
            pb.MainBarTick("Определение столбцов КП");
            Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);
            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            List<OfferAddress> addresses = GetAdderssLvl12();
            pb.Writeline("Копирование заголовков");
            PrintTitlesOffers11(addresses);

            int offersCount = addresses.Count;
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
                string number = _SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;

                string name = _SheetUrv12.Cells[row, 3].Value?.ToString() ?? "";

                number = number.Trim(new char[] { ' ', '.' });
                int levelNum = number.Split('.').Length;

                if (levelNum > 0 && levelNum < 3)
                {

                    string outRowNumber = _SheetUrv11.Cells[rowPaste, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(outRowNumber))
                    {
                        _SheetUrv11.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    }
                    Excel.Range numberCell = _SheetUrv11.Cells[rowPaste, 2];
                    numberCell.NumberFormat = "@";
                    numberCell.Value = number;

                    _SheetUrv11.Cells[rowPaste, 3].Value = name;
                    _SheetUrv11.Cells[rowPaste, 7].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, 4].Address}";

                    // Формат строки по уровню
                    if (pallets.TryGetValue(levelNum.ToString(), out Excel.Range pallet))
                    {
                        ExcelHelper.SetCellFormat(_SheetUrv11.Range[_SheetUrv11.Cells[rowPaste, 2], 
                                                  _SheetUrv11.Cells[rowPaste, lastCol]], pallet);
                    }

                    foreach (OfferAddress address in addresses)
                    {
                        PrintValuesFormuls(address, row, rowPaste, colPaste);
                        projectWorkbook.ColorCell(_SheetUrv11.Cells[rowPaste, colPaste + 1], levelNum.ToString());

                        // Вывод и форматирование значений
                        //_SheetUrv11.Cells[rowPaste, colPaste].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, address.ColTotalCost].Address}";
                        //string letterOutTotalDiff = ExcelHelper.GetColumnLetter(_SheetUrv12.Cells[rowPaste, colPaste]);
                        //_SheetUrv11.Cells[rowPaste, colPaste + 1].Formula = $"=${letterOutTotalDiff}{rowPaste}/$G{rowPaste}-1";
                        //_SheetUrv11.Cells[rowPaste, colPaste + 1].NumberFormat = "0%";
                        //projectWorkbook.ColorCell(_SheetUrv11.Cells[rowPaste, colPaste + 1], levelNum.ToString());
                        //_SheetUrv11.Cells[rowPaste, colPaste + 2].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, address.ColComments].Address}";
                        colPaste += 3;
                    }
                    colPaste = 9;
                    rowPaste++;
                }
            }
            TotalFormuls11();
            pb.MainBarTick("Удаление стр №13");
            Excel.Range rng = _SheetUrv11.Cells[13, 1];
            rng.EntireRow.Delete();
            pb.MainBarTick("Обновление диаграммы");
            UpdateDiagramm();
            _SheetUrv11.Activate();
        }

        /// <summary>
        /// 
        /// </summary>
        private void PrintValuesFormuls(OfferAddress address,int row, int rowPaste, int colPaste)
        {
            // Вывод и форматирование значений
            _SheetUrv11.Cells[rowPaste, colPaste].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, address.ColTotalCost ].Address}";
            string letterOutTotalDiff = ExcelHelper.GetColumnLetter(_SheetUrv12.Cells[rowPaste, colPaste]);
            _SheetUrv11.Cells[rowPaste, colPaste + 1].Formula = $"=${letterOutTotalDiff}{rowPaste}/$G{rowPaste}-1";
            _SheetUrv11.Cells[rowPaste, colPaste + 1].NumberFormat = "0%";
           
            _SheetUrv11.Cells[rowPaste, colPaste + 2].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, address.ColComments].Address}";
        }

        /// <summary>
        ///  Обновление Диаграммы
        /// </summary>
        internal void UpdateDiagramm()
        {
            Excel.ChartObject shp = _SheetUrv11.ChartObjects("Chart 2");
            Excel.Chart chartPage = shp.Chart;
            Excel.SeriesCollection seriesCol = (Excel.SeriesCollection)chartPage.SeriesCollection();
            Excel.FullSeriesCollection fullColl = chartPage.FullSeriesCollection();
            Debug.WriteLine(fullColl.Count);

            int rowStart = 13;
            int lastCol = GetLastColumnUrv(_SheetUrv11, 13);
            int lastRow = GetLastRowUrv11();
            int ix = 1;
            string letterCost = "G";
            fullColl.Item(ix).Name = $"={_SheetUrv11.Name}!${letterCost}10";
            fullColl.Item(ix).Values = $"={_SheetUrv11.Name}!${letterCost}{rowStart}:${letterCost}{lastRow}";
            fullColl.Item(ix).XValues = $"={_SheetUrv11.Name}!$C{rowStart}:$C{lastRow}";

            for (int col = 9; col <= lastCol; col += 3)
            {
                Excel.Range cellFirstCost = _SheetUrv11.Cells[13, col];
                string text = cellFirstCost.Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(text)) continue;
                letterCost = ExcelHelper.GetColumnLetter(cellFirstCost);
                ix++;
                if (ix > fullColl.Count)
                {
                    seriesCol.NewSeries();
                }
                fullColl.Item(ix).Name = $"={_SheetUrv11.Name}!${letterCost}10";
                fullColl.Item(ix).Values = $"={_SheetUrv11.Name}!${letterCost}{rowStart}:${letterCost}{lastRow}";
                fullColl.Item(ix).XValues = $"={_SheetUrv11.Name}!$C{rowStart}:$C{lastRow}";
            }
            if (ix < fullColl.Count)
            {
                for (int i = ix + 1; i <= fullColl.Count; i++)
                {
                    fullColl.Item(i).Delete();
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
            int lastCol = GetLastColumnUrv(_SheetUrv12, 13);

            for (int col = 6; col <= lastCol; col += 5)
            {
                string val = _SheetUrv12.Cells[13, col].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(val))
                {
                    OfferAddress address = new OfferAddress
                    {
                        ColTotalCost = col,
                        ColPercentTotal = col + 3,
                        ColComments = col + 4,
                        Name = _SheetUrv12.Cells[10, col].Value?.ToString() ?? ""
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
            return;
        }

        private int GetLastColumnUrv(Excel.Worksheet sh, int row)
        {
            int lastCol = sh.Cells[row, sh.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            return lastCol;
        }

        internal void UpdateUrv11()
        {
            pb.SetMainBarVolum(2);
            pb.MainBarTick("Обновление \"Урв 11\"");
            int lastRowSh12 = GetLastRowUrv12();
            int lastRowSh11 = GetLastRowUrv11();
            int colPaste = 9;

            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);

            List<OfferAddress> addresses = GetAdderssLvl12();
            pb.Writeline("Копирование заголовков");
            PrintTitlesOffers11(addresses);
            pb.SetSubBarVolume(addresses.Count);

            foreach (OfferAddress address in addresses)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();
                for (int rowPaste = 13; rowPaste <= lastRowSh11; rowPaste++)
                {
                    string number11 = _SheetUrv11.Cells[rowPaste, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number11)) continue;

                    for (int row = 13; row <= lastRowSh12; row++)
                    {
                        string number12 = _SheetUrv11.Cells[row, 2].Value?.ToString() ?? "";
                        if (number12 != number11) continue;
                        number11 = number11.Trim(new char[] { ' ', '.' });
                        int levelNum = number11.Split('.').Length;

                        if (levelNum > 0 && levelNum < 3)
                        {
                            // Формат строки по уровню
                            if (pallets.TryGetValue(levelNum.ToString(), out Excel.Range pallet))
                            {
                                ExcelHelper.SetCellFormat(_SheetUrv11.Range[_SheetUrv11.Cells[rowPaste, colPaste], _SheetUrv11.Cells[rowPaste, colPaste + 2]], pallet);
                            }

                            // Вывод и форматирование значений
                            _SheetUrv11.Cells[rowPaste, colPaste].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, address.ColTotalCost].Address}";
                            string letterOutTotalDiff = ExcelHelper.GetColumnLetter(_SheetUrv12.Cells[row, colPaste]);
                            _SheetUrv11.Cells[rowPaste, colPaste + 1].Formula = $"=${letterOutTotalDiff}{rowPaste}/$G{rowPaste}-1";
                            _SheetUrv11.Cells[rowPaste, colPaste + 1].NumberFormat = "0%";
                            projectWorkbook.ColorCell(_SheetUrv11.Cells[rowPaste, colPaste + 1], levelNum.ToString());
                            _SheetUrv11.Cells[rowPaste, colPaste + 2].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, address.ColComments].Address}";
                        }
                    }
                }
                colPaste += 3;
            }
            pb.MainBarTick("Обновление диаграммы");
            UpdateDiagramm();
        }
    }
}
