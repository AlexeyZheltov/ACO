using ACO.ExcelHelpers;
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
                    // _SheetUrv12.Cells[rowPaste, 4].Value = cost;
                    string letterOutName = ExcelHelper.GetColumnLetter(numberCell);
                    int colTotalCost = ExcelHelper.GetColumn(letterCost, _AnalisysSheet);
                    int column = colTotalCost - columnCellNumber + 1;
                    _SheetUrv12.Cells[rowPaste, 4].Formula =
                         $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {column}, FALSE)";

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

                        int col = address.ColTotalCost - columnCellNumber + 1;
                        //РУБ. РФ
                        _SheetUrv12.Cells[rowPaste, colPaste].Formula =
                           $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                        //_SheetUrv12.Cells[rowPaste, colPaste].Value = _AnalisysSheet.Cells[row, address.ColTotalCost].Value?.ToString() ?? "";

                        //% отклонения материалы
                        col = address.ColPercentWorks - columnCellNumber + 1;
                        _SheetUrv12.Cells[rowPaste, colPaste + 1].Formula =
                          $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                        _SheetUrv12.Cells[rowPaste, colPaste + 1].NumberFormat = "0%";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 1], level);
                        //_SheetUrv12.Cells[rowPaste, colPaste + 1].Value = _AnalisysSheet.Cells[row, address.ColPercentMaterial].Value?.ToString() ?? "";
                        // projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 1], level);

                        //% отклонения работы
                        col = address.ColPercentWorks - columnCellNumber + 1;
                        _SheetUrv12.Cells[rowPaste, colPaste + 2].Formula =
                          $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                        _SheetUrv12.Cells[rowPaste, colPaste + 2].NumberFormat = "0%";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 2], level);
                        //_SheetUrv12.Cells[rowPaste, colPaste + 2].Value = _AnalisysSheet.Cells[row, address.ColPercentWorks].Value?.ToString() ?? "";
                        //projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 2], level);

                        // % отклонения всего
                        string letterOutTotalDiff = ExcelHelper.GetColumnLetter(_SheetUrv12.Cells[rowPaste, colPaste]);
                        _SheetUrv12.Cells[rowPaste, colPaste + 3].Formula = $"=${letterOutTotalDiff}{rowPaste}/$D{rowPaste}-1";
                        _SheetUrv12.Cells[rowPaste, colPaste + 3].NumberFormat = "0%";
                        projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 3], level);
                        //  _SheetUrv12.Cells[rowPaste, colPaste + 4].Value = _AnalisysSheet.Cells[row, address.ColComments].Value?.ToString() ?? "";

                        //КОММЕНТАРИИ К СТОИМОСТИ
                        col = address.ColComments - columnCellNumber + 1;
                        _SheetUrv12.Cells[rowPaste, colPaste + 4].Formula =
                            $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";

                        //projectWorkbook.ColorCell(_SheetUrv12.Cells[rowPaste, colPaste + 3], level);
                        // _SheetUrv12.Cells[rowPaste, colPaste + 3].Value = _AnalisysSheet.Cells[row, address.ColPercentTotal].Value?.ToString() ?? "";

                        colPaste += 5;
                    }
                    colPaste = 6;
                    rowPaste++;
                }
            }
            TotalFormuls12();
            pb.MainBarTick("Удаление стр №13");
            Excel.Range rng = _SheetUrv12.Cells[13, 1];
            rng.EntireRow.Delete();
        }

        /// <summary>
        ///  Обновление значений урв 12
        /// </summary>
        /// <param name="pb"></param>
        internal void UpdateUrv12(IProgressBarWithLogUI pb)
        {
            pb.SetMainBarVolum(1);
            pb.MainBarTick("Обновление формул \"Урв 12\"");
            int rowStart = 13;
            int colPaste = 1;
            int ix = 0;
            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            Excel.Range dataRange = projectWorkbook.GetAnalysisRange();
            List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;
            //Excel.Range cellNumber = _SheetUrv12.Cells[rowStart, 2];
            //string letterNumber = 

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

                ix++;
                if (ix > 4)
                {
                    Excel.Range rngTitle = _SheetUrv11.Range["K10:O12"];
                    rngTitle.Copy();
                    _SheetUrv12.Cells[10, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
                    rngTitle = _SheetUrv12.Range[$"K{rowBottomTotal}:O{rowBottomTotal}"];
                    rngTitle.Copy();
                    _SheetUrv12.Cells[rowBottomTotal, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
                }

                colPaste += 5;
                pb.SubBarTick();
                int col = address.ColTotalCost - columnCellNumber + 1;
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
                        pallet.Copy();
                        _SheetUrv12.Range[_SheetUrv12.Cells[row, colPaste], _SheetUrv12.Cells[row, colPaste + 4]].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    }
                }

                //РУБ. РФ
                _SheetUrv12.Cells[rowStart, colPaste].Formula =
                   $"= VLOOKUP($B{rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                //% отклонения материалы
                col = address.ColPercentWorks - columnCellNumber + 1;
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
            }
        }
        //====================================================

        private void TotalFormuls12()
        {
            List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal - 2;
            //int lastRow = GetLastRowUrv12();
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
            }
        }



        private void TotalFormuls11()
        {
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv11, "СУММА ПРЕДЛОЖЕННЫХ ОПТИМИЗАЦИЙ (с НДС)").Row;
            int lastRow = rowBottomTotal - 2; //GetLastRowUrv11();
            int colPaste = 9;

            List<OfferAddress> addresses = GetAdderssLvl12();
            foreach (OfferAddress address in addresses)
            {
                string formulaSumm = "";
                for (int row = 13; row <= lastRow; row++)
                {
                    string number = _SheetUrv11.Cells[row, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number)) continue;
                    number = number.Trim(new char[] { ' ', '.' });
                    int levelNum = number.Split('.').Length;

                    if (levelNum == 1)
                    {
                        formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={_SheetUrv11.Cells[row, colPaste]}" :
                                                                                $"+{_SheetUrv11.Cells[row, colPaste]}";
                    }
                }
                // int colPaste = address.ColTotalCost;
                if (!string.IsNullOrEmpty(formulaSumm))
                {
                    _SheetUrv11.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
                }
                _SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Formula =
                                    $"={_SheetUrv12.Cells[rowBottomTotal, colPaste].Address}*0.2";
                _SheetUrv11.Cells[rowBottomTotal + 2, colPaste].Formula =
                                    $"={_SheetUrv12.Cells[rowBottomTotal, colPaste].Address}+" +
                                    $"{_SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Address}";

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
            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv11, "СУММА ПРЕДЛОЖЕННЫХ ОПТИМИЗАЦИЙ (с НДС)").Row;
            return rowBottomTotal - 2;
        }

        public void LoadUrv11(IProgressBarWithLogUI pb)
        {
            pb.SetMainBarVolum(5);
            pb.MainBarTick("Очистка");
            ClearDataRng11();
            int lastRow = GetLastRowUrv12();
            pb.MainBarTick("Определение столбцов КП");
            Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);
            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            List<OfferAddress> addresses = GetAdderssLvl12();
            int offersCount = addresses.Count;
            int ix = 0;
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
                // string cost = _SheetUrv12.Cells[row, 4].Value?.ToString() ?? "";
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
                    _SheetUrv11.Cells[rowPaste, 7].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, 4].Address}"; //cost;

                    // Формат строки по уровню

                    if (pallets.TryGetValue(levelNum.ToString(), out Excel.Range pallet))
                    {
                        pallet.Copy();
                        _SheetUrv11.Range[_SheetUrv11.Cells[rowPaste, 2], _SheetUrv11.Cells[rowPaste, lastCol]].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    }
                    ix = 0;
                    foreach (OfferAddress address in addresses)
                    {
                        ix++;
                        if (ix > 4)
                        {
                            Excel.Range rngTitle = _SheetUrv11.Range["I10:K11"];
                            rngTitle.Copy();
                            _SheetUrv11.Cells[10, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
                            int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv11, "СУММА ПРЕДЛОЖЕННЫХ ОПТИМИЗАЦИЙ (с НДС)").Row;
                            rngTitle = _SheetUrv11.Range[$"I{rowBottomTotal}:J{rowBottomTotal}"];
                            _SheetUrv11.Cells[rowBottomTotal, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
                        }
                        // Вывод и форматирование значений
                        _SheetUrv11.Cells[rowPaste, colPaste].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, address.ColTotalCost].Address}";
                        ///_SheetUrv11.Cells[rowPaste, colPaste].Value = _SheetUrv12.Cells[row, address.ColTotalCost].Value?.ToString() ?? "";

                        string letterOutTotalDiff = ExcelHelper.GetColumnLetter(_SheetUrv12.Cells[rowPaste, colPaste]);
                        _SheetUrv11.Cells[rowPaste, colPaste + 1].Formula = $"=${letterOutTotalDiff}{rowPaste}/$G{rowPaste}-1";
                        _SheetUrv11.Cells[rowPaste, colPaste + 1].NumberFormat = "0%";
                        projectWorkbook.ColorCell(_SheetUrv11.Cells[rowPaste, colPaste + 1], levelNum.ToString());
                        //_SheetUrv11.Cells[rowPaste, colPaste + 1].Value = _SheetUrv12.Cells[row, address.ColPercentTotal].Value?.ToString() ?? "";
                        //projectWorkbook.ColorCell(_SheetUrv11.Cells[rowPaste, colPaste + 1], levelNum.ToString());

                        _SheetUrv11.Cells[rowPaste, colPaste + 2].Formula = $"='{_SheetUrv12.Name}'!{_SheetUrv12.Cells[row, address.ColComments].Address}";
                        //_SheetUrv11.Cells[rowPaste, colPaste + 2].Value = _SheetUrv12.Cells[row, address.ColComments].Value?.ToString() ?? "";
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
        }

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

            int count = 10;
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
                        ColComments = col + 4
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

        private int GetLastColumnUrv(Excel.Worksheet sh, int row)
        {
            int lastCol = sh.Cells[row, sh.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            return lastCol;
        }



        internal void UpdateUrv11(IProgressBarWithLogUI pb)
        {
            pb.SetMainBarVolum(2);
            pb.MainBarTick("Обновление \"Урв 11\"");
            int lastRowSh12 = GetLastRowUrv12();
            int lastRowSh11 = GetLastRowUrv11();
            int ix = 0;
            int colPaste = 9;

            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);

            List<OfferAddress> addresses = GetAdderssLvl12();
            pb.SetSubBarVolume(addresses.Count);

            foreach (OfferAddress address in addresses)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();
                ix++;
                if (ix > 4)
                {
                    Excel.Range rngTitle = _SheetUrv11.Range["I10:K11"];
                    rngTitle.Copy();
                    _SheetUrv11.Cells[10, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
                    int rowBottomTotal = ExcelHelper.FindCell(_SheetUrv11, "СУММА ПРЕДЛОЖЕННЫХ ОПТИМИЗАЦИЙ (с НДС)").Row;
                    rngTitle = _SheetUrv11.Range[$"I{rowBottomTotal}:J{rowBottomTotal}"];
                    rngTitle.Copy();
                    _SheetUrv11.Cells[rowBottomTotal, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
                }

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
                                pallet.Copy();
                                _SheetUrv11.Range[_SheetUrv11.Cells[rowPaste, colPaste], _SheetUrv11.Cells[rowPaste, colPaste + 2]].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
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




    //internal void UpdateUrv12_1(IProgressBarWithLogUI pb)
    //{
    //    string letterNumber = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
    //    string letterLevel = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;

    //    int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
    //    List<OfferAddress> addresses = new ProjectWorkbook().OfferAddress;
    //    int rowPaste = 14;
    //    int colPaste = 6;
    //    for (int row = _project.RowStart; row <= lastRow; row++)
    //    {
    //        string number = _AnalisysSheet.Range[$"${letterNumber}{row}"].Value?.ToString() ?? "";
    //        if (string.IsNullOrEmpty(number)) continue;
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
    //    }
    //}
}
