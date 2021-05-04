using ACO.ExcelHelpers;
using ACO.ProjectBook;
using ACO.ProjectManager;
using System.Collections.Generic;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.PivotSheets
{
    /// <summary>
    ///  Загрузка и обновление сводных таблиц
    /// </summary>
    class Pivot
    {
        readonly Excel.Application _app = Globals.ThisAddIn.Application;
        public readonly Excel.Worksheet SheetUrv12;
        public readonly Excel.Worksheet SheetUrv11;
        readonly Excel.Worksheet _SheetPalette;
        readonly Excel.Worksheet _AnalisysSheet;
        readonly ProjectManager.ProjectManager _projectManager;
        readonly Project _project;
        readonly IProgressBarWithLogUI pb;
        private const int _rowStart = 13;

        public Pivot(IProgressBarWithLogUI pb)
        {
            this.pb = pb;
            Excel.Workbook wb = _app.ActiveWorkbook;
            SheetUrv12 = ExcelHelper.GetSheet(wb, "Урв12");
            SheetUrv11 = ExcelHelper.GetSheet(wb, "Урв11");
            _SheetPalette = ExcelHelper.GetSheet(wb, "Палитра");
            _projectManager = new ProjectManager.ProjectManager();
            _project = _projectManager.ActiveProject;
            string analisysSheetName = _project.AnalysisSheetName;
            _AnalisysSheet = ExcelHelper.GetSheet(wb, analisysSheetName);
        }

        private void PasteTitleOffer12(int colPaste)
        {
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int rowCostWorks = ExcelHelper.FindCell(SheetUrv12, "СТОИМОСТЬ НЕКОТОРЫХ РАБОТ").Row;
            Excel.Range rngTitle = SheetUrv12.Range["F10:J12"];
            rngTitle.Copy();
            SheetUrv12.Cells[10, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
            rngTitle = SheetUrv12.Range[$"F{rowBottomTotal - 1}:J{rowCostWorks}"];
            rngTitle.Copy();
            SheetUrv12.Cells[rowBottomTotal - 1, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);

            //Формат строк в столбце
           
            //int lastrow = SheetUrv12.UsedRange.Row + SheetUrv12.UsedRange.Rows.Count - 2;
            
            Excel.Range formatCell = SheetUrv12.Range[$"F{rowCostWorks + 1}"];
            Excel.Range rng = SheetUrv12.Range[SheetUrv12.Cells[rowCostWorks + 1, colPaste], SheetUrv12.Cells[rowCostWorks + 1, colPaste]];
            formatCell.Copy();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            //СТОИМОСТЬ НЕКОТОРЫХ РАБОТ
            formatCell = SheetUrv12.Range[$"G{rowCostWorks + 1}"];
            rng = SheetUrv12.Range[SheetUrv12.Cells[rowCostWorks + 1, colPaste], SheetUrv12.Cells[rowCostWorks + 1, colPaste + 4]];
            formatCell.Copy();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        }

        private void PrintTitlesOffers12(List<OfferColumns> addresses)
        {
            int i = 0;
            foreach (OfferColumns offer in addresses)
            {
                int colPaste = 6 + 5 * i;
                Excel.Range cell = SheetUrv12.Cells[10, colPaste];
                if (cell.Value is null)
                {
                    PasteTitleOffer12(colPaste);
                    string headerName = cell.Value?.ToString() ?? "";
                    cell.Value = headerName.Replace("УЧАСТНИК 1", offer.ParticipantName);
                }
                i++;
            }
        }

        /// Добавить список
        /// Добавить столбцы КП
        /// Проставить формулы
        public void LoadUrv12()
        {
            pb.SetMainBarVolum(7);
            pb.MainBarTick("Очистка");
            ClearDataRng12();
            string letterName = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Name]).ColumnSymbol;
            string letterNumber = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            string letterLevel = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            string letterCost = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol;

            int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
            pb.MainBarTick("Определение столбцов КП");
            List<OfferColumns> addresses = new ProjectWorkbook().OfferColumns;

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
                    SheetUrv12.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    Excel.Range numberCell = SheetUrv12.Cells[rowPaste, 2];
                    numberCell.NumberFormat = "@";
                    numberCell.Value = number;
                    SheetUrv12.Cells[rowPaste, 3].Value = name;
                    string letterOutName = ExcelHelper.GetColumnLetter(numberCell);
                    int colTotalCost = ExcelHelper.GetColumn(letterCost, _AnalisysSheet);
                    int column = colTotalCost - columnCellNumber + 1;
                    SheetUrv12.Cells[rowPaste, 4].Formula =
                         $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {column}, FALSE)";

                    // Формат строки по уровню
                    Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);
                    if (pallets.TryGetValue(level, out Excel.Range pallet))
                    {
                        ExcelHelper.SetCellFormat(SheetUrv12.Range[SheetUrv12.Cells[rowPaste, 2], SheetUrv12.Cells[rowPaste, lastCol]], pallet);
                    }
                    // Вывод и форматирование значений
                    foreach (OfferColumns address in addresses)
                    {
                        int col = address.ColCostTotalOffer - columnCellNumber + 1;
                        //РУБ. РФ
                        SheetUrv12.Cells[rowPaste, colPaste].Formula =
                           $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                        //% отклонения материалы
                        col = address.ColDeviationMaterials - columnCellNumber + 1;
                        SheetUrv12.Cells[rowPaste, colPaste + 1].Formula =
                          $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";

                        //% отклонения работы
                        col = address.ColDeviationWorks - columnCellNumber + 1;
                        SheetUrv12.Cells[rowPaste, colPaste + 2].Formula =
                          $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";

                        // % отклонения всего
                        string letterOutTotalDiff = ExcelHelper.GetColumnLetter(SheetUrv12.Cells[rowPaste, colPaste]);
                        SheetUrv12.Cells[rowPaste, colPaste + 3].Formula = $"=${letterOutTotalDiff}{rowPaste}/$D{rowPaste}-1";

                        //КОММЕНТАРИИ К СТОИМОСТИ
                        col = address.ColComments - columnCellNumber + 1;
                        SheetUrv12.Cells[rowPaste, colPaste + 4].Formula =
                            $"= VLOOKUP(${letterOutName}{rowPaste}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";

                        colPaste += 5;
                    }
                    colPaste = 6;
                    rowPaste++;
                }
            }

            SetConditionFormat12();
            pb.MainBarTick("Формулы итогов");
            TotalFormuls12();
            pb.MainBarTick("Формат ячеек");
            SetNumberFormat12(addresses.Count);
            pb.MainBarTick("Общие комментарии");
            new OfferInfo(projectWorkbook).SetInfo();

            pb.MainBarTick($"Удаление строки №{_rowStart}");
            Excel.Range rng = SheetUrv12.Cells[_rowStart, 1];
            rng.EntireRow.Delete();

        }

        /// <summary>
        /// Условное форматирование
        /// </summary>
        private void SetConditionFormat12()
        {
            int lastCol = SheetUrv12.UsedRange.Column + SheetUrv12.UsedRange.Columns.Count;
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;

            ConditonsFormatManager formatManager = new ConditonsFormatManager();
            /// Удаление правил
            Excel.Range rng = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, 1], SheetUrv12.Cells[rowBottomTotal, lastCol]];
            ExcelHelper.ClearFormatConditions(rng);
            // Правила для столбца материалы
            string colDeviationMat = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationMat];
            string colDeviationWorks = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationWorks];
            string colDeviationCost = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationCost];
            List<ConditionFormat> conditionsDeviationMat = formatManager.ListConditionFormats.FindAll(a => a.ColumnName == colDeviationMat);
            List<ConditionFormat> conditionsDeviationWorks = formatManager.ListConditionFormats.FindAll(a => a.ColumnName == colDeviationWorks);
            List<ConditionFormat> conditionsDeviationCost = formatManager.ListConditionFormats.FindAll(a => a.ColumnName == colDeviationCost);

            for (int col = 6; col < lastCol; col += 5)
            {
                Excel.Range columnMaterials = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, col + 1], SheetUrv12.Cells[rowBottomTotal - 2, col + 1]];
                conditionsDeviationMat.ForEach(x => x.SetCondition(columnMaterials));

                Excel.Range columnWorks = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, col + 2], SheetUrv12.Cells[rowBottomTotal - 2, col + 2]];
                conditionsDeviationWorks.ForEach(x => x.SetCondition(columnWorks));

                Excel.Range columnCost = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, col + 3], SheetUrv12.Cells[rowBottomTotal - 2, col + 3]];
                conditionsDeviationCost.ForEach(x => x.SetCondition(columnCost));
            }
        }

        private void SetConditionFormat11()
        {
            int lastCol = SheetUrv11.UsedRange.Column + SheetUrv11.UsedRange.Columns.Count;
            int rowBottomTotal = GetLastRowUrv11();
            // ExcelHelper.FindCell(SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;

            ConditonsFormatManager formatManager = new ConditonsFormatManager();
            /// Удаление правил
            Excel.Range rng = SheetUrv11.Range[SheetUrv11.Cells[_rowStart, 1], SheetUrv11.Cells[rowBottomTotal, lastCol]];
            ExcelHelper.ClearFormatConditions(rng);
            // Правила для столбца материалы           
            string colDeviationCost = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationCost];

            List<ConditionFormat> conditionsDeviationCost = formatManager.ListConditionFormats.FindAll(a => a.ColumnName == colDeviationCost);
            for (int col = 9; col <= lastCol; col += 3)
            {
                Excel.Range columnCost = SheetUrv11.Range[SheetUrv11.Cells[_rowStart, col + 1], SheetUrv11.Cells[rowBottomTotal, col + 1]];
                conditionsDeviationCost.ForEach(x => x.SetCondition(columnCost));
            }
        }

        /// <summary>
        ///  Формат 
        /// </summary>
        /// <param name="addresses"></param>
        private void SetNumberFormat12(int addressesCount)
        {
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal + 2;
            // int colPaste6 = 6;
            Excel.Range rng = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, 4], SheetUrv12.Cells[lastRow, 4]];
            rng.NumberFormat = "#,##0,00";

            int lastCol = 6 + 5 * addressesCount - 1;
            for (int col = 6; col <= lastCol; col += 5)
            {
                rng = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, col], SheetUrv12.Cells[lastRow, col]];
                rng.NumberFormat = "#,##0,00";
                rng = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, col + 1], SheetUrv12.Cells[lastRow, col + 1]];
                rng.NumberFormat = "0%";
                rng = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, col + 2], SheetUrv12.Cells[lastRow, col + 2]];
                rng.NumberFormat = "0%";
                rng = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, col + 3], SheetUrv12.Cells[lastRow, col + 3]];
                rng.NumberFormat = "0%";
            }
        }
        private void SetNumberFormat11(int addressesCount)
        {
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv11, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal + 2;
            Excel.Range rng = SheetUrv11.Range[SheetUrv11.Cells[_rowStart, 7], SheetUrv11.Cells[lastRow, 7]];
            rng.NumberFormat = "#,##0,00";

            int lastCol = 6 + 5 * addressesCount - 1;
            for (int col = 9; col <= lastCol; col += 3)
            {
                rng = SheetUrv11.Range[SheetUrv11.Cells[_rowStart, col], SheetUrv11.Cells[lastRow, col]];
                rng.NumberFormat = "#,##0,00";
                rng = SheetUrv11.Range[SheetUrv11.Cells[_rowStart, col + 1], SheetUrv11.Cells[lastRow, col + 1]];
                rng.NumberFormat = "0%";
            }
        }

        /// <summary>
        ///  Обновление значений урв 12
        /// </summary>
        /// <param name="pb"></param>
        internal void UpdateUrv12()
        {
            pb.SetMainBarVolum(4);
            pb.MainBarTick("Обновление формул \"Урв 12\"");
            int colPaste = 1;

            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            Excel.Range dataRange = projectWorkbook.GetAnalysisRange();
            List<OfferColumns> addresses = new ProjectWorkbook().OfferColumns;
            pb.Writeline("Копирование заголовков");
            PrintTitlesOffers12(addresses);

            string letterNumber = projectWorkbook.GetLetter(StaticColumns.Number);
            Excel.Range cellNumber = _AnalisysSheet.Range[$"${letterNumber}{_project.RowStart}"];
            int columnCellNumber = cellNumber.Column;
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal - 2;

            pb.SetSubBarVolume(addresses.Count);
            //// Вывод и форматирование значений
            foreach (OfferColumns address in addresses)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();
                colPaste += 5;
                pb.Writeline($"Формулы: {address.ParticipantName}");
                string textCost = SheetUrv12.Cells[_rowStart, colPaste].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(textCost)) continue;  // Пропустить заполненные КП

                string formulaSumm = "";
                for (int row = _rowStart; row <= lastRow; row++)
                {
                    string number = SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number)) continue;
                    number = number.Trim(new char[] { ' ', '.' });
                    int levelNum = number.Split('.').Length;
                    if (levelNum == 1)
                    {
                        formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={SheetUrv12.Cells[row, colPaste].Address}" :
                                                                                $"+{SheetUrv12.Cells[row, colPaste].Address}";
                    }
                    // Формат строки по уровню
                    Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);
                    string keyLvl = levelNum.ToString();
                    if (pallets.TryGetValue(keyLvl, out Excel.Range pallet))
                    {
                        ExcelHelper.SetCellFormat(SheetUrv12.Range[SheetUrv12.Cells[row, colPaste], SheetUrv12.Cells[row, colPaste + 4]], pallet);
                    }
                }

                int col = address.ColDeviationCost - columnCellNumber + 1;
                //РУБ. РФ
                SheetUrv12.Cells[_rowStart, colPaste].Formula =
                   $"= VLOOKUP($B{_rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                //% отклонения материалы
                col = address.ColDeviationMaterials - columnCellNumber + 1;
                SheetUrv12.Cells[_rowStart, colPaste + 1].Formula =
                  $"= VLOOKUP($B{_rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                //% отклонения работы
                col = address.ColDeviationWorks - columnCellNumber + 1;
                SheetUrv12.Cells[_rowStart, colPaste + 2].Formula =
                  $"= VLOOKUP($B{_rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";
                // % отклонения всего
                string letterOutTotalDiff = ExcelHelper.GetColumnLetter(SheetUrv12.Cells[_rowStart, colPaste]);
                SheetUrv12.Cells[_rowStart, colPaste + 3].Formula = $"=${letterOutTotalDiff}{_rowStart}/$D{_rowStart}-1";
                //КОММЕНТАРИИ К СТОИМОСТИ
                col = address.ColComments - columnCellNumber + 1;
                SheetUrv12.Cells[_rowStart, colPaste + 4].Formula =
                    $"= VLOOKUP($B{_rowStart}, '{_project.AnalysisSheetName}'! {dataRange.Address}, {col}, FALSE)";

                Excel.Range rng = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, colPaste], SheetUrv12.Cells[_rowStart, colPaste + 4]];
                Excel.Range destination = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, colPaste], SheetUrv12.Cells[lastRow, colPaste + 4]];
                rng.AutoFill(destination, Excel.XlAutoFillType.xlFillValues);
                destination.Columns[2].NumberFormat = "0%";
                destination.Columns[3].NumberFormat = "0%";
                destination.Columns[4].NumberFormat = "0%";

                if (!string.IsNullOrEmpty(formulaSumm))
                {
                    SheetUrv12.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
                    SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Formula =
                                        $"={SheetUrv12.Cells[rowBottomTotal, colPaste].Address}*0.2";
                    SheetUrv12.Cells[rowBottomTotal + 2, colPaste].Formula =
                                        $"={SheetUrv12.Cells[rowBottomTotal, colPaste].Address}+" +
                                        $"{SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Address}";
                }
                //TODO подсчитать кол-во.
                // PrintTotalComments(address);
            }
            pb.Writeline("Условное форматирование");
            SetConditionFormat12();
            pb.MainBarTick("Формулы итогов");
            TotalFormuls12();
            pb.MainBarTick("Формат ячеек");
            SetNumberFormat12(addresses.Count);
            pb.MainBarTick("Общие комментарии");
            new OfferInfo(projectWorkbook).SetInfo();
            //TODO загрузить наиболее дорогии позиции 
        }

        private void PrintTotalComments(OfferColumns address)
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
            List<OfferColumns> addresses = new ProjectWorkbook().OfferColumns;
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal - 2;
            string formulaSumm = "";
            for (int row = _rowStart; row <= lastRow; row++)
            {
                string number = SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;
                number = number.Trim(new char[] { ' ', '.' });
                int levelNum = number.Split('.').Length;

                if (levelNum == 1)
                {
                    formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={SheetUrv12.Cells[row, 4].Address}" :
                                                                         $"+{SheetUrv12.Cells[row, 4].Address}";
                }
            }
            if (!string.IsNullOrEmpty(formulaSumm))
            {
                SheetUrv12.Cells[rowBottomTotal, 4].Formula = formulaSumm;
                SheetUrv12.Cells[rowBottomTotal + 1, 4].Formula =
                                    $"={SheetUrv12.Cells[rowBottomTotal, 4].Address}*0.2";
                SheetUrv12.Cells[rowBottomTotal + 2, 4].Formula =
                                    $"={SheetUrv12.Cells[rowBottomTotal, 4].Address}+" +
                                    $"{SheetUrv12.Cells[rowBottomTotal + 1, 4].Address}";
            }

            int colPaste = 6;
            foreach (OfferColumns address in addresses)
            {
                formulaSumm = "";
                //string formulaSumm
                for (int row = _rowStart; row <= lastRow; row++)
                {
                    string number = SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number)) continue;
                    number = number.Trim(new char[] { ' ', '.' });
                    int levelNum = number.Split('.').Length;
                    if (levelNum == 1)
                    {
                        formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={SheetUrv12.Cells[row, colPaste].Address}" :
                                                                                $"+{SheetUrv12.Cells[row, colPaste].Address}";
                    }
                }
                if (!string.IsNullOrEmpty(formulaSumm))
                {
                    SheetUrv12.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
                    SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Formula =
                                        $"={SheetUrv12.Cells[rowBottomTotal, colPaste].Address}*0.2";
                    SheetUrv12.Cells[rowBottomTotal + 2, colPaste].Formula =
                                        $"={SheetUrv12.Cells[rowBottomTotal, colPaste].Address}+" +
                                        $"{SheetUrv12.Cells[rowBottomTotal + 1, colPaste].Address}";
                    colPaste += 5;
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void TotalFormuls11()
        {
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv11, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            int lastRow = rowBottomTotal - 2;
            int colPaste = 7;
            string formulaSumm = "";
            for (int row = _rowStart; row <= lastRow; row++)
            {

                string number = SheetUrv11.Cells[row, 2].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;
                number = number.Trim(new char[] { ' ', '.' });
                int levelNum = number.Split('.').Length;

                if (levelNum == 1)
                {
                    formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={SheetUrv11.Cells[row, colPaste].Address}" :
                                                                         $"+{SheetUrv11.Cells[row, colPaste].Address}";
                }
            }
            if (!string.IsNullOrEmpty(formulaSumm))
            {
              SheetUrv11.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
            }
            SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Formula =
                 $"={SheetUrv11.Cells[rowBottomTotal, colPaste].Address}*0.2";
            SheetUrv11.Cells[rowBottomTotal + 2, colPaste].Formula =
                                $"={SheetUrv11.Cells[rowBottomTotal, colPaste].Address}+" +
                                $"{SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Address}";

            List<OfferColumns> addresses = GetAdderssLvl12();

            colPaste = 9;
            foreach (OfferColumns address in addresses)
            {
                formulaSumm = "";
                for (int row = _rowStart; row <= lastRow; row++)
                {
                    string number = SheetUrv11.Cells[row, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number)) continue;
                    number = number.Trim(new char[] { ' ', '.' });
                    int levelNum = number.Split('.').Length;

                    if (levelNum == 1)
                    {
                        formulaSumm += (string.IsNullOrEmpty(formulaSumm)) ? $"={SheetUrv11.Cells[row, colPaste].Address}" :
                                                                                $"+{SheetUrv11.Cells[row, colPaste].Address}";
                    }
                }
                if (!string.IsNullOrEmpty(formulaSumm))
                {
                    SheetUrv11.Cells[rowBottomTotal, colPaste].Formula = formulaSumm;
                }
                SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Formula =
                                    $"={SheetUrv11.Cells[rowBottomTotal, colPaste].Address}*0.2";

                SheetUrv11.Cells[rowBottomTotal + 2, colPaste].Formula =
                                    $"={SheetUrv11.Cells[rowBottomTotal, colPaste].Address}+" +
                                    $"{SheetUrv11.Cells[rowBottomTotal + 1, colPaste].Address}";
                // SheetUrv11.Columns[colPaste].NumberFormat = "#,##0.00";
                colPaste += 3;
            }
        }

        private int GetLastRowUrv12()
        {
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            return rowBottomTotal - 2;
        }
        private int GetLastRowUrv11()
        {
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv11, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            return rowBottomTotal - 2;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="colPaste"></param>
        private void PasteTitleOffer11(int colPaste)
        {
            int rowBottomTotal = ExcelHelper.FindCell(SheetUrv11, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            Excel.Range rngTitle = SheetUrv11.Range["I10:K12"];
            rngTitle.Copy();
            SheetUrv11.Cells[10, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
            rngTitle = SheetUrv11.Range[$"I{_rowStart}:K{rowBottomTotal + 6}"];
            rngTitle.Copy();
            SheetUrv11.Cells[_rowStart, colPaste].PasteSpecial(Excel.XlPasteType.xlPasteAll);
            Excel.Range rng = SheetUrv11.Range[SheetUrv11.Cells[_rowStart, colPaste],
                                               SheetUrv11.Cells[rowBottomTotal - 1, colPaste + 2]];
            rng.ClearContents();
        }

        private void PrintTitlesOffers11(List<OfferColumns> addresses)
        {
            int i = 0;
            foreach (OfferColumns offer in addresses)
            {
                int colPaste = 9 + 3 * i;
                Excel.Range cell = SheetUrv11.Cells[10, colPaste];
                if (cell.Value is null)
                {
                    PasteTitleOffer11(colPaste);
                    string headerName = cell.Value?.ToString() ?? "";
                    cell.Value = offer.ParticipantName;
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
           // ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            List<OfferColumns> addresses = GetAdderssLvl12();
            pb.Writeline("Копирование заголовков");
            PrintTitlesOffers11(addresses);

            int offersCount = addresses.Count;
            int rowPaste = 14;
            int colPaste = 9;
            int lastCol = colPaste + 3 * offersCount - 1;
            int count = lastRow - 12;
            if (count < 1) throw new AddInException($"Строки отсутствуют лист: {SheetUrv12.Name}");
            pb.SetSubBarVolume(count);
            pb.MainBarTick("Печать строк");


            for (int row = _rowStart; row <= lastRow; row++)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();
                string number = SheetUrv12.Cells[row, 2].Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(number)) continue;

                string name = SheetUrv12.Cells[row, 3].Value?.ToString() ?? "";

                number = number.Trim(new char[] { ' ', '.' });
                int levelNum = number.Split('.').Length;

                if (levelNum > 0 && levelNum < 3)
                {

                    string outRowNumber = SheetUrv11.Cells[rowPaste, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(outRowNumber))
                    {
                        SheetUrv11.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    }
                    Excel.Range numberCell = SheetUrv11.Cells[rowPaste, 2];
                    numberCell.NumberFormat = "@";
                    numberCell.Value = number;

                    SheetUrv11.Cells[rowPaste, 3].Value = name;
                    SheetUrv11.Cells[rowPaste, 7].Formula = $"='{SheetUrv12.Name}'!{SheetUrv12.Cells[row, 4].Address}";

                    // Формат строки по уровню
                    if (pallets.TryGetValue(levelNum.ToString(), out Excel.Range pallet))
                    {
                        ExcelHelper.SetCellFormat(SheetUrv11.Range[SheetUrv11.Cells[rowPaste, 2],
                                                  SheetUrv11.Cells[rowPaste, lastCol]], pallet);
                    }

                    foreach (OfferColumns  address in addresses)
                    {
                        PrintValuesFormuls11(address, row, rowPaste, colPaste);
                        colPaste += 3;
                    }
                    colPaste = 9;
                    rowPaste++;
                }
            }
            pb.Writeline("Формат ячеек");
            SetConditionFormat11();
            SetNumberFormat11(offersCount);

            pb.Writeline("Формулы итогов");
            TotalFormuls11();
            if (rowPaste > 14)
            {
                pb.MainBarTick($"Удаление стр №{_rowStart}");
                Excel.Range rng = SheetUrv11.Cells[_rowStart, 1];
                rng.EntireRow.Delete();
            }

            pb.MainBarTick("Обновление диаграммы");
            UpdateDiagramm();
        }

        /// <summary>
        /// 
        /// </summary>
        internal void UpdateUrv11()
        {
            pb.SetMainBarVolum(2);
            pb.MainBarTick("Обновление \"Урв 11\"");
            int lastRowSh12 = GetLastRowUrv12();
            int lastRowSh11 = GetLastRowUrv11();
            int colPaste = 9;

           // ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPalette);

            List<OfferColumns> addresses = GetAdderssLvl12();
            pb.Writeline("Копирование заголовков");
            PrintTitlesOffers11(addresses);
            pb.SetSubBarVolume(addresses.Count);

            foreach (OfferColumns address in addresses)
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.SubBarTick();
                for (int rowPaste = _rowStart; rowPaste <= lastRowSh11; rowPaste++)
                {
                    string number11 = SheetUrv11.Cells[rowPaste, 2].Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(number11)) continue;

                    for (int row = _rowStart; row <= lastRowSh12; row++)
                    {
                        string number12 = SheetUrv11.Cells[row, 2].Value?.ToString() ?? "";
                        if (number12 != number11) continue;
                        number11 = number11.Trim(new char[] { ' ', '.' });
                        int levelNum = number11.Split('.').Length;

                        if (levelNum > 0 && levelNum < 3)
                        {
                            // Формат строки по уровню
                            if (pallets.TryGetValue(levelNum.ToString(), out Excel.Range pallet))
                            {
                                ExcelHelper.SetCellFormat(SheetUrv11.Range[SheetUrv11.Cells[rowPaste, colPaste], SheetUrv11.Cells[rowPaste, colPaste + 2]], pallet);
                            }
                            // Вывод и форматирование значений
                            Excel.Range rng = SheetUrv11.Cells[rowPaste, colPaste];
                            rng.Formula = $"='{SheetUrv12.Name}'!{SheetUrv12.Cells[row, address.ColCostTotalOffer].Address}";

                            string letterOutTotalDiff = ExcelHelper.GetColumnLetter(SheetUrv12.Cells[row, colPaste]);
                            SheetUrv11.Cells[rowPaste, colPaste + 1].Formula = $"=${letterOutTotalDiff}{rowPaste}/$G{rowPaste}-1";
                            SheetUrv11.Cells[rowPaste, colPaste + 1].NumberFormat = "0%";
                            // projectWorkbook.ColorCell(SheetUrv11.Cells[rowPaste, colPaste + 1], levelNum.ToString());
                            SheetUrv11.Cells[rowPaste, colPaste + 2].Formula = $"='{SheetUrv12.Name}'!{SheetUrv12.Cells[row, address.ColComments].Address}";
                        }
                    }
                }
                colPaste += 3;
            }
            pb.Writeline("Формулы итогов");
            TotalFormuls11();
            pb.Writeline("Формат ячеек");
            SetConditionFormat11();
            SetNumberFormat11(addresses.Count);
            pb.MainBarTick("Обновление диаграммы");
            UpdateDiagramm();
        }


        /// <summary>
        /// 
        /// </summary>
        private void PrintValuesFormuls11(OfferColumns address, int row, int rowPaste, int colPaste)
        {
            // Вывод и форматирование значений
            SheetUrv11.Cells[rowPaste, colPaste].Formula = $"='{SheetUrv12.Name}'!{SheetUrv12.Cells[row, address.ColCostTotalOffer].Address}";
            string letterOutTotalDiff = ExcelHelper.GetColumnLetter(SheetUrv12.Cells[rowPaste, colPaste]);
            SheetUrv11.Cells[rowPaste, colPaste + 1].Formula = $"=${letterOutTotalDiff}{rowPaste}/$G{rowPaste}-1";
            SheetUrv11.Cells[rowPaste, colPaste + 1].NumberFormat = "0%";
            SheetUrv11.Cells[rowPaste, colPaste + 2].Formula = $"='{SheetUrv12.Name}'!{SheetUrv12.Cells[row, address.ColComments].Address}";
        }

        /// <summary>
        ///  Номера столбцов заполненных кп
        /// </summary>
        /// <returns></returns>
        private List<OfferColumns> GetAdderssLvl12()
        {
            List<OfferColumns> addresses = new List<OfferColumns>();
            int lastCol = GetLastColumnUrv(SheetUrv12, _rowStart);

            for (int col = 6; col <= lastCol; col += 5)
            {
                string val = SheetUrv12.Cells[_rowStart, col].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(val))
                {
                    OfferColumns address = new OfferColumns
                    {
                        ColCostTotalOffer = col,
                        ColDeviationCost = col + 3,
                        ColComments = col + 4,
                        ParticipantName = SheetUrv12.Cells[10, col].Value?.ToString() ?? ""
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
            int lastColumn = SheetUrv12.UsedRange.Column + SheetUrv12.UsedRange.Columns.Count - 1;
            Excel.Range dataRng = SheetUrv12.Range[SheetUrv12.Cells[14, 2], SheetUrv12.Cells[lastRow, lastColumn]];
            dataRng.EntireRow.Delete();
            dataRng = SheetUrv12.Range[SheetUrv12.Cells[_rowStart, 2], SheetUrv12.Cells[_rowStart, lastColumn]];
            dataRng.ClearContents();
            return;
        }
        private void ClearDataRng11()
        {
            int lastRow = GetLastRowUrv11();
            if (lastRow <= 14) return;
            int lastColumn = SheetUrv11.UsedRange.Column + SheetUrv11.UsedRange.Columns.Count - 1;
            Excel.Range dataRng = SheetUrv11.Range[SheetUrv11.Cells[14, 2], SheetUrv11.Cells[lastRow, lastColumn]];
            dataRng.EntireRow.Delete();
            dataRng = SheetUrv11.Range[SheetUrv11.Cells[_rowStart, 2], SheetUrv11.Cells[_rowStart, lastColumn]];
            dataRng.ClearContents();
            return;
        }

        private int GetLastColumnUrv(Excel.Worksheet sh, int row)
        {
            int lastCol = sh.Cells[row, sh.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            return lastCol;
        }

        /// <summary>
        ///  Обновление Диаграммы
        /// </summary>
        internal void UpdateDiagramm()
        {
            Excel.ChartObject shp = SheetUrv11.ChartObjects("Chart 2");
            Excel.Chart chartPage = shp.Chart;
            Excel.SeriesCollection seriesCol = (Excel.SeriesCollection)chartPage.SeriesCollection();
            Excel.FullSeriesCollection fullColl = chartPage.FullSeriesCollection();
            Debug.WriteLine(fullColl.Count);
            int lastCol = GetLastColumnUrv(SheetUrv11, _rowStart);
            int lastRow = GetLastRowUrv11();
            int ix = 1;
            string letterCost = "G";
            fullColl.Item(ix).Name = $"={SheetUrv11.Name}!${letterCost}10";
            fullColl.Item(ix).Values = $"={SheetUrv11.Name}!${letterCost}{_rowStart}:${letterCost}{lastRow}";
            fullColl.Item(ix).XValues = $"={SheetUrv11.Name}!$C{_rowStart}:$C{lastRow}";

            for (int col = 9; col <= lastCol; col += 3)
            {
                Excel.Range cellFirstCost = SheetUrv11.Cells[_rowStart, col];
                string text = cellFirstCost.Value?.ToString() ?? "";
                if (string.IsNullOrEmpty(text)) continue;
                letterCost = ExcelHelper.GetColumnLetter(cellFirstCost);
                ix++;
                if (ix > fullColl.Count)
                {
                    seriesCol.NewSeries();
                }
                fullColl.Item(ix).Name = $"={SheetUrv11.Name}!${letterCost}10";
                fullColl.Item(ix).Values = $"={SheetUrv11.Name}!${letterCost}{_rowStart}:${letterCost}{lastRow}";
                fullColl.Item(ix).XValues = $"={SheetUrv11.Name}!$C{_rowStart}:$C{lastRow}";
            }
            if (ix < fullColl.Count)
            {
                for (int i = ix + 1; i <= fullColl.Count; i++)
                {
                    fullColl.Item(i).Delete();
                }
            }
        }

    }
}
