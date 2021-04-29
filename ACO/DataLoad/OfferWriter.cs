using Excel = Microsoft.Office.Interop.Excel;
using ACO.ExcelHelpers;
using ACO.ProjectManager;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;
using System;
using System.Drawing;

namespace ACO

{
    class OfferWriter
    {
        readonly Excel.Application _app = default;

        /// <summary>
        ///  Текущая книга с проектом
        /// </summary>
        readonly Excel.Workbook _wb = null;

        /// <summary>
        ///  Книга КП
        /// </summary>
        // ExcelFile
        readonly Excel.Workbook _offerBook = null;

        /// <summary>
        /// Лист  Анализ
        /// </summary>
        Excel.Worksheet _sheetProject = null;
        readonly OfferManager _offerManager = null;
        readonly Project _CurrentProject = null;

        public OfferWriter(ExcelFile offerBook)
        {
            _app = Globals.ThisAddIn.Application;
            _wb = _app.ActiveWorkbook;
            _offerBook = offerBook.WorkBook;
            _offerManager = new OfferManager();
            _CurrentProject = new ProjectManager.ProjectManager().ActiveProject;
            // Лист анализ в текущем проекте
            _sheetProject = ExcelHelper.GetSheet(_wb, _CurrentProject.AnalysisSheetName);
            _CurrentProject.SetColumnNumbers(_sheetProject);
        }


        /// <summary>
        /// Печать КП
        /// </summary>
        /// <param name="offer"></param>
        internal void Print(IProgressBarWithLogUI pb, string offerSettingsName)
        {
            // Ищем настройки столбцов
            OfferSettings offerSettings = _offerManager.Mappings.Find(s => s.Name == offerSettingsName);
            pb.Writeline($"Выбор листа {offerSettings.SheetName}");
            // Лист данных КП

            Excel.Worksheet offerSheet = ExcelHelper.GetSheet(_offerBook, offerSettings.SheetName);
            pb.Writeline("Разгруппировка строк");
            ShowSheetRows(offerSheet);

            ListAnalysis SheetAnalysis = new ListAnalysis(_sheetProject, _CurrentProject);

            pb.Writeline("Адресация полей");
            /// Адресация полей КП
            List<FieldAddress> addresslist = GetFields(offerSettings, SheetAnalysis.ColumnStartPrint);

            Excel.Worksheet tamplateSheet = ExcelHelper.GetSheet(_wb, "Шаблоны");
            pb.Writeline("Печать заголовков");
            SheetAnalysis.PrintTitle(tamplateSheet, addresslist);


            int lastRowOffer = offerSheet.UsedRange.Row + offerSheet.UsedRange.Rows.Count - 1;
            pb.Writeline("Чтение массива данных");
            // Массив загружаемых данных
            object[,] arrData = GetArrData(offerSheet, offerSettings.RowStart, lastRowOffer);

            int countRows = lastRowOffer - offerSettings.RowStart + 1;
            pb.SetSubBarVolume(countRows);
            pb.Writeline("Вывод строк");
            for (int i = 1; i <= countRows; i++)
            {
                pb.SubBarTick();
                if (pb.IsAborted) throw new AddInException("Процесс остановлен.");

                Record offerRecord = new Record
                {
                    Addresslist = addresslist
                };
                // Сбор данных
                foreach (FieldAddress field in addresslist)
                {
                    object val = arrData[i, field.ColumnOffer];
                    string text = val?.ToString() ?? "";

                    offerRecord.Values.Add(field.ColumnPaste, val);
                    if (field.MappingAnalysis.Name == Project.ColumnsNames[StaticColumns.Level])
                    {
                        offerRecord.Level = int.TryParse(text, out int lvl) ? lvl : 0;
                    }
                    if (field.MappingAnalysis.Name == Project.ColumnsNames[StaticColumns.Number])
                    {
                        offerRecord.Number = text;
                    }
                    if (field.MappingAnalysis.Check)
                    {
                        offerRecord.KeyFilds.Add(text);
                    }
                }
                SheetAnalysis.PrintRecord(offerRecord);

            }
            pb.Writeline("Группировка столбцов");
            SheetAnalysis.GroupColumn();
            if (pb.IsAborted) throw new AddInException("Процесс остановлен.");
            pb.Writeline("Формулы \"Комментарии Спектрум к заявке участника\"");
            SetFormuls();
        }



        /// <summary>
        ///  Комментарии Спектрум  
        /// </summary>
        private void SetFormuls()
        {
            //OfferSettings offerSettings
            int rowStart = _CurrentProject.RowStart;
            int colStart = GetColumnStartFormuls();
            Dictionary<string, string> columnsOffer = GetColumnsOffer();
            string GetLetter(Dictionary<string, string> columns, string name)
            {
                if (columns.ContainsKey(name))
                {
                    return columns[name];
                }
                throw new AddInException($"Столбец не найден: \"{name}\".");
            }
            Excel.Range cell = _sheetProject.Cells[rowStart, colStart];
            string letterNameSpectrum = _CurrentProject.GetColumn(StaticColumns.Name).ColumnSymbol;
            string letterNameOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.Name]);
            string letterAmountSpectrum = _CurrentProject.GetColumn(StaticColumns.Amount).ColumnSymbol;
            string letterAmountOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.Amount]);

            //Наименование вида работ
            cell.Formula = $"=${letterNameSpectrum}{rowStart}=${letterNameOffer}{rowStart}";
            //Комментарии Спектрум к описанию работ
            string letterCheckName = cell.Address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
            _sheetProject.Cells[rowStart, colStart + 1].Formula = $"=IF(${letterCheckName}{rowStart}=TRUE,\".\",Комментарии!$A$2)";
            //Отклонение по объемам
            _sheetProject.Cells[rowStart, colStart + 2].Formula = $"=IFERROR({letterAmountSpectrum}{rowStart}/{letterAmountOffer}{rowStart}-1,\"#НД\")";
            //Комментарии Спектрум к объемам работ
            cell = _sheetProject.Cells[rowStart, colStart + 2];
            string address = cell.Address;
            cell = _sheetProject.Cells[rowStart, colStart + 3];
            string letter = address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
            cell.Formula = $"=IF(${letter}{rowStart}>15%,Комментарии!$A$5,IF(${letter}{rowStart}<-15%,Комментарии!$A$6,\".\"))";

            //Отклонение по стоимости
            string letterTotalOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.CostTotal]);
            string letterTotalSpectrum = _CurrentProject.GetColumn(StaticColumns.CostTotal).ColumnSymbol;
            _sheetProject.Cells[rowStart, colStart + 4].Formula =
                        $"=IFERROR(IF(${letterTotalSpectrum}{rowStart}<>0," +
                        $"${letterTotalOffer}{rowStart}/${letterTotalSpectrum}{rowStart}-1,0),\"#НД\")";

            //Комментарии Спектрум к стоимости работ
            cell = _sheetProject.Cells[rowStart, colStart + 4];
            address = cell.Address;
            string letterDiffCost = address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
            _sheetProject.Cells[rowStart, colStart + 5].Formula =
                $"=IF(${letterDiffCost}{rowStart}>15%,Комментарии!$A$9,IF(${letterDiffCost}{rowStart}<-15%,Комментарии!$A$10,\".\"))";
            //Отклонение по стоимости РАБ
            string letterWorkslOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.CostWorksTotal]);
            string letterWorksSpectrum = _CurrentProject.GetColumn(StaticColumns.CostWorksTotal).ColumnSymbol;
            _sheetProject.Cells[rowStart, colStart + 6].Formula =
                        $"=IFERROR(IF(${letterWorkslOffer}{rowStart}<>0," +
                        $"${letterWorkslOffer}{rowStart}/${letterWorksSpectrum}{rowStart}-1,\"Отс-ет ст-ть мат.\"),\"#НД\")";
            //Отклонение по стоимости МАТ
            string letterMaterialslOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.CostMaterialsTotal]);
            string letterMaterialsSpectrum = _CurrentProject.GetColumn(StaticColumns.CostMaterialsTotal).ColumnSymbol;
            _sheetProject.Cells[rowStart, colStart + 7].Formula =
                        $"=IFERROR(IF(${letterMaterialslOffer}{rowStart}<>0," +
                        $"${letterMaterialslOffer}{rowStart}/${letterMaterialsSpectrum}{rowStart}-1,\"Отс-ет ст-ть работ\"),\"#НД\")";

            //Комментарии к строкам "0"
            _sheetProject.Cells[rowStart, colStart + 8].Formula =
                        $"=IF(${letterDiffCost}{rowStart}=-1,\"Указать стоимость единичной расценки и посчитать итог\",\".\")";


            // Протянуть до конца листов
            Excel.Range rng = _sheetProject.Range[_sheetProject.Cells[rowStart, colStart],
                                                _sheetProject.Cells[rowStart, colStart + 8]];
            int lastRow = GetLastRow(_sheetProject, letterNameOffer);

            if (lastRow > rowStart)
            {
                Excel.Range destination = _sheetProject.Range[_sheetProject.Cells[rowStart, colStart], _sheetProject.Cells[lastRow, colStart + 8]];
                rng.AutoFill(destination);
                destination.Interior.Color = Color.FromArgb(232, 242, 238);
                destination.Columns[3].NumberFormat = "0%";
                destination.Columns[5].NumberFormat = "0%";
                destination.Columns[7].NumberFormat = "0%";
                destination.Columns[8].NumberFormat = "0%";
            }
        }

        private Dictionary<string, string> GetColumnsOffer()
        {
            Dictionary<string, string> columnsOffer = new Dictionary<string, string>();
            int lastCol = _sheetProject.Cells[1, _sheetProject.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

            for (int col = lastCol; lastCol > 1; col--)
            {
                Excel.Range cell = _sheetProject.Cells[1, col];
                string text = cell.Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(text))
                {
                    string letter = cell.Address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
                    columnsOffer.Add(text, letter);

                    if (text == "offer_start")
                    {
                        break;
                    }
                }
            }

            return columnsOffer;
        }

        /// <summary>
        ///  Найти столбец начала комментариев. 
        /// </summary>
        /// <returns></returns>
        private int GetColumnStartFormuls()
        {
            int lastCol = _sheetProject.Cells[1, _sheetProject.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

            for (int col = lastCol; lastCol > 1; col--)
            {
                Excel.Range cell = _sheetProject.Cells[1, col];
                string text = cell.Value?.ToString() ?? "";
                if (text == "offer_end")
                {
                    return col;
                }
            }
            throw new AddInException("Столбец начала формул не найден.");
        }

        private List<FieldAddress> GetFields(OfferSettings offerSettings, int lastCol)
        {
            List<FieldAddress> fields = new List<FieldAddress>();
            int k = 0;
            foreach (OfferColumnMapping columnOffer in offerSettings.Columns)
            {
                if (string.IsNullOrEmpty(columnOffer.ColumnSymbol)) continue;
                ColumnMapping сolumnProject = _CurrentProject.Columns.Find(a => a.Name == columnOffer.Name);

                if (сolumnProject.Obligatory)
                {
                    сolumnProject.Column = ExcelHelper.GetColumn(сolumnProject.ColumnSymbol, _sheetProject);
                    int colPaste = lastCol + k;
                    int colOffer = ExcelHelper.GetColumn(columnOffer.ColumnSymbol, _sheetProject);
                    fields.Add(new FieldAddress()
                    {
                        ColumnOffer = colOffer,
                        ColumnPaste = colPaste,
                        MappingAnalysis = сolumnProject
                    });
                    k++;
                }
            }
            return fields;
        }

        /// <summary>
        ///  Диапазон в виде массива
        /// </summary>
        /// <param name="offerSheet"></param>
        /// <param name="rowStart"></param>
        /// <param name="lastRow"></param>
        /// <returns></returns>
        private object[,] GetArrData(Excel.Worksheet offerSheet, int rowStart, int lastRow)
        {
            int lastColumn = offerSheet.UsedRange.Column + offerSheet.UsedRange.Columns.Count - 1;
            Excel.Range RngData = offerSheet.Range[offerSheet.Cells[rowStart, 1], offerSheet.Cells[lastRow, lastColumn]];
            return RngData.Value;
        }

        /// <summary>
        /// Печать КП
        /// </summary>
        /// <param name="offer"></param>
        internal void PrintBaseEstimate(IProgressBarWithLogUI pb, string offerSettingsName)
        {
            OfferSettings offerSettings = _offerManager.Mappings.Find(s => s.Name == offerSettingsName);

            Excel.Worksheet offerSheet = ExcelHelper.GetSheet(_offerBook, offerSettings.SheetName);

            ShowSheetRows(offerSheet);
            _sheetProject = ExcelHelper.GetSheet(_wb, _CurrentProject.AnalysisSheetName);

            /// Столбец "номер п.п."
            OfferColumnMapping colNumber = offerSettings.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]);
            int lastRow = GetLastRow(offerSheet, colNumber.ColumnSymbol);

            int countRows = lastRow - offerSettings.RowStart + 1;
            pb.SetSubBarVolume(countRows);///100-1

            List<(int, int)> colPair = new List<(int projectCollumn, int offerColumn)>();
            int rightColumn = 10;
            foreach (OfferColumnMapping col in offerSettings.Columns)
            {
                if (string.IsNullOrEmpty(col.ColumnSymbol)) { continue; }
                ColumnMapping projectColumn = _CurrentProject.Columns.Find(a => a.Name == col.Name);
                if (!string.IsNullOrWhiteSpace(projectColumn?.ColumnSymbol ?? ""))
                {
                    int cnP = ExcelHelper.GetColumn(projectColumn.ColumnSymbol, _sheetProject);
                    int cnO = ExcelHelper.GetColumn(col.ColumnSymbol, _sheetProject);
                    colPair.Add((cnP, cnO));
                    if (rightColumn < cnO) rightColumn = cnO;
                }
            }

            Excel.Range RngData = offerSheet.Range[offerSheet.Cells[offerSettings.RowStart, 1], offerSheet.Cells[lastRow, rightColumn]];
            object[,] arrData = RngData.Value;
            for (int i = 1; i <= countRows; i++)
            {
                int rowPaste = _CurrentProject.RowStart + i - 1;
                pb.SubBarTick();
                if (pb.IsAborted) return;
                foreach ((int projectCollumn, int offerColumn) in colPair)
                {
                    object val = arrData[i, offerColumn];
                    string text = val?.ToString() ?? "";
                    Excel.Range cellPrint = _sheetProject.Cells[rowPaste, projectCollumn];
                    if (double.TryParse(text, out double number))
                    {
                        cellPrint.Value = Math.Round(number, 2);                       
                    }
                    else if (!string.IsNullOrEmpty(text))
                    {
                        cellPrint.Value = text;
                    }
                }
            }
            pb.ClearSubBar();
           
        }

        /// <summary>
        /// Показать скрытые строки на листе
        /// </summary>
        /// <param name="sh"></param>
        private void ShowSheetRows(Excel.Worksheet sh)
        {
            try
            {
                sh.Rows.Show();
                sh.UsedRange.EntireRow.Hidden = false;
            }
            catch (Exception)
            { }
        }

        /// <summary>
        ///  Найти последнюю заполненную строку в столбце
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="columnSymbol"></param>
        /// <returns></returns>
        private int GetLastRow(Excel.Worksheet sh, string columnSymbol)
        {
            Excel.Range rng = sh.Range[$"{columnSymbol}{sh.Rows.Count}"];
            int lastRow = rng.End[Excel.XlDirection.xlUp].Row;
            return lastRow;
        }

    }
}
