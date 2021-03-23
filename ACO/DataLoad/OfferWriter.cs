using Excel = Microsoft.Office.Interop.Excel;
using ACO.ExcelHelpers;
using ACO.ProjectManager;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;
using System;
using Microsoft.Office.Interop.Excel;

namespace ACO

{
    class OfferWriter
    {
        Excel.Application _app = default;
        /// <summary>
        ///  Текущая книга с проектом
        /// </summary>
        Excel.Workbook _wb = null;

        /// <summary>
        ///  Книга КП
        /// </summary>
        // ExcelFile
        Excel.Workbook _offerBook = null;

        /// <summary>
        /// Лист  Анализ
        /// </summary>
        Excel.Worksheet _sheetProject = null;
        OfferManager _offerManager = null;

        Project _CurrentProject = null;


        public OfferWriter(ExcelFile offerBook)
        {
            _app = Globals.ThisAddIn.Application;
            _wb = _app.ActiveWorkbook;
            _offerBook = offerBook.WorkBook;
            _offerManager = new OfferManager();
            _CurrentProject = new ProjectManager.ProjectManager().ActiveProject;
            // Лист анализ в текущем проекте
            _sheetProject = GetSheet(_wb, _CurrentProject.AnalysisSheetName);
            _CurrentProject.SetColumnNumbers(_sheetProject);
        }

        public OfferWriter(string file)
        {
            _app = Globals.ThisAddIn.Application;
            //_offerBook = offerBook;
            _wb = _app.ActiveWorkbook;
            // _offerBook = offerBook;
            _offerManager = new OfferManager();
            _CurrentProject = new ProjectManager.ProjectManager().ActiveProject;
            // Лист анализ в текущем проекте
            _sheetProject = GetSheet(_wb, _CurrentProject.AnalysisSheetName);
            _CurrentProject.SetColumnNumbers(_sheetProject);
            _offerBook = _app.Workbooks.Open(file);
            //Excel.Workbook wb = 
        }


        /// <summary>
        /// Печать КП
        /// </summary>
        /// <param name="offer"></param>
        internal void Print(IProgressBarWithLogUI pb, string offerSettingsName)
        {
            // Ищем настройки столбцов
            OfferSettings offerSettings = _offerManager.Mappings.Find(s => s.Name == offerSettingsName);
            // Лист данных КП
            Excel.Worksheet offerSheet = GetSheet(_offerBook, offerSettings.SheetName);//_offerBook.GetSheet(offerSettings.SheetName);
            ShowSheetRows(offerSheet);

            ListAnalysis SheetAnalysis = new ListAnalysis(_sheetProject, _CurrentProject);


            /// Адресация полей КП
            List<FieldAddress> addresslist = GetFields(offerSettings, SheetAnalysis.ColumnStartPrint);

            Excel.Worksheet tamplateSheet = GetSheet(_wb, "Шаблоны");
            SheetAnalysis.PrintTitle(tamplateSheet, addresslist);


            int lastRowOffer = GetLastRow(offerSheet);
            // Массив загружаемых данных
            object[,] arrData = GetArrData(offerSheet, offerSettings.RowStart, lastRowOffer);

            int countRows = lastRowOffer - offerSettings.RowStart - 1;
            pb.SetSubBarVolume(countRows);

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

                    if (field.MappingAnalysis.Name == Project.ColumnsNames[StaticColumns.Number])
                    {
                        offerRecord.Number = text;
                    }
                    if (field.MappingAnalysis.Check)
                    {
                        offerRecord.KeyFilds.Add(text);
                    }
                }
                SheetAnalysis.Print(offerRecord);
            }
            if (pb.IsAborted) throw new AddInException("Процесс остановлен.");
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
            string letterNameOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.Name]); //offerSettings.GetColumn(StaticColumns.Name).ColumnSymbol;
            string letterAmountSpectrum = _CurrentProject.GetColumn(StaticColumns.Amount).ColumnSymbol;
            string letterAmountOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.Amount]);

            //Наименование вида работ
            cell.Formula = $"=${letterNameSpectrum}{rowStart}=${letterNameOffer}{rowStart}";
            //Комментарии Спектрум к описанию работ
            _sheetProject.Cells[rowStart, colStart + 1].Formula = $"=IF(${letterNameSpectrum}{rowStart}=TRUE,\".\",Комментарии!$A$2)";
            //Отклонение по объемам
            _sheetProject.Cells[rowStart, colStart + 2].Formula = $"={letterAmountSpectrum}{rowStart}/{letterAmountOffer}{rowStart}-1";
            //Комментарии Спектрум к объемам работ
            cell = _sheetProject.Cells[rowStart, colStart + 2];
            string address = cell.Address;
            cell = _sheetProject.Cells[rowStart, colStart + 3];
            string letter = address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
            cell.Formula = $"=IF(${letter}{rowStart}>15%,Комментарии!$A$5,IF(${letter}{rowStart}<-15%,Комментарии!$A$6))";

            //Отклонение по стоимости
            string letterTotalOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.CostTotal]); //offerSettings.GetColumn(StaticColumns.CostTotal).ColumnSymbol;
            string letterTotalSpectrum = _CurrentProject.GetColumn(StaticColumns.CostTotal).ColumnSymbol;
            _sheetProject.Cells[rowStart, colStart + 4].Formula =
                        $"=IF(${letterTotalSpectrum}{rowStart}<>0," +
                        $"${letterTotalOffer}{rowStart}/${letterTotalSpectrum}{rowStart}-1,0)";

            //Комментарии Спектрум к стоимости работ
            cell = _sheetProject.Cells[rowStart, colStart + 4];
            address = cell.Address;
            string letterDiffCost = address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
            _sheetProject.Cells[rowStart, colStart + 5].Formula =
                $"=IF(${letterDiffCost}{rowStart}>15%,Комментарии!$A$9,IF(${letterDiffCost}{rowStart}<-15%,Комментарии!$A$10,\".\"))";
            //Отклонение по стоимости МАТ
            string letterWorkslOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.CostWorksTotal]); //offerSettings.GetColumn(StaticColumns.CostWorksTotal).ColumnSymbol;
            string letterWorksSpectrum = _CurrentProject.GetColumn(StaticColumns.CostWorksTotal).ColumnSymbol;
            _sheetProject.Cells[rowStart, colStart + 6].Formula =
                        $"=IF(${letterWorkslOffer}{rowStart}<>0;" +
                        $"${letterWorkslOffer}{rowStart}/${letterWorksSpectrum}{rowStart});\"Отс-ет ст-ть мат.\"";
            //Отклонение по стоимости РАБ
            string letterMaterialslOffer = GetLetter(columnsOffer, Project.ColumnsNames[StaticColumns.CostMaterialsTotal]);  //offerSettings.GetColumn(StaticColumns.CostMaterialsTotal).ColumnSymbol;
            string letterMaterialsSpectrum = _CurrentProject.GetColumn(StaticColumns.CostMaterialsTotal).ColumnSymbol;
            _sheetProject.Cells[rowStart, colStart + 7].Formula =
                        $"=IF(${letterMaterialslOffer}{rowStart}<>0;" +
                        $"${letterMaterialslOffer}{rowStart}/${letterMaterialsSpectrum}{rowStart});\"Отс-ет ст-ть мат.\"";

            //Комментарии к строкам "0"
            _sheetProject.Cells[rowStart, colStart + 8].Formula =
                        $"=IF(${letterDiffCost}{rowStart}=100,\"Указать стоимость единичной расценки и посчитать итог\",\".\")";


            // Протянуть до конца листов
            Excel.Range rng = _sheetProject.Range[_sheetProject.Cells[rowStart, colStart],
                                                _sheetProject.Cells[rowStart, colStart + 8]];
            int lastRow = GetLastRow(_sheetProject, letterNameOffer);
            //int lastRow = _sheetProject.UsedRange.Row + _sheetProject.UsedRange.Rows.Count - 1;
            if (lastRow > rowStart)
            {
                Excel.Range destination = _sheetProject.Range[_sheetProject.Cells[rowStart, colStart], _sheetProject.Cells[lastRow, colStart + 8]];
                rng.AutoFill(destination);
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
                    return col + 1;
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
                    сolumnProject.Column = GetColumn(сolumnProject.ColumnSymbol, _sheetProject);
                    int colPaste = lastCol + k;
                    int colOffer = GetColumn(columnOffer.ColumnSymbol, _sheetProject);
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
        internal void PrintSpectrum(IProgressBarWithLogUI pb)
        {
            OfferSettings offerSettings = OfferManager.GetSpectrumSettigsDefault();
            Excel.Worksheet offerSheet = GetSheet(_offerBook, offerSettings.SheetName);

            ShowSheetRows(offerSheet);
            _sheetProject = GetSheet(_wb, _CurrentProject.AnalysisSheetName);

            /// Столбец "номер п.п."
            OfferColumnMapping colNumber = offerSettings.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]);
            int lastRow = GetLastRow(offerSheet, colNumber.ColumnSymbol);

            int countRows = lastRow - offerSettings.RowStart + 1;
            pb.SetSubBarVolume(countRows);

            List<(int, int)> colPair = new List<(int projectCollumn, int offerColumn)>();
            int rightColumn = 10;
            foreach (OfferColumnMapping col in offerSettings.Columns)
            {
                if (string.IsNullOrEmpty(col.ColumnSymbol)) { continue; }
                ColumnMapping projectColumn = _CurrentProject.Columns.Find(a => a.Name == col.Name);
                if (!string.IsNullOrWhiteSpace(projectColumn?.ColumnSymbol ?? ""))
                {
                    int cnP = GetColumn(projectColumn.ColumnSymbol, _sheetProject);  // _sheetProjerct.Range[$"{projectColumn.ColumnSymbol}1"].Column;
                    int cnO = GetColumn(col.ColumnSymbol, _sheetProject);//_sheetProjerct.Range[$"{col.ColumnSymbol}1"].Column;
                    colPair.Add((cnP, cnO));
                    if (rightColumn < cnO) rightColumn = cnO;
                }
            }
            //int lastCol = offerSheet.UsedRange.Column + offerSheet.UsedRange.Columns.Count - 1;
            Excel.Range RngData = offerSheet.Range[offerSheet.Cells[offerSettings.RowStart, 1], offerSheet.Cells[lastRow, rightColumn]];
            object[,] arrData = RngData.Value;

            for (int i = 1; i <= countRows; i++)
            {
                // int rowRead = offerSettings.RowStart + i - 1;
                int rowPaste = _CurrentProject.RowStart + i - 1;
                pb.SubBarTick();
                if (pb.IsAborted) return;
                // throw new AddInException("Процесс остановлен");
                //foreach (OfferColumnMapping col in offerSettings.Columns)
                //{
                // if (string.IsNullOrEmpty(col.ColumnSymbol)) continue;
                foreach ((int projectCollumn, int offerColumn) in colPair)
                {
                    object val = arrData[i, offerColumn];
                    if (val != null) _sheetProject.Cells[rowPaste, projectCollumn].Value = val;
                }
            }
            pb.ClearSubBar();
            _offerBook.Close(false);
        }

        /// <summary>
        /// Показать скрытые строки на листе
        /// </summary>
        /// <param name="sh"></param>
        private void ShowSheetRows(Excel.Worksheet sh)
        {
            //offerSheet.Rows.Show();
            //sh.Outline.ShowLevels();
            try
            {
                sh.Rows.Show();
                sh.UsedRange.EntireRow.Hidden = false;
            }
            catch (Exception)
            { }
        }

        /// <summary>
        ///  Номер стодбца по его буквенному обозначению
        /// </summary>
        /// <param name="columnSymbol"></param>
        /// <param name="sh"></param>
        /// <returns></returns>
        private int GetColumn(string columnSymbol, Excel.Worksheet sh)
        {
            int col = sh.Range[$"{columnSymbol}1"].Column;
            return col;
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

        /// <summary>
        /// Получить лист по номеру
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        private Excel.Worksheet GetSheet(int index)
        {
            return _wb.Worksheets[index];
        }
        /// <summary>
        ///  Получить лист по имени
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private Excel.Worksheet GetSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
            {
                if (sh.Name == name)
                {
                    return sh;
                }
            }
            throw new AddInException($"Лист {name} отсутствует");
        }

        private int GetLastRow(Excel.Worksheet sh)
        {
            int lastRow = sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1;
            return lastRow;
        }
    }
}
