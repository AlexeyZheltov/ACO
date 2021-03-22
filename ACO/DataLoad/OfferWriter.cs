using Excel = Microsoft.Office.Interop.Excel;
using ACO.ExcelHelpers;
using ACO.ProjectManager;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;
using System;

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
        ExcelFile _offerBook = null;

        /// <summary>
        /// Лист  Анализ
        /// </summary>
        Excel.Worksheet _sheetProject = null;

        Project _project = default;
        OfferManager _offerManager = null;
        Project _CurrentProject = null;
        int _offsetPasteRange = 0;

        public OfferWriter(ExcelFile offerBook)
        {
            _app = Globals.ThisAddIn.Application;
            _wb = _app.ActiveWorkbook;
            _offerBook = offerBook;
            _offerManager = new OfferManager();
            _CurrentProject = new ProjectManager.ProjectManager().ActiveProject;
            // Лист анализ в текущем проекте
            _sheetProject = GetSheet(_CurrentProject.AnalysisSheetName);
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
            // Лист данных КП
            Excel.Worksheet offerSheet = _offerBook.GetSheet(offerSettings.SheetName);
            ShowSheetRows(offerSheet);

            ListAnalysis SheetAnalysis = new ListAnalysis(_sheetProject, _CurrentProject);


            /// Адресация полей КП
            List<FieldAddress> addresslist = GetFields(offerSettings, SheetAnalysis.ColumnStartPrint);

            Excel.Worksheet tamplateSheet = GetSheet("Шаблоны");
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

                Record offerRecord = new Record();
                offerRecord.Addresslist = addresslist;
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
            Excel.Worksheet offerSheet = _offerBook.GetSheet(offerSettings.SheetName);
            
            ShowSheetRows(offerSheet);
            _sheetProject = GetSheet(_CurrentProject.AnalysisSheetName);

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
                foreach ((int projectCollumn, int offerColumn) pair in colPair)
                {
                    object val = arrData[i, pair.offerColumn];
                    if (val != null) _sheetProject.Cells[rowPaste, pair.projectCollumn].Value = val;
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
        private Excel.Worksheet GetSheet(string name)
        {
            foreach (Excel.Worksheet sh in _wb.Worksheets)
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
