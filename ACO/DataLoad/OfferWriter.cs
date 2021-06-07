using Excel = Microsoft.Office.Interop.Excel;
using ACO.ExcelHelpers;
using ACO.ProjectManager;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;
using System;
using System.Drawing;
using System.Diagnostics;

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

        /// <summary>
        /// Собрать пары маппингов столбцов
        /// </summary>
        /// <param name="offerSettings"></param>
        /// <param name="lastCol"></param>
        /// <returns></returns>
        private List<FieldAddress> GetFields(OfferSettings offerSettings, int lastCol)
        {
            List<FieldAddress> fields = new List<FieldAddress>();
            int k = 0;
            foreach (OfferColumnMapping columnOffer in offerSettings.Columns)
            {
                if (string.IsNullOrEmpty(columnOffer.ColumnSymbol)) continue;
                ColumnMapping сolumnProject = _CurrentProject.Columns.Find(a => a.Name == columnOffer.Name);
                if (сolumnProject == null) continue;
                if (сolumnProject.Obligatory)
                {
                    try
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
                    catch (Exception ex) 
                    { Debug.WriteLine(ex.Message); }
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
        public static int GetLastRow(Excel.Worksheet sh, string columnSymbol)
        {
            Excel.Range rng = sh.Range[$"{columnSymbol}{sh.Rows.Count}"];
            int lastRow = rng.End[Excel.XlDirection.xlUp].Row;
            return lastRow;
        }

    }
}
