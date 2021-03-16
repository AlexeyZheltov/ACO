using ACO.ExcelHelpers;
using ACO.ProjectManager;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.Offers
{
    class OfferWriter
    {
        Excel.Application _app = default;
        Excel.Workbook _wb = null;
        ExcelFile _offerBook = null;
        Excel.Worksheet _sheet = null;

        Project _project = default;
        OfferManager _offerManager = null;
        Project _CurrentProject = null;

        int _offsetPasteRange = 0;


        public OfferWriter()
        {
            //ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            _project = new ProjectManager.ProjectManager().ActiveProject;

            _wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            _sheet = _wb.Worksheets[_project.AnalysisSheetName];
            _offsetPasteRange = GetOffset(_sheet);
        }

        public OfferWriter(ExcelFile offerBook)
        {
            _app = Globals.ThisAddIn.Application;
            //Excel.Workbook _wb = _app.ActiveWorkbook;
            _offerBook = offerBook;
            _offerManager = new OfferManager();
            _CurrentProject = new ProjectManager.ProjectManager().ActiveProject;
        }

        /// <summary>
        ///  
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private int GetOffset(Excel.Worksheet sheet)
        {
            //sheet.

            return 0;
        }

        /// <summary>
        /// Печать КП
        /// </summary>
        /// <param name="offer"></param>
        internal void Print(IProgressBarWithLogUI pb, string offerSettingsName)
        {            
            OfferSettings offerSettings = _offerManager.Mappings.Find(s=>s.Name == offerSettingsName);

            Excel.Worksheet offerSheet = _offerBook.GetSheet(offerSettings.SheetName);
            _sheet = _offerBook.GetSheet(1);
            
            int lastRow = GetLastRow(offerSheet);
            int countRows = lastRow - offerSettings.RowStart - 1;

            pb.SetSubBarVolume(countRows);

            // for (int row = offerSettings.RowStart; row <= lastRow; row++)
            //string columnName = _CurrentProject.Columns.Find(a => a.Name == Project.ColumnsNames[StaticColumns.Name]).ColumnSymbol;

            for (int i = 1; i<=countRows; i++)
            {
                int row = offerSettings.RowStart + i - 1;
                pb.SubBarTick();
                Record record = new Record();
                                              
                OfferColumnMapping colNumber = offerSettings.Columns.Find(x => x.Name == record.Number );
                foreach (OfferColumnMapping col in offerSettings.Columns)
                {
                    if (string.IsNullOrEmpty(col.ColumnSymbol)) continue;                    
                    ColumnMapping projectColumn = _CurrentProject.Columns.Find(a => a.Name == col.Name);
                    if (projectColumn != null)
                    {
                        object val = offerSheet.Range[$"${col.ColumnSymbol}${row}"].Value;

                        if (projectColumn.Check)
                        {
                           
                        }
                        if (projectColumn.Obligatory)
                        {

                        }
                    }

                    //if (!string.IsNullOrEmpty(col.Name))
                    //{
                        //if (projectCol is null) continue;
                        ////string val = offerSheet.Cells[row, col.Column].value?.ToString() ?? "";
                        ////if (string.IsNullOrWhiteSpace(val))
                        //object val = offerSheet.Range[$"${col.ColumnSymbol}${row}"].Value;
                        //if(val != null)
                        //{
                        //    if (projectCol.Check)
                        //    {
                        //        string v = val.ToString();
                        //        record.KeyFilds.Add(v);
                        //    record.Values.Add(col.Name, val);
                        //    }
                        //}
                    //}
                }
               int rowPaste = GetRow(record);
                
                //вывод по столбцам
                foreach (string key in record.Values.Keys)
                {
                   // ColumnMapping projectCol = _CurrentProject.Columns.Find(a => a.Value == key);
                   // _sheet.Cells[rowPaste, projectCol.Column].value = record.Values[key];
                }

                /// Вставка диазазона сумм
                //Excel.Range copyRng = offerSheet.Range[
                //    offerSheet.Cells[row, offerSettings.RangeValuesStart],
                //    offerSheet.Cells[row, offerSettings.RangeValuesEnd]];


                //Excel.Range pasteRng = _sheet.Range[
                //    _sheet.Cells[rowPaste, _CurrentProject. RangeValuesStart],
                //    _sheet.Cells[rowPaste, _CurrentProject.RangeValuesEnd]];
                //if (pasteRng.Cells.Count == copyRng.Cells.Count)
                //{
                //    pasteRng.Value = copyRng.Value;
                //}
                //else { 
                //    throw new AddInException("Укажите одинаковое количество столбцов " +
                //                            "в диапазонах значений КП и проекта"); }

                //_offerManager.
            }
            pb.ClearSubBar();
        }

        private Worksheet GetSheet(string analysisSheetName)
        {
           return _wb.Worksheets[analysisSheetName];
        }

        private int GetLastRow(Worksheet offerSheet)
        {
            int lastRow = offerSheet.UsedRange.Row + offerSheet.UsedRange.Rows.Count - 1;
            return lastRow;
        }

        //internal void PrintOffer(Offer offer)
        //{
        //    //TODO Определить место вставки  
        //    Excel.Range rng = CopyRange();

        //    List<ColumnMapping> columnsMapping = _project.Columns;

        //    foreach (Record record in offer.Records)
        //    {
        //        int rowPrint = GetRow(record.KeyFilds);
        //        if (rowPrint == 0) throw new AddInException("Не удалось определить строку вставки. Номер перечня: " + record.Number);
        //        foreach (ColumnMapping col in columnsMapping)
        //        {
        //            int columnPrint = 0;
        //            //TODO определить столбец вставки 
        //            if (record.Values.ContainsKey(col.Value))
        //            {
        //                object val = record.Values[col.Value];
        //                Excel.Range cellPrint = rng.Cells[rowPrint, columnPrint];
        //                cellPrint.Value = val;
        //            }
        //        }
        //    }
        //}

        //private Excel.Range CopyRange()
        //{
        //    int firstCol = _project.FirstColumnOffer;
        //    int lastCol = _project.LastColumnOffer;
        //    int lastRow = _project.RowStart;

        //    Excel.Range rng = _sheet.Range[_sheet.Cells[1, firstCol], _sheet.Cells[lastRow, lastCol]];
        //    int col = _project.FirstColumnOffer;
        //    int colCount = _project.LastColumnOffer - _project.FirstColumnOffer + 1;

        //    while (_sheet.Cells[1, col].Value != "")
        //    {
        //        col = col + colCount;
        //    }
        //    rng.Copy(_sheet.Cells[1, col]);
        //    rng = _sheet.Range[_sheet.Cells[1, col], _sheet.Cells[lastRow, col + colCount - 1]];
        //    return rng;
        //}



        //TODO определить строку вставки 
        private int GetRow(Record record)
        {
            int row = 0;

            if (row == 0) row = InsertRow(record);
            return row;
        }

        private int InsertRow(Record recordist)
        {
            //TODO если такого пункта нет вставить строку
            int row = 0;
            //Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            row = GetRowByLevel();
            //Excel.Worksheet sh= 
            Excel.Range rowToInsert = _sheet.Rows[row];
            rowToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            return row;
        }

        private int GetRowByLevel()
        {
            return 3;
        }

        Dictionary<string, int> _rowKeys = default;
        private ExcelFile excelBook;

        private void UpdateRowKeys()
        {
            _rowKeys = new Dictionary<string, int>();
            int columnNumbr = 10;
            int colEnd = 13;
            int lastRow = _sheet.Cells[_sheet.Rows.Count, columnNumbr].End[Excel.XlDirection.xlUp].row;
            Excel.Range rng = _sheet.Range[_sheet.Cells[_project.RowStart, 1], _sheet.Cells[lastRow, colEnd]];
            object[,] data = rng.Value;

            for (int i = 1; i < data.GetUpperBound(0); i++)
            {
                string key = "";
                foreach (ColumnMapping col in _project.Columns)
                {
                    if (col.Check)
                    {   //columnNumbr
                       // string val = data[i, col.Column]?.ToString() ?? "";
                       // key += val;
                    }
                    int row = i + _project.RowStart - 1;
                    _rowKeys.Add(key, row);
                }
            }
        }

        private int GetRowBy(Record record)
        {
            // string key = "";
            //foreach (ColumnMapping col in record.Key)
            // return _rowKeys[record.Key];

            //int columnNumbr = 10;//Project.staticColumns[StaticColumns.Number]
            //int colEnd = 13;
            ////_sheet
            //int lastRow = _sheet.Cells[_sheet.Rows.Count, columnNumbr].End[Excel.XlDirection.xlUp].row;
            //Excel.Range rng = _sheet.Range[_sheet.Cells[_project.RowStart, 1], _sheet.Cells[7, colEnd]];
            //object[,] data = rng.Value;
            ////data[1,2]
            //for (int i = 1; i < data.GetUpperBound(0); i++)
            //{
            //    string key = "";
            //    foreach( ColumnMapping col in _project.Columns )
            //    {
            //        if (col.Check)
            //        {   //columnNumbr
            //            string val = data[i, col.Column ]?.ToString() ?? "";
            //            key += val;
            //        }
            //    }
            //}
            return 1;
        }

    }
}
