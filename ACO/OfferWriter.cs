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
        Excel.Worksheet _sheetProjerct = null;

        Project _project = default;
        OfferManager _offerManager = null;
        Project _CurrentProject = null;

        int _offsetPasteRange = 0;


        public OfferWriter()
        {
            //ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            //_project = new ProjectManager.ProjectManager().ActiveProject;

            //_wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            //_sheetProjerct = _wb.Worksheets[_project.AnalysisSheetName];
            //_offsetPasteRange = GetOffset(_sheetProjerct);
        }

        public OfferWriter(ExcelFile offerBook)
        {
            _app = Globals.ThisAddIn.Application;
            _wb = _app.ActiveWorkbook;
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
            OfferSettings offerSettings = _offerManager.Mappings.Find(s => s.Name == offerSettingsName);

            Excel.Worksheet offerSheet = _offerBook.GetSheet(offerSettings.SheetName);
            _sheetProjerct = GetSheet(1);

            int lastRow = GetLastRow(offerSheet);
            int countRows = lastRow - offerSettings.RowStart - 1;

            PasteHeaderRange(_sheetProjerct);

            pb.SetSubBarVolume(countRows);


            for (int i = 1; i <= countRows; i++)
            {
                int row = offerSettings.RowStart + i - 1;
                pb.SubBarTick();
                Record record = new Record();

                OfferColumnMapping colNumber = offerSettings.Columns.Find(x => x.Name == record.Number);
                foreach (OfferColumnMapping col in offerSettings.Columns)
                {
                    if (string.IsNullOrEmpty(col.ColumnSymbol)) continue;
                    ColumnMapping projectColumn = _CurrentProject.Columns.Find(a => a.Name == col.Name);
                    if (projectColumn != null)
                    {
                        object val = offerSheet.Range[$"{col.ColumnSymbol}{row}"].Value;

                        if (projectColumn.Check)
                        {
                            if (val == _sheetProjerct.Range[$"{projectColumn.ColumnSymbol}{row}"].Value)
                            {
                                //_sheetProjerct.Range[$"{projectColumn.ColumnSymbol}{row}"].Value = val;
                            }
                        }
                    }

                }

                //вывод по столбцам
                foreach (string key in record.Values.Keys)
                {
                    // ColumnMapping projectCol = _CurrentProject.Columns.Find(a => a.Value == key);
                    // _sheet.Cells[rowPaste, projectCol.Column].value = record.Values[key];
                }
            }
            pb.ClearSubBar();
        }

        private int FindNextOfferRange()
        {
            int col = _sheetProjerct.UsedRange.Column + _sheetProjerct.UsedRange.Columns.Count;
            return col;
        }

        private Excel.Range PasteHeaderRange(Worksheet sheetProjerct)
        {
            Excel.Worksheet tamplateShets = _wb.Worksheets["Шаблоны"];
            Excel.Range rangerCopy = tamplateShets.Range[tamplateShets.Cells[1, 1], tamplateShets.Range[1, 4]];

            int offsetColumn = FindNextOfferRange();

            return rangerCopy;
        }

        /// <summary>
        /// Печать КП
        /// </summary>
        /// <param name="offer"></param>
        internal void PrintSpectrum(IProgressBarWithLogUI pb)
        {
            OfferSettings offerSettings = OfferManager.GetSpectrumSettigsDefault();

            /*_offerManager.Mappings.Find(s => s.Name == "Спектрум");*/
            //if (offerSettings is null ) 

            Excel.Worksheet offerSheet = _offerBook.GetSheet(offerSettings.SheetName);
            _sheetProjerct = GetSheet(1);//_CurrentProject.AnalysisSheetName);


            /// Столбец "номер п.п."
            OfferColumnMapping colNumber = offerSettings.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]);
            int lastRow = GetLastRow(offerSheet, colNumber.ColumnSymbol);

            int countRows = lastRow - offerSettings.RowStart + 1;
            pb.SetSubBarVolume(countRows);

            //List<(string, string)> colPair = new List<(string projectCollumn, string offerColumn)>();
            List<(int, int)> colPair = new List<(int projectCollumn, int offerColumn)>();
            int rightColumn = 10;
            foreach (OfferColumnMapping col in offerSettings.Columns)
            {
                if (string.IsNullOrEmpty(col.ColumnSymbol)) { continue; }
                ColumnMapping projectColumn = _CurrentProject.Columns.Find(a => a.Name == col.Name);
                if (!string.IsNullOrWhiteSpace(projectColumn?.ColumnSymbol ?? ""))
                {
                    int cnP = GetColumn(projectColumn.ColumnSymbol, _sheetProjerct);  // _sheetProjerct.Range[$"{projectColumn.ColumnSymbol}1"].Column;
                    int cnO = GetColumn(col.ColumnSymbol, _sheetProjerct); // _sheetProjerct.Range[$"{col.ColumnSymbol}1"].Column;
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
                if (pb.IsAborted) return;// throw new AddInException("Процесс остановлен");
                //foreach (OfferColumnMapping col in offerSettings.Columns)
                //{
                // if (string.IsNullOrEmpty(col.ColumnSymbol)) continue;
                foreach ((int projectCollumn, int offerColumn) pair in colPair)
                {
                    object val = arrData[i, pair.offerColumn];
                    if (val != null) _sheetProjerct.Cells[rowPaste, pair.projectCollumn].Value = val;
                }
            }

            // pb.ClearSubBar();
        }

        //private int GetColumn(Excel.Worksheet sh,StaticColumns staticColumn)
        //{
        //    string columnSymbol = Project.ColumnsNames[staticColumn];
        //    Excel.Range rng = sh.Range[$"${columnSymbol}${sh.Rows.Count}"];
        //    return rng.Column;
        //}
        private int GetColumn(string columnSymbol, Excel.Worksheet sh)
        {
            //Excel.Worksheet sh = _app.ActiveSheet;
            int col = sh.Range[$"{columnSymbol}1"].Column;
            return col;
        }

        private int GetLastRow(Excel.Worksheet sh, string columnSymbol)
        {
            //Excel.Range rng = sh.Range[$"{columnSymbol}{sh.Rows.Count}"];
            int col = GetColumn(columnSymbol, sh);
            Excel.Range rng = sh.Cells[sh.Rows.Count, col];
            int lastRow = rng.End[Excel.XlDirection.xlUp].Row;
            return lastRow;
        }

        private Excel.Worksheet GetSheet(string analysisSheetName)
        {

            return _wb.Sheets[analysisSheetName];
        }
        private Excel.Worksheet GetSheet(int index)
        {
            return _wb.Worksheets[index];
        }

        private int GetLastRow(Excel.Worksheet sh, int col)
        {
            int lastRow = sh.Cells[sh.Rows.Count, col].End[Excel.XlDirection.xlUp].row;
            return lastRow;
        }

        private int GetLastRow(Excel.Worksheet sh)
        {
            int lastRow = sh.UsedRange.Row + sh.UsedRange.Rows.Count - 1;
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
            Excel.Range rowToInsert = _sheetProjerct.Rows[row];
            rowToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            return row;
        }

        private void InsertRow()
        {
            //TODO если такого пункта нет вставить строку
            int row = 0;
            //Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            row = GetRowByLevel();
            //Excel.Worksheet sh= 
            Excel.Range rowToInsert = _sheetProjerct.Rows[row];
            rowToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

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
            int lastRow = _sheetProjerct.Cells[_sheetProjerct.Rows.Count, columnNumbr].End[Excel.XlDirection.xlUp].row;
            Excel.Range rng = _sheetProjerct.Range[_sheetProjerct.Cells[_project.RowStart, 1], _sheetProjerct.Cells[lastRow, colEnd]];
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
