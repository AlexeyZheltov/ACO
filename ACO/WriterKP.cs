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
        Project _project = default;
        Excel.Worksheet _sheet = null;
        Excel.Workbook _wb = null;
        int _offsetPasteRange = 0;
        public OfferWriter()
        {
            //ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            _project = new ProjectManager.ProjectManager().ActiveProject;
            _wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            _sheet = _wb.Worksheets[_project.AnalysisSheetName];
            _offsetPasteRange = GetOffset(_sheet);
        }

        /// <summary>
        ///  
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private int GetOffset(Worksheet sheet)
        {
            //sheet.

            return 0;
        }

        /// <summary>
        /// Печать КП
        /// </summary>
        /// <param name="offer"></param>
        internal void PrintOffer(Offer offer)
        {
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            //
            //TODO Определить место вставки  
            List<ColumnMapping> columnsMapping = _project.Columns;

            foreach (Record record in offer.Records)
            {
                int rowPrint = GetRow(record.Key);
                if (rowPrint == 0) throw new AddInException("Не удалось определить строку вставки. Номер перечня: " + record.Number);
                foreach (ColumnMapping col in columnsMapping)
                {
                    int columnPrint = 0; //TODO определить столбец вставки 
                    if (record.Values.ContainsKey(col.Value))
                    {
                        object val = record.Values[col.Value];
                        Excel.Range cellPrint = sh.Cells[rowPrint, columnPrint];
                        cellPrint.Value = val;
                    }
                }
            }
        }

        private int GetRow(string number)
        {
            //TODO определить строку вставки 
            int row = 0;
            if (row == 0) row = InsertRow(number);

            return row;
        }

        private int InsertRow(string number)
        {
            //TODO если такого пункта нет вставить строку
            int row = 0;
            Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            row = GetRowByLevel();
            //Excel.Worksheet sh= 
            Excel.Range rowToInsert = sh.Rows[row];
            rowToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            return row;
        }

        private int GetRowByLevel()
        {
            return 3;
        }

        Dictionary<string, int> _rowKeys = default;
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
                        string val = data[i, col.Column]?.ToString() ?? "";
                        key += val;
                    }
                    int row = i + _project.RowStart-1 ;
                    _rowKeys.Add(key, row);
                }
            }
        }
        
        private int GetRowBy(Record record)
        {
            // string key = "";
            //foreach (ColumnMapping col in record.Key)
            return _rowKeys[record.Key];

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
        }

    }
}
