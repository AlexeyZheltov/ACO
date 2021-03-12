using ACO.ProjectManager;
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
       public OfferWriter()
        {
            //ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            _project = new ProjectManager.ProjectManager().ActiveProject;
          
            _wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            _sheet = _wb.Worksheets[_project.AnalysisSheetName];
        }


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

    

    }
}
