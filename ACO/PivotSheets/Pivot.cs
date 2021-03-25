using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.PivotSheets
{
    /// <summary>
    ///  Загрузка и обновление сводных таблиц
    /// </summary>
    class Pivot
    {
        Excel.Application _app = Globals.ThisAddIn.Application;
        Excel.Worksheet _SheetUrv11 ;
        Excel.Worksheet _AnalisysSheet;
        ProjectManager.ProjectManager _projectManager;
        ProjectManager.Project _project;
        public Pivot()
        {
           // _app = Globals.ThisAddIn.Application;
            string sheetName = "Урв 11";
            Excel.Workbook wb = _app.ActiveWorkbook;
            _SheetUrv11 = OfferWriter.GetSheet(wb, sheetName);
            _projectManager = new ProjectManager.ProjectManager();
            _project = _projectManager.ActiveProject;
            string analisysSheetName = _project.AnalysisSheetName;
            _AnalisysSheet = OfferWriter.GetSheet(wb, analisysSheetName);
        }

       public void LoadUrv12(IProgressBarWithLogUI pb)
        {
            /// Добавить список
            /// Добавить столбцы КП
            /// Проставить формулы
            ///
            string letterName = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Name]).ColumnSymbol;
            string letterNumber = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            string letterLevel = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            string letterCost = _project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol;

            int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
            //foreach (Excel.Range row in _AnalisysSheet.UsedRange.Rows)
            int rowPaste = 13;
            for (int row = _project.RowStart; row<=lastRow; row++)
            {
               string name=  _AnalisysSheet.Range[$"${letterName}{row}"].Value.ToString()??"";
               string number=  _AnalisysSheet.Range[$"${letterNumber}{row}"].Value.ToString()??"";
               string level = _AnalisysSheet.Range[$"${letterLevel}{row}"].Value.ToString() ?? "";
                string cost = _AnalisysSheet.Range[$"${letterCost}{row}"].Value.ToString() ?? "";
                int levelNum = int.TryParse(level, out int ln) ? ln : 0;
                
                if (levelNum > 0 && levelNum < 6)
                {
                    _SheetUrv11.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    _SheetUrv11.Cells[rowPaste, 2].Value = number;
                    _SheetUrv11.Cells[rowPaste, 3].Value = name;
                    PrintOffers(rowPaste,row);
                    rowPaste++;
                }
            }

        }

        private void PrintOffers(int rowPaste, int row)
        {
            //_SheetUrv11
        }
        class OfferComments
        {
            public string ParticipantName { get; set; }
            public string PercentMaterial { get; set; }
            public string PercentWorks { get; set; }
            public string PercentTotal { get; set; }
            public double TotalCost { get; set; }
        }
    }
}
