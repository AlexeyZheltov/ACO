using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;

namespace ACO.ProjectManager
{
    class ListAnalysis
    {
        public Excel.Worksheet SheetProjerct { get; }
        public Project CurrentProject { get; }


        public ListAnalysis()
        {

        }

        public ListAnalysis(Excel.Worksheet sheetProjerct, Project currentProject)
        {
            SheetProjerct = sheetProjerct;
            CurrentProject = currentProject;
        }

        internal void Print(Record record)
        {
            int rowPaste = CurrentProject.RowStart + record.Index;

            int colNumber = GetNumber(record);
            int ixColumn = 1;
            foreach (Field field in record.Fields)
            {
                ixColumn++;
                Excel.Range rngFirst = SheetProjerct.Cells[rowPaste, field.ColumnAnalysis.Column];
                object val = record.Values[ixColumn];
                if (field.ColumnAnalysis.Check && rngFirst.Value != val && val != null)
                {
                    SheetProjerct.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                }
            }

            /// Печать значений
            ixColumn = 1;
            foreach (Field field in record.Fields)
            {
                ixColumn++;
                object val = record.Values[ixColumn];
                if (val != null) SheetProjerct.Cells[rowPaste, field.ColumnAnalysis.Column].Value = val;
            }
        }

        private int GetNumber(Record record)
        {
            int lastRow = SheetProjerct.UsedRange.Rows.Count + SheetProjerct.UsedRange.Row - 1;
            int lastCol = CurrentProject.Columns.Max(s => s.Column);
            object[,] data = SheetProjerct.Range[SheetProjerct.Cells[CurrentProject.RowStart, 1],
                                                  SheetProjerct.Cells[lastRow, lastCol]];
            string num = record.Number;
            ColumnMapping mapping = CurrentProject.Columns.Find(n => n.Name == Project.ColumnsNames[StaticColumns.Number]);
            int columnNumber = mapping.Column;

            for (int i = 1; i <= data.GetUpperBound(0); i++)
            {
              


            }
            return 2;
        }

        //int colr()
        //{
        //    List<Col>
        //}

        //internal void Print(List<(ColumnMapping, int)> listColumnPair, int rowPaste)
        //{
        //    //   int lastRow = SheetProjerct.UsedRange.Row + SheetProjerct.UsedRange.Rows.Count + 1;
        //    foreach ((ColumnMapping projectColumn, int offerColumn) pair in listColumnPair)
        //    {
        //       // object val = arrData[i, pair.offerColumn];
        //       // Excel.Range rngFirst = SheetProjerct.Cells[rowPaste, pair.projectColumn.Column];
        //       // if (pair.projectColumn.Check && rngFirst.Value != val && val != null)
        //        {
        //            rowPaste++;
        //            SheetProjerct.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
        //            break;
        //        }
        //    }
        //    foreach ((ColumnMapping projectCollumn, int offerColumn) pair in listColumnPair)
        //    {
        //        object val = arrData[i, pair.offerColumn];
        //        if (val != null) SheetProjerct.Cells[rowPaste, pair.projectCollumn.Column].Value = val;
        //    }
        //}

        internal void PrintMarks(List<(ColumnMapping, int)> listColumnPair)
        {
            foreach ((ColumnMapping projectColumn, int offerColumn) pair in listColumnPair)
            {
                SheetProjerct.Cells[1, pair.projectColumn.Column].Value = pair.projectColumn.Name;
            }
        }
    }
}
