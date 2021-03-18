using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;

namespace ACO.ProjectManager
{
    class ListAnalysis
    {
        int _ColumnStartPrint = 0;

        public List<Record> Records
        {
            get
            {
                _Records = GetListRecordsAnalysis();
                return _Records;
            }
            set
            {
                _Records = value;
            }
        }
        List<Record> _Records;

        public Excel.Worksheet SheetProjerct { get; }
        public Project CurrentProject { get; }


        public ListAnalysis()
        {

        }

        public ListAnalysis(Excel.Worksheet sheetProjerct, Project currentProject)
        {
            SheetProjerct = sheetProjerct;
            CurrentProject = currentProject;
            _ColumnStartPrint = sheetProjerct.UsedRange.Column + sheetProjerct.UsedRange.Columns.Count;
            _rowStart = CurrentProject.RowStart;
        }

        int _rowStart =1;

        internal void Print(Record recordPrint)
        {
            int ixColumn = 1;          
            int rowPaste = _rowStart;
            int lastRow = SheetProjerct.UsedRange.Row + SheetProjerct.UsedRange.Rows.Count + 1;
            //recordPrint.Number
            bool curentLevel = false;
            for (int row = _rowStart; row <= lastRow; row++)
            {
                Record recordAnalysis = GetRecocdAnalysis(row);
                //  recordAnalysis.Number = SheetProjerct.Cells[CurrentProject.Columns.]

                if (recordAnalysis.LevelEqual(recordPrint))
                {
                    curentLevel = true;
                    _rowStart = row;

                    if (recordAnalysis.Equal(recordPrint))
                    {
                        rowPaste = row;
                        break;
                    }
                }
                else if (curentLevel)
                {
                    SheetProjerct.Rows[row-1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    rowPaste = row;                    
                    break;
                }
            }

            /// Печать значений
            ixColumn = 1;
            foreach (FieldAddress field in recordPrint.Addresslist)
            {
                object val = recordPrint.Values[ixColumn];
                if (val != null) SheetProjerct.Cells[rowPaste, field.MappingAnalysis.Column].Value = val;
                ixColumn++;
            }
        }

        private Record GetRecocdAnalysis(int row)
        {
            Record recordAnalysis = new Record();
            ColumnMapping mappingNumber = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Name]);
            object number = SheetProjerct.Cells[$"{mappingNumber.ColumnSymbol}{row}"].value;
            recordAnalysis.Number = number.ToString();

            foreach (ColumnMapping columnMapping in CurrentProject.Columns)
            {
                if (columnMapping.Check)
                {
                    object val = SheetProjerct.Cells[$"{columnMapping.ColumnSymbol}{row}"].value;
                    string key = val?.ToString();

                    recordAnalysis.KeyFilds.Add(key);
                }
            }
            return recordAnalysis;
        }

        internal void Print1(IProgressBarWithLogUI pb, string offerSettingsName)
        {
            //offerSheet.UsedRange.Rows.s  //Outline. ShowLevels();// Rows. EntireRow.
            // int lastRowOffer = GetLastRow(offerSheet);

            // PasteHeaderRange(_sheetProjerct);
            //List<(int, int)> listColumnPair = GetHeaders(offerSettings);


            /// Столбец проект \ столбец КП.
            //List<(ColumnMapping, int)> listColumnPair = GetColumnHeaders(offerSettings);

            // SheetAnalysis.PrintMarks(listColumnPair);

            //int rowPaste = _CurrentProject.RowStart - 1;
            //for (int i = 1; i <= countRows; i++)
            //{
            //    pb.SubBarTick();
            //    if (pb.IsAborted) return; //throw new AddInException("К");
            //    int row = offerSettings.RowStart + i - 1;
            //    Record record = new Record();
            //    record.Fields = fields;
            //    record.Index = i;
            //    for(int k = 1; k <= fields.Count; k++)
            //    {
            //        record.Values.Add(k, arrData[i, k]);
            //         //object val = arrData[i, 1];
            //    }
            //    SheetAnalysis.Print(record);
            //rowPaste += i;
            //SheetAnalysis.Print(listColumnPair, rowPaste);

            // Пропустить строки                
            //int lastRow = _sheetProjerct.UsedRange.Row + _sheetProjerct.UsedRange.Rows.Count + 1;                
            //foreach ((ColumnMapping projectColumn, int offerColumn) pair in listColumnPair)
            //{
            //    object val = arrData[i, pair.offerColumn];
            //    Excel.Range rngFirst = _sheetProjerct.Cells[rowPaste, pair.projectColumn.Column];
            //    if (pair.projectColumn.Check && rngFirst.Value != val && val !=null)
            //    { 
            //        rowPaste++;
            //        _sheetProjerct.Rows[rowPaste].Insert(Excel.XlInsertShiftDirection.xlShiftDown);                        
            //        break;
            //    }
            //}
            //foreach ((ColumnMapping projectCollumn, int offerColumn) pair in listColumnPair)
            //{
            //    object val = arrData[i, pair.offerColumn];
            //    if (val != null) _sheetProjerct.Cells[rowPaste, pair.projectCollumn.Column].Value = val;
            //}
            //{
            //    if (string.IsNullOrEmpty(col.ColumnSymbol)) continue;
            //    ColumnMapping projectColumn = _CurrentProject.Columns.Find(a => a.Name == col.Name);
            //    if (projectColumn != null)
            //    {
            //        object val = offerSheet.Range[$"{col.ColumnSymbol}{row}"].Value;

            //        if (projectColumn.Check)
            //        {
            //            if (val == _sheetProjerct.Range[$"{projectColumn.ColumnSymbol}{row}"].Value)
            //            {
            //                //_sheetProjerct.Range[$"{projectColumn.ColumnSymbol}{row}"].Value = val;
            //            }
            //        }
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

            int rowsCount = data.GetUpperBound(0);
            for (int i = 1; i <= rowsCount; i++)
            {
                string cellNumber = data[i, columnNumber]?.ToString() ?? "";
                cellNumber = cellNumber.Trim(new Char[] { ' ', '.' });
                if (cellNumber == num)
                {

                }
            }
            return 2;
        }


        /// <summary>
        ///  Копирование заголовков
        /// </summary>
        /// <param name="addresslist"></param>
        internal void PrintTitle(Excel.Worksheet tamplateSheet, List<FieldAddress> addresslist)
        {
            //FieldAddress firstColMapping = addresslist.Min(x => x.MappingAnalysis.Column).MappingAnalysis.Column;
            //int firstCol = addresslist.First().MappingAnalysis.Column;
            int lastCol = addresslist.Last().MappingAnalysis.Column;

            foreach (FieldAddress address in addresslist)
            {
                int col = address.MappingAnalysis.Column;
                if (col > lastCol) lastCol = col;
                //if (col < firstCol) firstCol = col;
            }
            Excel.Range titleTamplate = SheetProjerct.Range[SheetProjerct.Cells[7, 1], SheetProjerct.Cells[8, lastCol]];
            //Excel.Range title = null;
            int columnPaste = _ColumnStartPrint;
            foreach (FieldAddress address in addresslist)
            {
                Excel.Range rngCoulumn = titleTamplate.Columns[address.MappingAnalysis.Column];
                rngCoulumn.Copy(SheetProjerct.Cells[7, columnPaste]);
                columnPaste++;
                //if (title is null)
                //{
                //    title = rngCoulumn;
                //}
                //else
                //{
                //    title = Globals.ThisAddIn.Application.Union(title, rngCoulumn);
                //}
            }
            Excel.Range commentsTitleRng = tamplateSheet.Range["A10:I14"];
            commentsTitleRng.Copy(SheetProjerct.Cells[5, columnPaste]);
            //if (title != null)
            //{
            //    title.Copy(SheetProjerct.Cells[7, _ColumnStartPrint]);
            //}
        }




        private void FindRecord()
        {

        }

        private List<Record> GetListRecordsAnalysis()
        {
            List<Record> records = new List<Record>();
            int lastRow = SheetProjerct.UsedRange.Rows.Count + SheetProjerct.UsedRange.Row - 1;
            int lastCol = CurrentProject.Columns.Max(s => s.Column);
            object[,] data = SheetProjerct.Range[SheetProjerct.Cells[CurrentProject.RowStart, 1],
                                                   SheetProjerct.Cells[lastRow, lastCol]];
            int ixCol = 1;
            int rowsCount = data.GetUpperBound(0);

            for (int row = 1; row < rowsCount; row++)
            {
                Record record = new Record();
                //record.Number = ;
                //string val = data[row, ]?.ToString() ?? "";

                foreach (ColumnMapping mapping in CurrentProject.Columns)
                {
                    if (mapping.Name == Project.ColumnsNames[StaticColumns.Number])
                    {
                        record.Number = data[row, mapping.Column]?.ToString() ?? "";
                    }
                    if (mapping.Check)
                    {
                        string val = data[row, ixCol]?.ToString() ?? "";
                        record.KeyFilds.Add(val);
                        ixCol++;
                    }

                }
                records.Add(record);
            }
            return records;
        }

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
