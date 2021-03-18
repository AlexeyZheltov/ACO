using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;
using ACO.ProjectManager;

namespace ACO
{
    /// <summary>
    ///  Загружает данные на лист Анализ
    /// </summary>
    class ListAnalysis
    {         

        public Excel.Worksheet SheetAnalysis { get; }
        public Project CurrentProject { get; }

        public int ColumnStartPrint
        {
            get
            {
                if (_ColumnStartPrint == 0)
                {
                    _ColumnStartPrint = SheetAnalysis.UsedRange.Column + SheetAnalysis.UsedRange.Columns.Count +1 ;
                }
                return _ColumnStartPrint;
            }
            set
            {
                _ColumnStartPrint = value;
            }
        }
        int _ColumnStartPrint = 0;

        public ListAnalysis()
        {

        }

        public ListAnalysis(Excel.Worksheet sheetProjerct, Project currentProject)
        {
            SheetAnalysis = sheetProjerct;
            CurrentProject = currentProject;
            _rowStart = CurrentProject.RowStart;
        }

        int _rowStart = 1;

        internal void Print(Record recordPrint)
        {           
            int rowPaste = _rowStart;

            /// Последняя строка списка 
            int lastRow = SheetAnalysis.UsedRange.Row + SheetAnalysis.UsedRange.Rows.Count - 3;
            //recordPrint.Number
            bool curentLevel = false;
            for (int row = _rowStart; row <= lastRow; row++)
            {

                Record recordAnalysis = GetRecocdAnalysis(row);
                if (string.IsNullOrEmpty(recordAnalysis.Number)) continue;

                /// Проверка уровня: совпадение номера предпоследнего номера
                if (recordAnalysis.LevelEqual(recordPrint))
                {
                    curentLevel = true;
                    _rowStart = row;
                    // Проверка ключевых значений 
                    if (recordAnalysis.KeyEqual(recordPrint))
                    {
                        rowPaste = row;
                        break;
                    }
                }

                else if (curentLevel || row==lastRow)
                {
                    SheetAnalysis.Rows[row].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    rowPaste = row;
                    break;
                }
            }

            /// Печать значений
            foreach (FieldAddress field in recordPrint.Addresslist)
            {
                object val = recordPrint.Values[field.ColumnPaste];
                if (val != null) SheetAnalysis.Cells[rowPaste, field.ColumnPaste].Value = val;
            }
        }

        private Record GetRecocdAnalysis(int row)
        {
            Record recordAnalysis = new Record();
            ColumnMapping mappingNumber = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]);
            object number = SheetAnalysis.Range[$"{mappingNumber.ColumnSymbol}{row}"].Value;
            recordAnalysis.Number = number?.ToString() ?? "";

            foreach (ColumnMapping columnMapping in CurrentProject.Columns)
            {
                if (columnMapping.Check)
                {
                    object val = SheetAnalysis.Range[$"{columnMapping.ColumnSymbol}{row}"].Value;
                    string key = val?.ToString();

                    recordAnalysis.KeyFilds.Add(key);
                }
            }
            return recordAnalysis;
        }


        internal void PrintMarks(List<(ColumnMapping, int)> listColumnPair)
        {
            foreach ((ColumnMapping projectColumn, int offerColumn) pair in listColumnPair)
            {
                SheetAnalysis.Cells[1, pair.projectColumn.Column].Value = pair.projectColumn.Name;
            }
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
            Excel.Range titleTamplate = SheetAnalysis.Range[SheetAnalysis.Cells[7, 1], SheetAnalysis.Cells[8, lastCol]];
            //Excel.Range title = null;
            int columnPaste = ColumnStartPrint;
            foreach (FieldAddress address in addresslist)
            {
                Excel.Range rngCoulumn = titleTamplate.Columns[address.MappingAnalysis.Column];
                rngCoulumn.Copy(SheetAnalysis.Cells[7, columnPaste]);
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
            //Excel.Range commentsTitleRng = tamplateSheet.Range["A10:I14"];
            try
            {
                Excel.Range commentsTitleRng = tamplateSheet.Range["ШаблонКомментарии"];
                commentsTitleRng.Copy(SheetAnalysis.Cells[5, columnPaste]);
            }
            catch (Exception e)
            {
                throw new AddInException($"При копировании диапазона \"ШаблонКомментарии\" возникла ошибка: {e.Message}");
            }
            //if (title != null)
            //{
            //    title.Copy(SheetProjerct.Cells[7, _ColumnStartPrint]);
            //}
        }

        private List<Record> GetListRecordsAnalysis()
        {
            List<Record> records = new List<Record>();
            int lastRow = SheetAnalysis.UsedRange.Rows.Count + SheetAnalysis.UsedRange.Row - 1;
            int lastCol = CurrentProject.Columns.Max(s => s.Column);
            object[,] data = SheetAnalysis.Range[SheetAnalysis.Cells[CurrentProject.RowStart, 1],
                                                   SheetAnalysis.Cells[lastRow, lastCol]];
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

    }
}
