using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;
using ACO.ProjectManager;
using System.Windows.Forms;
using ACO.ExcelHelpers;

namespace ACO
{
    /// <summary>
    ///  Загружает данные на лист Анализ
    /// </summary>
    class ListAnalysis
    {

        public Excel.Worksheet SheetAnalysis { get; }
        public Project CurrentProject { get; }

        /// <summary>
        ///  Столбец для вставки загруд
        /// </summary>
        public int ColumnStartPrint
        {
            get
            {
                if (_ColumnStartPrint == 0)
                {
                    _ColumnStartPrint = SheetAnalysis.UsedRange.Column + SheetAnalysis.UsedRange.Columns.Count + 1;
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
        int _rowStart = 1;
        int _lastRow = 1;

        public ListAnalysis(Excel.Worksheet sheetProjerct, Project currentProject)
        {
            SheetAnalysis = sheetProjerct;
            CurrentProject = currentProject;
            _rowStart = CurrentProject.RowStart;
            _lastRow = SheetAnalysis.UsedRange.Row + SheetAnalysis.UsedRange.Rows.Count;
        }

        /// <summary>
        ///  Запись строки КП на лист Анализ. Вставка строк.
        /// </summary>
        /// <param name="recordPrint"></param>
        internal void Print(Record recordPrint)
        {
            int rowPaste = _rowStart;

            /// Последняя строка списка 

            bool existRecord = false;
            Record recordAnalysis = null;
            for (int row = _rowStart; row <= _lastRow; row++)
            {
                recordAnalysis = GetRecocdAnalysis(_rowStart);
                if (!string.IsNullOrEmpty(recordAnalysis.Number))
                {
                    _rowStart = row;
                    existRecord = true;
                    break;
                }
            }
            if (recordAnalysis != null && existRecord)
            {
                // Проверка ключевых значений 
                if (!recordAnalysis.KeyEqual(recordPrint))
                {
                    SheetAnalysis.Rows[_rowStart].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    _lastRow++;
                }
            }

            /// Печать значений
            foreach (FieldAddress field in recordPrint.Addresslist)
            {
                object val = recordPrint.Values[field.ColumnPaste];
                Excel.Range cell = SheetAnalysis.Cells[rowPaste, field.ColumnPaste];
                if (val != null)
                { // Ошибка формулы в загружаемом файле
                    if (double.TryParse(val.ToString(), out double dv))
                    {
                        if (dv < 0) cell.Interior.Color = System.Drawing.Color.FromArgb(176, 119, 237);                       
                        cell.NumberFormat = "#,##0.##";                        
                        cell.Value = Math.Round(dv, 2);
                    }
                    else
                    {
                        cell.Value = val;
                    }
                }
            }
            _rowStart++;
        }

        /// <summary>
        ///  Считывает поля для проверки с листа Анализ.
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
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

        /// <summary>
        ///  Копирование заголовков
        /// </summary>
        /// <param name="addresslist"></param>
        internal void PrintTitle(Excel.Worksheet tamplateSheet, List<FieldAddress> addresslist)
        {
            int lastCol = addresslist.Last().MappingAnalysis.Column;
                              

            foreach (FieldAddress address in addresslist)
            {
                int col = address.MappingAnalysis.Column;
                if (col > lastCol) lastCol = col;
            }
            Excel.Range titleTamplate = SheetAnalysis.Range[SheetAnalysis.Cells[7, 1], SheetAnalysis.Cells[8, lastCol]];
            int columnPaste = ColumnStartPrint;
            foreach (FieldAddress address in addresslist)
            {
                Excel.Range rngCoulumn = titleTamplate.Columns[address.MappingAnalysis.Column];
                rngCoulumn.Copy(SheetAnalysis.Cells[7, columnPaste]);
                SheetAnalysis.Cells[1, columnPaste].Value = address.MappingAnalysis.Name;

                SheetAnalysis.Cells[7, columnPaste].Copy();
                SheetAnalysis.Cells[9, columnPaste].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                if (address.MappingAnalysis.Name == Project.ColumnsNames[StaticColumns.CostMaterialsTotal] ||
                    address.MappingAnalysis.Name == Project.ColumnsNames[StaticColumns.CostWorksTotal])
                {
                    SheetAnalysis.Range[SheetAnalysis.Cells[7, columnPaste - 1], SheetAnalysis.Cells[7, columnPaste]].Merge();                   
                }
                columnPaste++;
            }
            /// Цвет шапки
            Excel.Range pallet = SheetAnalysis.Cells[6, 1];
            //Top
            
            Excel.Range rng = SheetAnalysis.Range[SheetAnalysis.Cells[6, ColumnStartPrint], SheetAnalysis.Cells[6, columnPaste - 1]];
            // Globals.ThisAddIn.Application.ScreenUpdating = true;            
            rng.EntireColumn.AutoFit();
            //Globals.ThisAddIn.Application.ScreenUpdating = false;
            pallet.Copy();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            rng.Merge();

            //Bottom
          //  rng = SheetAnalysis.Range[SheetAnalysis.Cells[9, ColumnStartPrint], SheetAnalysis.Cells[9, columnPaste - 1]];
          //  pallet.Copy();
           // rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            // Left
            rng = SheetAnalysis.Range[SheetAnalysis.Cells[6, ColumnStartPrint - 1], SheetAnalysis.Cells[9, ColumnStartPrint - 1]];
            pallet.Copy();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            /// Участник
            


            SheetAnalysis.Cells[1, ColumnStartPrint - 1].Value = "offer_start";
            SheetAnalysis.Cells[1, columnPaste].Value = "offer_end";
            try
            {
                Excel.Range commentsTitleRng = tamplateSheet.Range["ШаблонКомментарии"];
                commentsTitleRng.Copy();
                Excel.Range rngPaste = SheetAnalysis.Cells[5, columnPaste];
                rngPaste.PasteSpecial(Excel.XlPasteType.xlPasteAll);

               
            }
            catch (Exception e)
            {
                throw new AddInException($"При копировании диапазона \"ШаблонКомментарии\" возникла ошибка: {e.Message}");
            }
        }
    }
}
