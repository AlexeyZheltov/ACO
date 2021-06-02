using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using ACO.Offers;
using ACO.ProjectManager;
using System.Windows.Forms;
using ACO.ExcelHelpers;
using ACO.ProjectBook;

namespace ACO
{
    /// <summary>
    ///  Загружает данные на лист Анализ
    /// </summary>
    class ListAnalysis
    {
        public Excel.Worksheet SheetAnalysis { get; }
        public Project CurrentProject { get; }

        int _rowStart = 1;
        int _lastRow = 1;
        //  private readonly List<Record> _levelRecords;
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

        public ListAnalysis(Excel.Worksheet sheetProjerct, Project currentProject)
        {
            SheetAnalysis = sheetProjerct;
            CurrentProject = currentProject;
            _rowStart = CurrentProject.RowStart;
            _lastRow = SheetAnalysis.UsedRange.Row + SheetAnalysis.UsedRange.Rows.Count - 1;
        }

        /// <summary>
        ///  Сопоставление полей для вставки строки КП
        /// </summary>
        /// <param name="recordPrint"></param>
        public void PrintRecord(Record recordPrint)
        {
            int rowNextLvl;
            int row = _rowStart;
            Record recordAnalysis = GetRecocdAnalysis(_rowStart);

            /// Последний уровень ищем пока уровень не изменится на любой другой
            if (recordPrint.Level == 0) return;
            if (recordPrint.Level == 6)
            {
                rowNextLvl = FindNextLevelRow(recordAnalysis.Level);
                if (PrintEqualRecord(recordPrint, row, rowNextLvl)) return;
            }
            else
            {
                /// Уровень не последний, ищем пока не изменится на меньший чем текущий
                rowNextLvl = FindPevLvlRow(recordPrint.Level);
                int rowPrevLvl = FindPrevLvlBackRow(recordPrint.Level);
                if (PrintEqualRecord(recordPrint, rowPrevLvl, rowNextLvl)) return;
                _rowStart = rowNextLvl;
            }

            // Вставка новой строки
            SheetAnalysis.Rows[rowNextLvl].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            PrintValues(recordPrint, rowNextLvl);
            SetLevel(recordPrint.Level, rowNextLvl);
        }

        private bool PrintEqualRecord(Record recordPrint, int rowStart, int rowEnd)
        {
            if (rowEnd >= rowStart)
            {
                for (int i = rowStart; i <= rowEnd; i++)
                {
                    Record recordAnalysis = GetRecocdAnalysis(i);
                    if (recordAnalysis.Level == recordPrint.Level && recordAnalysis.KeyEqual(recordPrint))
                    {
                        PrintValues(recordPrint, i);
                        _rowStart = i + 1;
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        ///  Вставлет в базовый список номер уровня добавленной строки
        /// </summary>
        /// <param name="lvl"></param>
        /// <param name="row"></param>
        private void SetLevel(int lvl, int row)
        {
            string letterLevel = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            SheetAnalysis.Range[$"{letterLevel}{row}"].Value = lvl.ToString();
        }

        private int FindPrevLvlBackRow(int level)
        {
            string letterLevel = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            if (level > 0)
            {
                for (int row = _rowStart; row >= CurrentProject.RowStart; row--)
                {
                    string text = SheetAnalysis.Range[$"{letterLevel}{row}"].Value?.ToString() ?? "";
                    if (int.TryParse(text, out int lvl))
                    {
                        if (level > lvl) return row;
                    }
                }
            }
            return _rowStart;
        }

        /// <summary>
        /// 
        /// </summary>
        private int FindNextLevelRow(int level)
        {
            string letterLevel = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;

            if (level > 0)
            {
                for (int row = _rowStart; row <= _lastRow; row++)
                {
                    string text = SheetAnalysis.Range[$"{letterLevel}{row}"].Value?.ToString() ?? "";
                    if (int.TryParse(text, out int lvl))
                    {
                        if (level != lvl) return row;
                    }
                }
            }
            _lastRow++;
            return _lastRow;
        }

        /// <summary>
        /// Поиск строки указанного уровня вниз по списку, если не найдена возвращает 
        /// </summary>
        /// <param name="level"></param>
        /// <param name="rowNextLevel"></param>
        /// <returns></returns>
        private int FindPevLvlRow(int level)
        {
            string letterLevel = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;

            if (level > 0)
            {
                for (int row = _rowStart; row <= _lastRow; row++)
                {
                    string text = SheetAnalysis.Range[$"{letterLevel}{row}"].Value?.ToString() ?? "";
                    if (int.TryParse(text, out int lvl))
                    {
                        if (level > lvl) return row;
                    }
                }
            }
            _lastRow++;
            return _lastRow;
        }


        private void PrintValues(Record recordPrint, int rowPaste)
        {
            foreach (FieldAddress field in recordPrint.Addresslist)
            {
                object val = recordPrint.Values[field.ColumnPaste];
                Excel.Range cell = SheetAnalysis.Cells[rowPaste, field.ColumnPaste];
                if (val != null)
                { // Ошибка формулы в загружаемом файле
                    if (double.TryParse(val.ToString(), out double dv))
                    {
                        if (dv < 0) cell.Interior.Color = System.Drawing.Color.FromArgb(176, 119, 237);
                        // cell.NumberFormat = "#,##0.00";
                        cell.Value = Math.Round(dv, 2);
                    }
                    else
                    {
                        cell.Value = val;
                    }
                }
            }
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
                recordAnalysis = GetRecocdAnalysis(row);
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
            ColumnMapping mappingLevel = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]);
            object number = SheetAnalysis.Range[$"{mappingNumber.ColumnSymbol}{row}"].Value;
            recordAnalysis.Number = number?.ToString() ?? "";

            string levelTtext = SheetAnalysis.Range[$"{mappingLevel.ColumnSymbol}{row}"].Value?.ToString() ?? "";
            levelTtext = levelTtext.Trim();
            recordAnalysis.Level = int.TryParse(levelTtext, out int lvl) ? lvl : 0;
            foreach (ColumnMapping columnMapping in CurrentProject.Columns)
            {
                if (columnMapping.Check)
                {
                    object val = SheetAnalysis.Range[$"{columnMapping.ColumnSymbol}{row}"].Value;
                    string key = val?.ToString() ?? "";
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
                SheetAnalysis.Cells[9, columnPaste].PasteSpecial(Excel.XlPasteType.xlPasteFormats,
                                            Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

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
            pallet.Copy();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            rng.Merge();
            SheetAnalysis.Cells[6, ColumnStartPrint].HorizontalAlignment = HorizontalAlignment.Center;
            // Left
            rng = SheetAnalysis.Range[SheetAnalysis.Cells[6, ColumnStartPrint - 1], SheetAnalysis.Cells[9, ColumnStartPrint - 1]];
            pallet.Copy();
            rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

            /// Метки
            SheetAnalysis.Cells[1, ColumnStartPrint - 1].Value = "offer_start";
            SheetAnalysis.Cells[1, columnPaste].Value = "offer_end";
            ColumnCommentsMark(columnPaste);
            rng = SheetAnalysis.Range[SheetAnalysis.Cells[6, ColumnStartPrint], SheetAnalysis.Cells[6, columnPaste + 8]];
            rng.EntireColumn.AutoFit();

            SheetAnalysis.Rows[1].Hidden = true;

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

        /// <summary>
        ///  Метки комментариев 
        /// </summary>
        /// <param name="column"></param>
        private void ColumnCommentsMark(int column)
        {
            SheetAnalysis.Cells[1, column + 1].Value = ColumnCommentsValues[StaticColumnsComments.CommentDiscriptionWorks];//  "Комментарии к описанию работ";
            SheetAnalysis.Cells[1, column + 2].Value = ColumnCommentsValues[StaticColumnsComments.DeviationVolume]; //"Отклонение по объемам";
            SheetAnalysis.Cells[1, column + 3].Value = ColumnCommentsValues[StaticColumnsComments.CommentsVolume];  //"Комментарии к объемам работ";
            SheetAnalysis.Cells[1, column + 4].Value = ColumnCommentsValues[StaticColumnsComments.DeviationCost];   //"Отклонение по стоимости";
            SheetAnalysis.Cells[1, column + 5].Value = ColumnCommentsValues[StaticColumnsComments.CommentsCost];    //"Комментарии к стоимости работ";
            SheetAnalysis.Cells[1, column + 6].Value = ColumnCommentsValues[StaticColumnsComments.DeviationMat];    //"Отклонение МАТ";
            SheetAnalysis.Cells[1, column + 7].Value = ColumnCommentsValues[StaticColumnsComments.DeviationWorks];  //"Отклонение РАБ";
            SheetAnalysis.Cells[1, column + 8].Value = ColumnCommentsValues[StaticColumnsComments.Comments];        //"Комментарии к строкам";
        }

        /*
          "Комментарии к описанию работ";
          "Отклонение по объемам";
          "Комментарии к объемам работ";
          "Отклонение по стоимости";
          "Комментарии к стоимости работ";
          "Отклонение МАТ";
          "Отклонение РАБ";
          "Комментарии к строкам";
        */
        public static Dictionary<StaticColumnsComments, string> ColumnCommentsValues =
              new Dictionary<StaticColumnsComments, string>()
              {
                  {StaticColumnsComments.CommentDiscriptionWorks, "Комментарии к описанию работ" },
                  {StaticColumnsComments.DeviationVolume,   "Отклонение по объемам" },
                  {StaticColumnsComments.CommentsVolume,    "Комментарии к объемам работ" },
                  {StaticColumnsComments.DeviationCost,     "Отклонение по стоимости" },
                  {StaticColumnsComments.CommentsCost ,     "Комментарии к стоимости работ" },
                  {StaticColumnsComments.DeviationMat ,     "Отклонение МАТ" },
                  {StaticColumnsComments.DeviationWorks ,   "Отклонение РАБ" },
                  {StaticColumnsComments.Comments,          "Комментарии к строкам" }
              };

        public void GroupColumn()
        {
            ExcelHelper.UnGroupColumns(SheetAnalysis);
            GroupColumnsBasis();
            GroupOfferColumns();
            List<OfferColumns> addresslist = new ProjectWorkbook().GetOfferColumns();
            foreach (OfferColumns address in addresslist)
            {
                GroupCommetnsColumns(address);
            }
        }
        private void GroupOfferColumns()
        {
            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            foreach (OfferColumns address in projectWorkbook.OfferColumns)
            {
                Excel.Range rngCoulumn = SheetAnalysis.Range[SheetAnalysis.Cells[1, address.ColStartOffer + 1], SheetAnalysis.Cells[1, address.ColCostTotalOffer - 1]];
                rngCoulumn.Columns.Group();
            }
        }
        private void GroupCommetnsColumns(OfferColumns address)
        {
            //Excel.Range rngCoulumn = SheetAnalysis.Cells[1, address.ColStartOfferComments-1];
            //rngCoulumn.Columns.Group();

            Excel.Range rngColumn = SheetAnalysis.Range[SheetAnalysis.Cells[1, address.ColStartOfferComments], SheetAnalysis.Cells[1, address.ColDeviationCost - 1]];
            rngColumn.Columns.Group();
        }

        /// <summary>
        ///  Группировать столбцы базовой оценки
        /// </summary>
        public void GroupColumnsBasis()
        {
            // Группировать столбцы A:B
            string letterNumber = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            Excel.Range rng = SheetAnalysis.Range[$"{letterNumber}1"];
            int colNumber = rng.Column;
            rng = SheetAnalysis.Range[SheetAnalysis.Cells[1, 1],
                                      SheetAnalysis.Cells[1, colNumber - 1]];
            rng.Columns.Group();

            string letterBasisComment = CurrentProject.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Comment]).ColumnSymbol;
            SheetAnalysis.Range[$"{letterBasisComment}1"].Columns.Group();

            int colCostTotal = ExcelHelper.GetColumn(CurrentProject.Columns.Find(
                x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol, SheetAnalysis);
            int colNames = ExcelHelper.GetColumn(CurrentProject.Columns.Find(
               x => x.Name == Project.ColumnsNames[StaticColumns.Name]).ColumnSymbol, SheetAnalysis);
            int colEndBasis = ExcelHelper.GetColumn(CurrentProject.Columns.Find(
               x => x.Name == Project.ColumnsNames[StaticColumns.Comment]).ColumnSymbol, SheetAnalysis);

            // Комментарий базовой оценки
            // rng = SheetAnalysis.Cells[1, colCostTotal + 1];
            // rng.Columns.Group();

            string letterUnit = CurrentProject.Columns.
                            Find(x => x.Name == Project.ColumnsNames[StaticColumns.Unit]).ColumnSymbol;
            Excel.Range cellUnit = SheetAnalysis.Range[$"{letterUnit}1"];

            // доп аттрибуты базовой оценки
            rng = SheetAnalysis.Range[SheetAnalysis.Cells[1, colNames + 1], cellUnit];
            rng.Columns.Group();

            //Базовая оценка стоимости
            int colCostMaterialsPerUnit = colCostTotal - 5;
            rng = SheetAnalysis.Range[SheetAnalysis.Cells[1, colCostMaterialsPerUnit],
                                      SheetAnalysis.Cells[1, colCostTotal - 1]];
            rng.Columns.Group();

            rng = SheetAnalysis.Range[SheetAnalysis.Cells[1, colEndBasis + 1],
                                      SheetAnalysis.Cells[1, colEndBasis + 7]];
            rng.Columns.Group();
        }
    }

    public enum StaticColumnsComments
    {
        CommentDiscriptionWorks,
        DeviationVolume,
        CommentsVolume,
        DeviationCost,
        CommentsCost,
        DeviationMat,
        DeviationWorks,
        Comments
    }
}
