using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ACO.ExcelHelpers
{
    /// <summary>
    /// Ксласс со вспомогательным функционал для работы с Excel
    /// </summary>
    static class ExcelHelper
    {
        /// <summary>
        /// Перекрашивает таблицу, используя палитру
        /// </summary>
        /// <remarks>В палитре Range, что бы копировать весь стиль оформления установленный пользователем на листе палитры</remarks>
        /// <param name="ws">Лист в котором будет произведена закраска</param>
        /// <param name="pallet">Палитра</param>
        public static void Repaint(Excel.Worksheet ws, Dictionary<string, Excel.Range> pallets, IProgressBarWithLogUI pb)
        {
            Excel.Application application = ws.Application;

            pb.SetSubBarVolume(ws.UsedRange.Rows.Count);
            foreach (Excel.Range row in ws.UsedRange.Rows)
            {
                if (pb.IsAborted) break;
                pb.SubBarTick();

                if (pallets.TryGetValue(row.Cells[1, 1].Text, out Excel.Range pallet))
                {
                    pallet.Copy();
                    row.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                }
            }
            application.CutCopyMode = (Excel.XlCutCopyMode)0;
        }

        /// <summary>
        /// Перекрашивает таблицу, используя палитру. Для второго этапа
        /// </summary>
        /// <param name="ws">Лист в котором будет произведена закраска</param>
        /// <param name="pallets">Палитра</param>
        /// <param name="startRow">Начальная строка с которой идет перекраска</param>
        /// <param name="levelColumn">Буквенное имя столбца в котором проставлены уровни</param>
        /// <param name="pb">Прогресс бар. Колличество нужно инициализировать до вызова метода</param>
        /// <param name="columns">Набор юуквенных имен колонок (нач, кон), (начб кон)... в которых будет производится покраска</param>
        public static void Repaint(Excel.Worksheet ws, Dictionary<string, Excel.Range> pallets, int startRow, string levelColumn, IProgressBarWithLogUI pb, params (string, string)[] columns)
        {
           // Excel.Application application = ws.Application;

            foreach (Excel.Range r_row in ws.UsedRange.Rows)
            {
                if (pb?.IsAborted ?? false) break;
                pb?.SubBarTick();

                int row = r_row.Row;
                if (row < startRow)
                    continue;

                if (pallets.TryGetValue(ws.Range[$"{levelColumn}{row}"].Text, out Excel.Range pallet))
                {
                    //pallet.Copy();
                    foreach (var columns_pair in columns)
                    {
                        (string f_column, string l_column) = columns_pair;
                        SetCellFormat(ws.Range[$"{f_column}{row}:{l_column}{row}"], pallet);
                        //ws.Range[$"{f_column}{row}:{l_column}{row}"].PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    }
                }
            }
           // application.CutCopyMode = (Excel.XlCutCopyMode)0;
        }

        public static void SetCellFormat(Excel.Range cell,Excel.Range cellFormat)
        {
          cell.Interior.Color = cellFormat.Interior.Color;
          cell.Font.Name = cellFormat.Font.Name;
          cell.Font.Bold = cellFormat.Font.Bold;
          cell.Font.Color = cellFormat.Font.Color;
        }

        /// <summary>
        /// Группирует строки в зависимости от уровня
        /// </summary>
        /// <param name="ws">Лист с таблицей итоговых данных</param>
        /// <param name="pb"></param>
        public static void Group(Excel.Worksheet ws, IProgressBarWithLogUI pb, string letterLevel)
        {
            Excel.Application application = ws.Application;
            int max = (int)application.WorksheetFunction.Max(ws.Range[$"{letterLevel}:{letterLevel}"]);
            bool flag = false;
            int firstRow = 0, lastRow, currentRow = 0;

            pb.SetSubBarVolume(max * ws.UsedRange.Rows.Count);
            for (int level = 1; level < max; level++)
            {
                if (pb.IsAborted) break;
                foreach (Excel.Range row in ws.UsedRange.Rows)
                {
                    if (pb.IsAborted) break;
                    pb.SubBarTick();
                    currentRow = row.Row;
                    if (row.Row < 10) continue; 

                    if (int.TryParse(ws.Cells[currentRow, 1].Text, out int value))
                    {
                        if (!flag && value == level)
                        {
                            firstRow = currentRow + 1;
                            flag = true;
                        }
                        else if (flag && value <= level)
                        {
                            lastRow = currentRow - 1;
                            if (lastRow - firstRow > -1)
                                ws.Range[$"{firstRow}:{lastRow}"].Rows.Group();

                            if (value == level)
                                firstRow = currentRow + 1;
                            else
                                flag = false;
                        }
                    }
                }
                if (flag && currentRow > firstRow)
                    ws.Range[$"{firstRow}:{currentRow}"].Rows.Group();
                flag = false;
            }
        }

        /// <summary>
        /// Разгрупировывает группировку по строкам
        /// </summary>
        /// <param name="ws">Лист для разгруппировки</param>
        public static void UnGroupRows(Excel.Worksheet ws)
        {
            try
            {
                while (true)
                    ws.UsedRange.Rows.Ungroup();
            }
            catch { }
        }
        public static void UnGroupColumns(Excel.Worksheet ws)
        {
            try
            {
                while (true)
                    ws.UsedRange.Columns.Ungroup();
            }
            catch { }
        }


        /// <summary>
        /// поиск ячейки по тексту 
        /// </summary>
        /// <param name="sh"></param>
        /// <param name="findText"></param>
        /// <returns></returns>
        internal static Excel.Range FindCell(Excel.Worksheet sh, string findText)
        {
            Excel.Range cell = sh.UsedRange.Find(findText);
            if (cell is null) throw new AddInException($"Не удалось найти ячейку с текстом: \"{findText}\" на листе: {sh.Name}");
            return cell;
        }

        /// <summary>
        /// Расставляет формулы в зависимости от маркера
        /// </summary>
        /// <param name="ws">Лист в котором проставляются формулы</param>
        /// <param name="mapping">маппинг для формул по типу FMapping</param>
        /// <param name="data"></param>
        /// <param name="pb">Прогресс бар, служит только для отлавливания нажатия кнопки отмена</param>
        public static void SetFormulas(Excel.Worksheet ws, FMapping mapping, HItem data, IProgressBarWithLogUI pb)
        {
            if (pb?.IsAborted ?? false) return;

            StringBuilder builder = new StringBuilder();

            var lvl = data.GetSubLevel();

            foreach (var item in lvl)
            {
                if (pb?.IsAborted ?? false) break;
                builder.Clear();
                builder.Append("=");

                var sub_lvl = item.GetSubLevel();

                if (sub_lvl.Count == 0)
                {
                    int s_row = item.Row;
                    ws.Range[$"{mapping.MaterialTotal}{s_row}"].Formula = $"=ROUND({mapping.MaterialPerUnit}{s_row}*{mapping.Amount}{s_row},2)";
                    ws.Range[$"{mapping.WorkTotal}{s_row}"].Formula = $"=ROUND({mapping.WorkPerUnit}{s_row}*{mapping.Amount}{s_row},2)";
                    ws.Range[$"{mapping.PricePerUnit}{s_row}"].Formula = $"=ROUND({mapping.MaterialPerUnit}{s_row}+{mapping.WorkPerUnit}{s_row},2)";
                    ws.Range[$"{mapping.Total}{s_row}"].Formula = $"=ROUND({mapping.PricePerUnit}{s_row}*{mapping.Amount}{s_row},2)";
                    //ws.Range[$"{mapping.MaterialTotal}{s_row}"].NumberFormat = "# ##0,00";
                    //ws.Range[$"{mapping.WorkTotal}{s_row}"].NumberFormat = "# ##0,00";
                    //ws.Range[$"{mapping.PricePerUnit}{s_row}"].NumberFormat = "# ##0,00";
                    //ws.Range[$"{mapping.Total}{s_row}"].NumberFormat = "# ##0,00";
                    continue;
                }

                if (sub_lvl.IsSolid())
                    builder.Append($"SUM({mapping.Total}{sub_lvl.First().Row}:{mapping.Total}{sub_lvl.Last().Row})");
                else
                {
                        foreach (var sub_item in sub_lvl)
                        builder.Append($"{mapping.Total}{sub_item.Row}+");

                    builder.Remove(builder.Length - 1, 1);
                }

                int t_row = item.Row;
                ws.Range[$"{mapping.Total}{t_row}"].Formula = builder.ToString();

                if (item.Level > 1)
                {
                    ws.Range[$"{mapping.PricePerUnit}{t_row}"].Formula = $"=ROUND({mapping.Total}{t_row}/{mapping.Amount}{t_row},2)";
                   // ws.Range[$"{mapping.PricePerUnit}{t_row}"].NumberFormat = "# ##0,00";
                }

                SetFormulas(ws, mapping, item, pb);
            }
        }

        /// <summary>
        /// Записывает новыю нумерацию
        /// </summary>
        /// <param name="ws">Лист куда записываются номера</param>
        /// <param name="data"></param>
        /// <param name="pb">Прогресс бар. Колличество нужно инициализировать до вызова метода</param>
        /// <param name="column">Буквенное имя колонки куда записывать номера</param>
        public static void Write(Excel.Worksheet ws, HItem data, IProgressBarWithLogUI pb, string column = "A")
        {
            foreach (var item in data.Items())
            {
                if (pb?.IsAborted ?? false) break;
                pb?.SubBarTick();
                ws.Range[$"{column}{item.Row}"].Value = item.Number;
            }
        }

        /// <summary>
        /// Маркировка ошибочных ячеек.
        /// </summary>
        /// <param name="ws">Лист в котором эти ошибки надо промаркировать</param>
        /// <param name="errorAddress">Список адрессов ячеек с ошибками</param>
        /// <param name="pb">Прогрессбар</param>
        public static void MarkErrors(Excel.Worksheet ws, List<string> errorAddress, IProgressBarWithLogUI pb)
        {
            foreach (var addr in errorAddress)
            {
                if (pb.IsAborted) break;
                pb.SubBarTick();
                ws.Range[addr].Interior.Color = Color.FromArgb(0, 172, 117, 213);
            }
        }

        /// <summary>
        ///  Получить лист по имени
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public static Excel.Worksheet GetSheet(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
            {
                if (sh.Name == name)
                {
                    return sh;
                }
            }
            throw new AddInException($"Лист {name} отсутствует");
        }

        /// <summary>
        ///  Получить текст из ячейки
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        //public static string GetText(Excel.Range cell)
        //{
        //    bool IsXLCVErr(object obj)
        //    {
        //        return (obj) is Int32; // Ошибка Формулы Excel
        //    }
        //    string text = "";
        //    Excel.Application app = Globals.ThisAddIn.Application;          
        //    if (!IsXLCVErr(cell.Value))
        //    {
        //        text = cell?.Value?.ToString() ?? "";
        //    }
        //    return text;
        //}

        /// <summary>
        /// Ячейка в ржиме редактирования
        /// </summary>
        /// <returns></returns>
        public static bool IsEditing()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            if (excelApp.Interactive)
            {
                try
                {
                    excelApp.Interactive = false;
                    excelApp.Interactive = true;
                }
                catch (Exception)
                {
                    MessageBox.Show("Завершите редактирование ячейки", "Ячйка в режиме редактирования",
                       MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return true;
                }
            }
            return false;
        }

        public static string GetColumnLetter(Excel.Range cell)
        {
            string address = cell.Address;
            string letter = address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
            return letter;
        }

        /// <summary>
        ///  Номер стодбца по его буквенному обозначению
        /// </summary>
        /// <param name="columnSymbol"></param>
        /// <param name="sh"></param>
        /// <returns></returns>
        public static int GetColumn(string columnSymbol, Excel.Worksheet sh)
        {
            int col = sh.Range[$"{columnSymbol}1"].Column;
            return col;
        }

        internal static void SetNumberFormat(Worksheet ws, int rowStart, (string, string)[] columns)
        {
            int lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
            if (lastRow <= rowStart) return;
            foreach ((string, string) itm in columns)
            {
                Excel.Range rng = ws.Range[$"{itm.Item1}{rowStart}:{itm.Item2}{lastRow}"];
                rng.NumberFormat = "#,##0.00";
            }
        }

        internal static void SetNumberFormat(Worksheet ws, int rowStart, string letterAmount)
        {
            int lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1;
            if (lastRow <= rowStart) return;            
                Excel.Range rng = ws.Range[$"{letterAmount}{rowStart}:{letterAmount}{lastRow}"];
                rng.NumberFormat = "#,##0.00";
        }

        /// <summary>
        /// Удалить условное форматирование
        /// </summary>
        /// <param name="rng"></param>
        internal static void ClearFormatConditions(Excel.Range rng)
        {
            rng.FormatConditions.Delete();
        }
    }
}
