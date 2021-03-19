using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.ExcelHelpers
{
    /// <summary>
    /// Считывает разнообразные данные с листов
    /// </summary>
    static class ExcelReader
    {
        /// <summary>
        /// Считывает палитру
        /// </summary>
        /// <param name="ws">Лист с которого читается палитра</param>
        /// <returns>Словарь, где ключь это уровень маркера, а значение - ссылка на ячейку с эталонным стилем</returns>
        public static Dictionary<string, Excel.Range> ReadPallet(Excel.Worksheet ws)
        {
            Dictionary<string, Excel.Range> result = new Dictionary<string, Excel.Range>();

            foreach(Excel.Range row in ws.UsedRange.Rows)
            {
                string key = row.Cells[1, 1].Text;
                if (key != "")
                {
                    if (!result.ContainsKey(key))
                        result.Add(key, row.Cells[1, 1]);
                }
            }

            return result;
        }

        /// <summary>
        /// Определяет Верхний уровень в исходной таблице
        /// </summary>
        /// <param name="ws">Лист с которого загружаются данные</param>
        /// <returns>Название верхнего уровня в виде string</returns>
        public static string GetTopLevel(Excel.Worksheet ws) => ws.Range["L1"].Text;

        /// <summary>
        /// Считывает данные с листа
        /// </summary>
        /// <param name="ws">Лист с которого данные читаются</param>
        /// <param name="levelColumn">Буквенное имя колонки в которой проставлены уровни</param>
        /// <param name="startRow">Стартовая строка</param>
        /// <returns></returns>
        public static IEnumerable<(int Row, int Level)> ReadSourceItems(Excel.Worksheet ws, string levelColumn, int startRow = 1)
        {
            foreach(Excel.Range r_row in ws.UsedRange.Rows)
            {
                int row = r_row.Row;
                if (row < startRow) continue;
                string str_lvl = ws.Range[$"{levelColumn}{row}"].Text;
                if (int.TryParse(str_lvl, out int lvl))
                    yield return (row, lvl);
            }
        }
    }
}
