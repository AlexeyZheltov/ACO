using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

                if(pallets.TryGetValue(row.Cells[1,1].Text, out Excel.Range pallet))
                {
                    pallet.Copy();
                    row.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                }
            }
            application.CutCopyMode = (Excel.XlCutCopyMode)0;
        }

        /// <summary>
        /// Группирует строки в зависимости от уровня
        /// </summary>
        /// <param name="ws">Лист с таблицей итоговых данных</param>
        /// <param name="pb"></param>
        public static void Group(Excel.Worksheet ws, IProgressBarWithLogUI pb)
        {
            Excel.Application application = ws.Application;
            int max = (int)application.WorksheetFunction.Max(ws.Range["A:A"]);
            bool flag = false;
            int firstRow = 0, lastRow, currentRow = 0;

            pb.SetSubBarVolume(max * ws.UsedRange.Rows.Count);
            for(int level = 1; level < max; level++)
            {
                if (pb.IsAborted) break;
                foreach (Excel.Range row in ws.UsedRange.Rows)
                {
                    if (pb.IsAborted) break;
                    pb.SubBarTick();
                    currentRow = row.Row;

                    if (row.Row < 10) continue; //ws.Cells[currentRow, 1].Text == ""

                    if(int.TryParse(ws.Cells[currentRow, 1].Text, out int value))
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
        /// Расставляет формулы в зависимости от маркера
        /// </summary>
        /// <remarks>Столбцы куда записывать формулы пока будут захардкорены</remarks>
        /// <param name="ws">Лист в котором проставляются формулы</param>
        /// <param name="markColumn">Столбец с маркерами</param>
        public static void SetFormulas(Excel.Worksheet ws)
        {
            for (int ptr = 10; ptr <= ws.UsedRange.Rows.Count; ptr++)
            {

            }

        }

        /// <summary>
        /// Записывает итоговые данные
        /// </summary>
        /// <param name="ws">Лист шаблона</param>
        /// <param name="data">Считанные и структурированные данные</param>
        /// <param name="markColumn">Столбцец маркеров</param>
        //public static void WriteResult(Excel.Worksheet ws, HierarchyDictionary data, IProgressBarWithLogUI pb)
        //{
        //  //  int ptr = 10;
        //  //  foreach(var item in data.Items())
        //    //{
        //     //   if (pb.IsAborted) break;
        //        //pb.SubBarTick();
        //        //ws.Cells[ptr, 1].Value = item.Level;
        //        //ws.Cells[ptr, 2].Value = item.Number;
        //        //ws.Cells[ptr, 4].Value = item.WorkName;
        //        //ws.Cells[ptr, 5].Value = item.Marking;
        //        //ws.Cells[ptr, 6].Value = item.Material;
        //        //ws.Cells[ptr, 7].Value = item.Format;
        //        //ws.Cells[ptr, 8].Value = item.Type;
        //        //ws.Cells[ptr, 9].Value = item.Article;
        //        //ws.Cells[ptr, 10].Value = item.Maker;
        //        //ws.Cells[ptr, 11].Value = item.Unit;
        //        //ws.Cells[ptr, 12].Value = item.Amount;
        //        //ws.Cells[ptr, 19].Value = item.Note;
        //        //ptr++;
        //   // }
        //}
    }
}
