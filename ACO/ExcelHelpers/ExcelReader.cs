using System;
using System.Collections.Generic;
using System.Linq;
using Spectrum.SpLoader.XMLSetting;
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
        /// Получает данные из Omni файла
        /// </summary>
        /// <param name="ws">Лист с которого читаются данные о Omni классах</param>
        /// <returns>Словарь, где ключ - это OmniClass, а значение структура в виде списка.</returns>
        public static Dictionary<string, string[]> ReadOmni(Excel.Worksheet ws, IProgressBarWithLogUI pb)
        {
            Dictionary<string, string[]> result = new Dictionary<string, string[]>();
            string[] buffer = new string[4];

            pb.SetSubBarVolume(ws.UsedRange.Rows.Count);
            foreach (Excel.Range row in ws.UsedRange.Rows)
            {
                if (pb.IsAborted) break;
                pb.SubBarTick();
                if (row.Row == 1) continue;
                string omni = row.Cells[1, 1].Text;
                omni = omni.Trim().Replace("\n", "");
                if (!String.IsNullOrEmpty(omni))
                {
                    for (int i = 3; i > -1; i--)
                    {
                        string temp = row.Cells[1, i + 2].Text;
                        temp = temp.Trim().Replace("\n", "");
                        buffer[i] = temp;
                        if (!string.IsNullOrEmpty(temp))
                            break;
                    }
                   // if (!buffer.All(x => string.IsNullOrEmpty(x)) && !result.ContainsKey(omni))
                   //     result.Add(omni, buffer.CopyFilled());
                }
            }

            return result;
        }

        public static string GetTopLevel(Excel.Worksheet ws) => ws.Range["L1"].Text;

        /// <summary>
        /// Считывает ресурсные файлы
        /// </summary>
        /// <remarks>Выполнить через yield return</remarks>
        /// <returns>Итератор с типов TargetItem</returns>
        //public static IEnumerable<TargetItem> ReadSourceItems(Excel.Worksheet ws, Mapping mapping, string[] omniClasses)
        //{
        //    foreach(Excel.Range row in ws.UsedRange.Rows)
        //    {
        //        string key = ws.Range[$"{mapping.Omni}{row.Row}"].Text;
        //        if (key != "" && omniClasses.Contains(key))
        //            yield return TrimAllProp(new TargetItem
        //            {
        //                OmniClass = ws.Range[$"{mapping.Omni}{row.Row}"].Text,
        //                WorkName = ws.Range[$"{mapping.WorkName}{row.Row}"].Text,
        //                Marking = ws.Range[$"{mapping.Marking}{row.Row}"].Text,
        //                Material = ws.Range[$"{mapping.Material}{row.Row}"].Text,
        //                Format = ws.Range[$"{mapping.Format}{row.Row}"].Text,
        //                Type = ws.Range[$"{mapping.Type}{row.Row}"].Text,
        //                Article = ws.Range[$"{mapping.Article}{row.Row}"].Text,
        //                Maker = ws.Range[$"{mapping.Maker}{row.Row}"].Text,
        //                Unit = ws.Range[$"{mapping.Unit}{row.Row}"].Text,
        //                Amount = ws.Range[$"{mapping.Amount}{row.Row}"].Text,
        //                Note = ws.Range[$"{mapping.Note}{row.Row}"].Text
        //            });
        //        else yield return null;
        //    }
        //}

        //private static TargetItem TrimAllProp(TargetItem item)
        //{
        //    item.OmniClass = item.OmniClass.Trim();
        //    item.WorkName = item.WorkName.Trim();
        //    item.WorkName = item.WorkName.Trim();
        //    item.Material = item.Material.Trim();
        //    item.Format = item.Format.Trim();
        //    item.Type = item.Type.Trim();
        //    item.Article = item.Article.Trim();
        //    item.Maker = item.Maker.Trim();
        //    item.Unit = item.Unit.Trim();
        //    item.Amount = item.Amount.Trim();
        //    item.Note = item.Note.Trim();
        //    return item;
        //}
    }
}
