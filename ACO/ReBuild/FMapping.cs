using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO
{
    /// <summary>
    /// Класс маппинга для повторной вставки формул
    /// </summary>
    class FMapping
    {
        /// <summary>
        /// Кол-во
        /// </summary>
        public string Amount { get; set; }

        /// <summary>
        /// Цена метериалов за еденицу
        /// </summary>
        public string MaterialPerUnit { get; set; }

        /// <summary>
        /// Цена материалов всего
        /// </summary>
        public string MaterialTotal { get; set; }

        /// <summary>
        /// Цена работ за еденицу
        /// </summary>
        public string WorkPerUnit { get; set; }

        /// <summary>
        /// Цена работ всего
        /// </summary>
        public string WorkTotal { get; set; }

        /// <summary>
        /// Цена за еденицу
        /// </summary>
        public string PricePerUnit { get; set; }

        /// <summary>
        /// Итого
        /// </summary>
        public string Total { get; set; }

        /// <summary>
        /// Сдвигает маппинг на указанное колличество столбцов
        /// </summary>
        /// <param name="Shift">Величина сдвига</param>
        /// <param name="ws">Лист к которому относится маппинг</param>
        /// <returns>Новый маппинг столбцоы</returns>
        public FMapping Shift(Excel.Worksheet ws, int columnOfAmount)
        {
            int firstColAmount = ws.Range[$"{Amount}1"].Column;
            int shift = columnOfAmount - firstColAmount ;
            string GetShifted(string CName)
            {
                int col = ws.Range[$"{CName}1"].Column + shift;
                string address = ws.Cells[1, col].Address;
                return address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
            }

            FMapping result = new FMapping
            {
                Amount = GetShifted(Amount),
                MaterialPerUnit = GetShifted(MaterialPerUnit),
                MaterialTotal = GetShifted(MaterialTotal),
                WorkPerUnit = GetShifted(WorkPerUnit),
                WorkTotal = GetShifted(WorkTotal),
                PricePerUnit = GetShifted(PricePerUnit),
                Total = GetShifted(Total)
            };

            return result;
        }

        /// <summary>
        /// Преобразует номер столбца в буквенное представление
        /// </summary>
        /// <param name="num">номер столбца</param>ъ
        /// <param name="ws">лист в эксель</param>
        /// <returns></returns>
        public string Number2Letter(Excel.Worksheet ws, int num)
        {
            string address = ws.Cells[1, num].Address;
            return address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
        }
    }
}
