using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ACO.ExcelHelpers
{
    /// <summary>
    ///  Ячейка Excel. 
    /// </summary>
    class Cell
    {     
        /// <summary>
        ///  Значение столбца
        /// </summary>
        public string Value { get; set; }
        /// <summary>
        /// Адрес ячейки
        /// </summary>
        public string Address { get; set; }
        /// <summary>
        ///  Номер строки
        /// </summary>
        public int Row { get; set; }
        /// <summary>
        ///  Номер столбца
        /// </summary>
        public int Column { get; set; }
    }
}
