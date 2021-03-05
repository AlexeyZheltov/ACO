using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO.ExcelHelpers
{
    struct Cell
    {
       // public string Name { get; set; }
        /// <summary>
        ///  Номер строки
        /// </summary>
        public int Row { get; set; }
        public int Column { get; set; }
        public string ColumnString { get; set; }
        public string Value { get; set; }
        //public string Addres { get => $"{ ColumnString + Row }"; }
        public string Address { get; set; }
    }
}
