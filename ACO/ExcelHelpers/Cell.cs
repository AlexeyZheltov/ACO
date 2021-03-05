using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ACO.ExcelHelpers
{
    struct Cell
    {
       // public string Name { get; set; }
        /// <summary>
        ///  Номер строки
        /// </summary>
        public string Name { get; set; }
        public string Address { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public string ColumnString { get; set; }
        public string Value { get; set; }
        //public string Addres { get => $"{ ColumnString + Row }"; }

        public static Cell GetCellFromXElement(XElement xElement)
        {
            return new Cell()
            {
                Name = xElement.Attribute("Name").Value,
                Value = xElement.Attribute("Value").Value,
                Row = int.Parse(xElement.Attribute("Row").Value),
                Column = int.Parse(xElement.Attribute("Column").Value),
               // ColumnString = xElement.Attribute("ColumnString").Value,
                Address = xElement.Attribute("Address").Value
            };

        }
        public  XElement GetXElement()
        {         
           XElement xeColumn = new XElement("Column");
            xeColumn.Add(new XAttribute("Name", Name));
            xeColumn.Add(new XAttribute("Value", Value));
            xeColumn.Add(new XAttribute("Row", Row));
            xeColumn.Add(new XAttribute("Column", Column));
            xeColumn.Add(new XAttribute("Address", Address));
            return xeColumn;
        }

    }
}
