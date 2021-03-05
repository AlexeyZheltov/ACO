using ACO.ExcelHelpers;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ACO.ProjectManager
{
    class ColumnMapping : Cell
    {
        /// <summary>
        ///  Название ячейки
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        ///  Проверять
        /// </summary>
        public bool Check { get; set; }
        /// <summary>
        /// Обязательный
        /// </summary>
        public bool Obligatory { get; set; }


        public static ColumnMapping GetCellFromXElement(XElement xElement)
        {
            return new ColumnMapping()
            {
                Name = xElement.Attribute("Name").Value,
                Value = xElement.Attribute("Value").Value,
                Row = int.Parse(xElement.Attribute("Row").Value),
                Column = int.Parse(xElement.Attribute("Column").Value),
                Address = xElement.Attribute("Address").Value
            };

        }
        public XElement GetXElement()
        {
            XElement xeColumn = new XElement("Column");
            xeColumn.Add(new XAttribute("Name", Name));
            xeColumn.Add(new XAttribute("Value", Value));
            xeColumn.Add(new XAttribute("Row", Row));
            xeColumn.Add(new XAttribute("Column", Column));
            xeColumn.Add(new XAttribute("Address", Address));
            return xeColumn;
        }

        internal bool CheckSheet(Excel.Worksheet sheet)
        {
            bool VeiwCheck = true;
            if (Check)
            {
                Excel.Range cell = sheet.Cells[Row, Column];
                VeiwCheck = cell.Value == Value;
            }
            return VeiwCheck;
        }
    }
}
