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
    /// <summary>
    ///  Ячейка для сохранения в настройках
    /// </summary>
    class ColumnMapping : Cell
    {
        /// <summary>
        ///  Название ячейки
        /// </summary>
      //  public string Name { get; set; }
        /// <summary>
        ///  Проверять
        /// </summary>
        public bool Check { get; set; }
        /// <summary>
        /// Обязательный
        /// </summary>
        public bool Obligatory { get; set; }

       public ColumnMapping() { }
       public ColumnMapping(Excel.Range cell) 
        {
          //  Name = cell.Value;
            Value = cell.Value;
            Check = false;
            Obligatory = false;
            Address = cell.Address;
            Row = cell.Row;
            Column = cell.Column;
        }
        public static ColumnMapping GetCellFromXElement(XElement xElement)
        {
            return new ColumnMapping()
            {
                //Value = xElement.Attribute("Name").Value,
                Value = xElement.Attribute("Value").Value,
                Row = int.Parse(xElement.Attribute("Row").Value),
                Column = int.Parse(xElement.Attribute("Column").Value),
                Address = xElement.Attribute("Address").Value,
                Check = bool.Parse(xElement.Attribute("Check").Value),
                Obligatory = bool.Parse(xElement.Attribute("Obligatory").Value)
            };
        }

        public XElement GetXElement()
        {
            XElement xeColumn = new XElement("column");
           // xeColumn.Add(new XAttribute("Name", Name));
            xeColumn.Add(new XAttribute("Value", Value));
            xeColumn.Add(new XAttribute("Row", Row));
            xeColumn.Add(new XAttribute("Column", Column));
            xeColumn.Add(new XAttribute("Address", Address));
            xeColumn.Add(new XAttribute("Check", Check.ToString()));
            xeColumn.Add(new XAttribute("Obligatory", Obligatory.ToString()));
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
