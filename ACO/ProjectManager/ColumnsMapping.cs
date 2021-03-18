﻿using ACO.ExcelHelpers;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace ACO.ProjectManager
{
    /// <summary>
    ///  Ячейка для сохранения в настройках
    /// </summary>
    class ColumnMapping 
    {
        /// <summary>
        ///  Проверять
        /// </summary>
        public bool Check { get; set; }
        /// <summary>
        /// Обязательный
        /// </summary>
        public bool Obligatory { get; set; }

        /// <summary>
        ///  Наименование столбца на листе анализ
        /// </summary>
        public string Name { get; set; }

        public string ColumnSymbol { get; set; }

        public int Column { get; set; }

        public ColumnMapping() { }
       public ColumnMapping(Excel.Range cell) 
        {
            //Value = cell.Value?.ToString()??"";
            Check = false;
            Obligatory = false;
            ColumnSymbol = Regex.Match(cell.Address, @"[A-Z]*")?.Value??"";
           // Address = cell.Address;
            //Row = cell.Row;
            //Column = cell.Column;
        }
        public static ColumnMapping GetCellFromXElement(XElement xElement)
        {
            return new ColumnMapping()
            {
                Name = xElement.Attribute("Name").Value,
               // Value = xElement.Attribute("Value").Value,
               // Row = int.Parse(xElement.Attribute("Row").Value),
              //  Column = int.Parse(xElement.Attribute("Column").Value),
                ColumnSymbol= xElement.Attribute("ColumnSymbol").Value?.ToString()??"",
                Check = bool.Parse(xElement.Attribute("Check").Value),
                Obligatory = bool.Parse(xElement.Attribute("Obligatory").Value)
            };
        }

        public XElement GetXElement()
        {
            XElement xeColumn = new XElement("column");
            xeColumn.Add(new XAttribute("Name", Name));
            //xeColumn.Add(new XAttribute("Value", Value));
            //xeColumn.Add(new XAttribute("Row", Row));
            //xeColumn.Add(new XAttribute("Column", Column));
            //xeColumn.Add(new XAttribute("Address", Address));
            xeColumn.Add(new XAttribute("Check", Check.ToString()));
            xeColumn.Add(new XAttribute("Obligatory", Obligatory.ToString()));
            xeColumn.Add(new XAttribute("ColumnSymbol", ColumnSymbol));
            return xeColumn;
        }

        internal bool CheckSheet(Excel.Worksheet sheet)
        {
            bool VeiwCheck = true;
            //if (Check)
            //{
            //    Excel.Range cell = sheet.Cells[Row, Column];
            //    VeiwCheck = cell.Value == Value;
            //}
            return VeiwCheck;
        }
    }
}