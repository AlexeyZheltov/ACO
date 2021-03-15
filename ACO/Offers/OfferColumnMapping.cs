using ACO.ExcelHelpers;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;

namespace ACO.Offers
{
    class OfferColumnMapping
    {
        /// <summary>
        ///  Наименование столбца на листе анализ
        /// </summary>
        public string Name { get; set; }

        public string ColumnSymbol { get; set; }


        public OfferColumnMapping() { }
        //public OfferColumnMapping(Excel.Range cell)
        //{
        //    Link = "";
        //    Value = cell.Value?.ToString() ?? "";           
        //    Address = cell.Address;
        //    Row = cell.Row;
        //    Column = cell.Column;
        //}
        public static OfferColumnMapping GetCellFromXElement(XElement xElement)
        {
            return new OfferColumnMapping()
            {
                Name = xElement.Attribute("Name").Value,
                ColumnSymbol = xElement.Attribute("ColumnSymbol").Value
                //Link = xElement.Attribute("Link").Value,
                //Value = xElement.Attribute("Value").Value,
                //Row = int.Parse(xElement.Attribute("Row").Value),
                //Column = int.Parse(xElement.Attribute("Column").Value),
                //Address = xElement.Attribute("Address").Value
            };
        }

        public XElement GetXElement()
        {
            XElement xeColumn = new XElement("column");
            xeColumn.Add(new XAttribute("Name", Name));
            xeColumn.Add(new XAttribute("ColumnSymbol", ColumnSymbol));
            //xeColumn.Add(new XAttribute("Link", Link));
            //xeColumn.Add(new XAttribute("Value", Value));
            //xeColumn.Add(new XAttribute("Row", Row));
            //xeColumn.Add(new XAttribute("Column", Column));
            //xeColumn.Add(new XAttribute("Address", Address));
            return xeColumn;
        }

        //internal bool CheckSheet(Excel.Worksheet sheet)
        //{
        //    bool VeiwCheck = true;          
        //        Excel.Range cell = sheet.Cells[Row, Column];
        //        VeiwCheck = cell.Value == Value;            
        //    return VeiwCheck;
        //}

    }
}
