using ACO.ExcelHelpers;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;
using System.Text.RegularExpressions;

namespace ACO.Offers
{
    public class OfferColumnMapping
    {
        /// <summary>
        ///  Наименование столбца на листе анализ
        /// </summary>
        public string Name { get; set; }

        public string ColumnSymbol { get; set; }


        public OfferColumnMapping() { }
       
        private static string ClearColumnSymbol(string symbol)
        {
            return Regex.Replace(symbol, "[^A-Za-z]", "");
        }

        public static OfferColumnMapping GetCellFromXElement(XElement xElement)
        {
            string symbol = xElement.Attribute("ColumnSymbol").Value;
            symbol = ClearColumnSymbol(symbol);

            return new OfferColumnMapping()
            {
                Name = xElement.Attribute("Name").Value,
                ColumnSymbol = symbol
            };
        }

        public XElement GetXElement()
        {
            string symbol = ClearColumnSymbol(ColumnSymbol ?? "");
            XElement xeColumn = new XElement("column");
            xeColumn.Add(new XAttribute("Name", Name));
            xeColumn.Add(new XAttribute("ColumnSymbol", symbol));           
            return xeColumn;
        }
            
    }
}
