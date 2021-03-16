using ACO.ExcelHelpers;
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml.Linq;

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
       
        public static OfferColumnMapping GetCellFromXElement(XElement xElement)
        {
            return new OfferColumnMapping()
            {
                Name = xElement.Attribute("Name").Value,
                ColumnSymbol = xElement.Attribute("ColumnSymbol").Value               
            };
        }

        public XElement GetXElement()
        {
            XElement xeColumn = new XElement("column");
            xeColumn.Add(new XAttribute("Name", Name));
            xeColumn.Add(new XAttribute("ColumnSymbol", ColumnSymbol??""));           
            return xeColumn;
        }
            
    }
}
