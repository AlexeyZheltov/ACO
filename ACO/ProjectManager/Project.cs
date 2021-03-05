using ACO.ExcelHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ACO.ProjectManager
{
    class Project
    {

        public bool Active { get; set; }
        public string Name { get; set; }
        public string FileName { get; set; }
        public List<Cell> Columns { get; set; }

        //public ColumnsMapping MyProperty { get; set; }
        //public class SettingsProject
        public Project()
        {

        }

        public void Save()
        {
            XElement root = new XElement("project");
            XAttribute xaName = new XAttribute("ProjectName", Name);
            XAttribute xaActive = new XAttribute("Active", true);
            root.Add(xaName);
            root.Add(xaActive);
            XElement xeColumns = new XElement("Columns");           
          
            foreach (Cell cell in Columns)
            {
                XElement xeColumn = cell.GetXElement();
                //XElement xeColumn = new XElement("Column");
                //xeColumn.Add( new XAttribute("Name", cell.Name));
                //xeColumn.Add( new XAttribute("Value", cell.Value));
                //xeColumn.Add( new XAttribute("Row", cell.Row));
                //xeColumn.Add( new XAttribute("Column", cell.Column));
                //xeColumn.Add( new XAttribute("Address", cell.Address));
                xeColumns.Add(xeColumn);
            }
            root.Add(xeColumns);
            XDocument xdoc = new XDocument(root);
            xdoc.Save(FileName);
        }
        public void CreateXML()
        {

        }
      
    }
}
