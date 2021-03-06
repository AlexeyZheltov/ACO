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
        public List<ColumnMapping> Columns { get; set; }

        public Project(){}

        public void Save()
        {
            XElement root = new XElement("project");
            XAttribute xaName = new XAttribute("ProjectName", Name);
            XAttribute xaActive = new XAttribute("Active", true);
            root.Add(xaName);
            root.Add(xaActive);
            XElement xeColumns = new XElement("Columns");           
          
            foreach (ColumnMapping cell in Columns)
            {
                XElement xeColumn = cell.GetXElement();
                xeColumns.Add(xeColumn);
            }
            root.Add(xeColumns);
            XDocument xdoc = new XDocument(root);
            xdoc.Save(FileName);
        }

        public static Project GetFromXML(string filename)
        {
            Project project = new Project();
            XDocument xdoc = XDocument.Load(filename);
            XElement root = xdoc.Root;
            project.FileName = filename;
            XAttribute xeName = root.Attribute("Name");
            project.Name = root.Attribute("ProjectName").Value?.ToString() ?? "";
            project.Active = bool.Parse(root.Attribute("Active").Value?.ToString() ?? "false");
            project.Columns = LoadColumnsFromXElement(root.Element("Columns"));
            return project;
        }
        private static List<ColumnMapping> LoadColumnsFromXElement(XElement xElement)
        {
            List<ColumnMapping> columns = new List<ColumnMapping>();
            if (xElement != null)
            {
                foreach (XElement xcol in xElement.Elements())
                {
                    columns.Add(ColumnMapping.GetCellFromXElement(xcol));
                }
            }
            return columns;
        }

        public void CreateXML()
        {

        }
      
    }
}
