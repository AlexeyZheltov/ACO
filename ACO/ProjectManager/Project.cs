using ACO.ExcelHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ACO.ProjectManager
{
    public enum StaticColumns 
    { 
        Number,
        Name
    }
    class Project
    {
        /// <summary>
        ///  является ли проект активным // Используется в DataGridView
        /// </summary>
        public bool Active { get; set; }

        /// <summary>
        ///  Название проекта
        /// </summary>
        public string Name { get; set; }


        /// <summary>
        ///  Путь к файлу
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        ///  Название листа 
        /// </summary>
        public string AnalysisSheetName { get; set; }

        public int RangeValuesStart { get; set; }
        public int RangeValuesEnd { get; set; }
        public int RowStart { get; set; }

        /// <summary>
        ///  Адреса ячеек шапки используемых столбцов
        /// </summary>
        public List<ColumnMapping> Columns { get; set; }

        public static Dictionary<StaticColumns, string> staticColumns =
            new Dictionary<StaticColumns, string>
            {
                { StaticColumns.Number, "№ п/п" },                
                { StaticColumns.Name, "Наименование работ и затрат" } 
            };
            


        public Project() { }

        public void Save()
        {
            XElement root = new XElement("project");
            XAttribute xaName = new XAttribute("ProjectName", Name);           
            root.Add(xaName);
           

            XElement xeSheets = new XElement("Sheets");
            XElement xeAnalysisSheet = new XElement("AnalysisSheet");
            xeAnalysisSheet.Add(new XAttribute("Name", AnalysisSheetName));
            XElement xeRows = new XElement("Rows");
            XElement xeRowStart = new XElement("RowStart");
            xeRowStart.Add(new XAttribute("Row", RowStart.ToString()));
            xeRows.Add(xeRowStart);
            xeAnalysisSheet.Add(xeRows);

            XElement xeColumns = new XElement("Columns");
            // Диапазон значения
            XElement xeRangeValues = new XElement("RangeValues");
            xeRangeValues.Add(new XAttribute("StartColumn", RangeValuesStart.ToString()));
            xeRangeValues.Add(new XAttribute("EndColumn", RangeValuesEnd.ToString()));
            xeAnalysisSheet.Add(xeRangeValues);

            foreach (ColumnMapping cell in Columns)
            {
                XElement xeColumn = cell.GetXElement();
                xeColumns.Add(xeColumn);
            }

            xeAnalysisSheet.Add(xeColumns);
            xeSheets.Add(xeAnalysisSheet);
            root.Add(xeSheets);
            XDocument xdoc = new XDocument(root);
            xdoc.Save(FileName);
        }

        public static Project GetFromXML(string filename)
        {
            Project project = new Project();
            XDocument xdoc = XDocument.Load(filename);
            XElement root = xdoc.Root;
            project.FileName = filename;           
            project.Name = root.Attribute("ProjectName").Value?.ToString() ?? "";
            XElement xeSheets = root.Element("Sheets");
            XElement xeAnalysisSheet = xeSheets.Element("AnalysisSheet");

            project.AnalysisSheetName = xeAnalysisSheet.Attribute("Name").Value?.ToString() ?? "";

            XElement xeRows = xeAnalysisSheet.Element("Rows");
            XElement xeRowStart = xeRows.Element("RowStart");
            project.RowStart = int.TryParse(xeRowStart.Attribute("Row").Value, out int r) ? r : 0;
            XElement xeRangeValues = xeAnalysisSheet.Element("RangeValues");
            project.RangeValuesStart = int.TryParse(xeRangeValues.Attribute("StartColumn").Value, out int sc) ? sc : 0; 
            project.RangeValuesEnd = int.TryParse(xeRangeValues.Attribute("EndColumn").Value, out int ec) ? ec : 0; 
            project.Columns = LoadColumnsFromXElement(xeAnalysisSheet.Element("Columns"));

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
    }
}
