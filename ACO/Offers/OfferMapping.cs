using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ACO.Offers
{
    /// <summary>
    ///  Настройки КП. Чтение\ Создание XML
    /// </summary>
    class OfferMapping
    {
        public string Name { get; set; }
        public string FileName { get; set; }
        public OfferMapping() { }
        public OfferMapping(string filename)
        {
             GetFromXML(filename);
        }

        /// <summary>
        /// Ячейки заголовков
        /// </summary>
        public List<ColumnMapping> Columns { get; set; }

        //public List<> Mapping { get; set; }
            

        internal static void Create(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) { return; }
            string path = GetNamesSettingsKP();
            string filename = Path.Combine(path, name + ".xml");
            if (!File.Exists(filename))
            {
                CreateNewProjectXML(name, filename);
            }
            else
            {
                if (MessageBox.Show("Удалить старый файл?", "Файл уже существует!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    File.Delete(filename);
                    CreateNewProjectXML(name, filename);
                }
            }
        }
        private static string GetNamesSettingsKP()
        {
            string path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Spectrum",
            "ACO",
            "Offers"
            );
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
          //  string filename = Path.Combine(path, $"{name}.xml");
            return path;
        }

        /// <summary>
        ///  Создать новый файл 
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="path"></param>
        public static void CreateNewProjectXML(string projectname, string path)
        {            
            XElement root = new XElement("OfferSettings");
            root.Add(new XAttribute("OfferName", projectname));         
            XElement xeColumns = new XElement("Columns");
            root.Add(xeColumns);
            XDocument xdoc = new XDocument(root);
            xdoc.Save(path);
        }
        public void GetFromXML(string filename)
        {
            // OfferMapping mapping = new OfferMapping();
            XDocument xdoc = XDocument.Load(filename);
            XElement root = xdoc.Root;
            FileName = filename;
            // XAttribute xeName = root.Attribute("Name");
            Name = root.Attribute("OfferName").Value?.ToString() ?? "";
            //mapping.Active = bool.Parse(root.Attribute("Active").Value?.ToString() ?? "false");
            Columns = LoadColumnsFromXElement(root.Element("Columns"));
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

        public void Save()
        {
            XElement root = new XElement("OfferSettings");
            XAttribute xaName = new XAttribute("OfferName", Name);
            root.Add(xaName);
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
    }
}
