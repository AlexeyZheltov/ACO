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
  public  class OfferSettings
    {
        public string Name { get; set; }
        public string FileName { get; set; }

        public string SheetName { get; set; }
      
        /// <summary>
        /// Строка начала данных
        /// </summary>
        public int RowStart { get; set; }

        public OfferSettings() { }
        public OfferSettings(string filename)
        {
            GetFromXML(filename);
        }

        /// <summary>
        /// Ячейки заголовков
        /// </summary>
        public List<OfferColumnMapping> Columns { get; set; }       


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
            return path;
        }

        /// <summary>
        ///  Создать новый файл 
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="path"></param>
        public static void CreateNewProjectXML(string projectname, string path)
        {
            OfferSettings offerMapping = new OfferSettings
            {
                Name = projectname,
                FileName = path
            };
            offerMapping.Save();
        }
        public void GetFromXML(string filename)
        {
            XDocument xdoc = XDocument.Load(filename);
            XElement root = xdoc.Root;
            FileName = filename;
            Name = root.Attribute("OfferName").Value?.ToString() ?? "";
            XElement xeSheets = root.Element("Sheets");
            XElement xeDataSheet = xeSheets.Element("DataSheet");
            SheetName = xeDataSheet.Attribute("SheetName").Value?.ToString()??"";
            XElement xeRows = xeDataSheet.Element("Rows");
            XElement xeRowStart = xeRows.Element("RowStart");
            RowStart = int.TryParse(xeRowStart.Attribute("Row").Value?.ToString() ?? "", out int rs) ? rs : 0;           
            Columns = LoadColumnsFromXElement(xeDataSheet.Element("Columns"));
        }

        /// <summary>
        ///  Получить столбцы из XML.
        /// </summary>
        /// <param name="xElement"></param>
        /// <returns></returns>
        private static List<OfferColumnMapping> LoadColumnsFromXElement(XElement xElement)
        {
            List<OfferColumnMapping> columns = new List<OfferColumnMapping>();
            if (xElement != null)
            {
                foreach (XElement xcol in xElement.Elements())
                {
                    columns.Add(OfferColumnMapping.GetCellFromXElement(xcol));
                }
            }
            return columns;
        }

        public void Save()
        {
            XElement root = new XElement("OfferSettings");
            XAttribute xaName = new XAttribute("OfferName", Name ?? "");
            root.Add(xaName);

            XElement xeSheets = new XElement("Sheets");
            XElement xeDataSheet = new XElement("DataSheet");
            xeDataSheet.Add(new XAttribute("SheetName", SheetName ?? ""));
            XElement xeRangeValues = new XElement("RangeVaues");
            //XAttribute xaStart = new XAttribute("Start", RangeValuesStart.ToString());
            //XAttribute xaEnd = new XAttribute("End", RangeValuesEnd.ToString());
            //xeRangeValues.Add(xaStart);
            //xeRangeValues.Add(xaEnd);
            xeDataSheet.Add(xeRangeValues);

            XElement xeRows = new XElement("Rows");
            XElement xeRowStart = new XElement("RowStart");
            xeRowStart.Add(new XAttribute("Row", RowStart.ToString()));
            xeRows.Add(xeRowStart);
            xeDataSheet.Add(xeRows);

            XElement xeColumns = new XElement("Columns");
            if ((Columns?.Count ?? 0) > 0)
            {
                foreach (OfferColumnMapping cell in Columns)
                {
                    XElement xeColumn = cell.GetXElement();
                    xeColumns.Add(xeColumn);
                }
            }
            xeDataSheet.Add(xeColumns);
            xeSheets.Add(xeDataSheet);
            root.Add(xeSheets);

            XDocument xdoc = new XDocument(root);
            xdoc.Save(FileName);
        }
    }
}
