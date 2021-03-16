using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ACO.Offers;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Excel;
using ACO.ProjectManager;
using System.Diagnostics;
using ACO.ExcelHelpers;

namespace ACO
{
    /// <summary>
    /// Собирает данные из КП
    /// </summary>
    class OfferManager
    {
        private Excel.Worksheet _sheet;

        public OfferManager() { }
             
        public Offer Offer { get; set; }

        private List<OfferSettings> _Mappings;
        public List<OfferSettings> Mappings
        {
            get
            {
                if (_Mappings == null)
                {
                    _Mappings = GetMappings();
                }
                return _Mappings;
            }
            set { _Mappings = value; }
        }


        public List<OfferSettings> GetMappings()
        {
            List<OfferSettings> mappings = new List<OfferSettings>();
            string folder = GetFolderSettingsKP();
            string[] files = Directory.GetFiles(folder);
            foreach (string file in files)
            {
                mappings.Add(new OfferSettings(file));
            }
            return mappings;
        }
        public static string GetFolderSettingsKP()
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
              
    }
}
