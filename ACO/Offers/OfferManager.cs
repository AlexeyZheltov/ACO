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

        public OfferManager() { }

        /// <summary>
        ///  Адреса столбцов
        /// </summary>
        public List<OfferSettings> Mappings
        {
            get
            {
                if (_Mappings == null)
                {
                    UpdateMappings();
                }
                return _Mappings;
            }
            set { _Mappings = value; }
        }
        private List<OfferSettings> _Mappings;


        public void UpdateMappings()
        {
            List<OfferSettings> mappings = new List<OfferSettings>();
            string folder = GetFolderSettingsKP();
            string[] files = Directory.GetFiles(folder);
            foreach (string file in files)
            {
                mappings.Add(new OfferSettings(file));
            }


            string filename = GetBaseEstimateFilename();
            if (mappings.Find(x => x.FileName == filename) == null)
            {
                mappings.Add(CreateBaseEstimateSettings(filename));
            }

            _Mappings = mappings;
        }

        private static string GetBaseEstimateFilename() //  SpectrumFilename()
        {
            string folder = GetFolderSettingsKP();
            string filename = Path.Combine(folder, "Базовая оценка.xml");
            return filename;
        }

        private static OfferSettings CreateBaseEstimateSettings(string filename)
        {
            List<OfferColumnMapping> columns = new List<OfferColumnMapping>
            {
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Name],
                    ColumnSymbol = "M"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Level],
                    ColumnSymbol = "I"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Cipher],
                    ColumnSymbol = "L"
                },
                  new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Label],
                    ColumnSymbol = ""
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Number],
                    ColumnSymbol = "K"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Unit],
                    ColumnSymbol = "T"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Amount],
                    ColumnSymbol = "U"
                },
                  new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.ContractorAmount],
                    ColumnSymbol = ""
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostMaterialsPerUnit],
                    ColumnSymbol = "V"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostMaterialsTotal],
                    ColumnSymbol = "W"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostWorksPerUnit],
                    ColumnSymbol = "X"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostWorksTotal],
                    ColumnSymbol = "Y"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostTotalPerUnit],
                    ColumnSymbol = "Z"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostTotal],
                    ColumnSymbol = "AA"
                },
                new OfferColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Comment],
                    ColumnSymbol = "AB"
                }
            };

            OfferSettings BaseEstimateSettings = new OfferSettings()
            {
                Name = "Базовая оценка",
                FileName = filename,
                RowStart = 23,
                SheetName = "Рсч-П",
                Columns = columns
            };
            BaseEstimateSettings.Save();
            return BaseEstimateSettings;
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
