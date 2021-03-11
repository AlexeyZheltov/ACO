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

        public OfferManager(ExcelFile excelBook)
        {
            ExcelBook = excelBook;
            //  Excel.Worksheet sheet = excelBook.GetSheet(Offer.SheetName);
            _sheet = excelBook.GetSheet(Offer.SheetName);
        }
        public ExcelFile ExcelBook { get; }

        public Offer Offer { get; set; }

        private List<OfferMapping> _Mappings;
        public List<OfferMapping> Mappings
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


        public List<OfferMapping> GetMappings()
        {
            List<OfferMapping> mappings = new List<OfferMapping>();
            string folder = GetFolderSettingsKP();
            string[] files = Directory.GetFiles(folder);
            foreach (string file in files)
            {
                mappings.Add(new OfferMapping(file));
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


        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private Offer SetOffer(Excel.Worksheet sheet)
        {
            Offer offer = new Offer();
            offer.Date = sheet.Cells[1, 1].Value?.ToSting() ?? "";
            offer.Customer = sheet.Cells[1, 1].Value?.ToSting() ?? "";
            offer.ProjectName = sheet.Cells[1, 1].Value?.ToSting() ?? "";
            offer.ProjectNumber = sheet.Cells[1, 1].Value?.ToSting() ?? "";
            return offer;
        }

        /// <summary>
        ///  считать КП
        /// </summary>
        /// <returns></returns>
        public bool ReadOffer()
        {
            OfferMapping mapping = FindColumnsMapping();
            bool validation = mapping != null;
            if (validation)
            {
                int rowStart = GetRowStart(_sheet);
                int rowEnd = _sheet.UsedRange.Row + _sheet.UsedRange.Rows.Count - 1;
                Offer = new Offer();
                //int columnNumber = mapping.Columns.Find(c => c.Value == "П.П.").Column;

                for (int row = rowStart; row <= rowEnd; row++)
                {
                    try
                    {
                        Record record = new Record();
                        foreach (ColumnMapping col in mapping.Columns)
                        {
                            if (col.Value == "П.П.")
                            {
                                record.Number = _sheet.Cells[row, col.Column].Value ?? "";
                            }
                            object val = _sheet.Cells[row, col.Column].Value;
                            string key = col.Value; // Заголовок
                            if (!record.Values.ContainsKey(key))
                            { //TODO Поправить ключ для составной шапки
                                record.Values.Add(key, val);
                            }
                        }
                        /// Сохранение  строки 
                        Offer.Records.Add(record);
                    }
                    catch (AddInException ex)
                    {
                        validation = ex.StopProcess;
                        if (ex.StopProcess) break;
                    }
                }
            }
            return validation;
        }

        /// <summary>
        /// Выбрать маппинг. Проверить столбцы КП на листе. 
        /// </summary>
        /// <returns></returns>
        private OfferMapping FindColumnsMapping()
        {
            OfferMapping checkedMapping = null;
            foreach (OfferMapping mapping in Mappings)
            {
                foreach (ColumnMapping col in mapping.Columns)
                {
                    try
                    {
                        string val = _sheet.Range[col.Address].Value?.ToString() ?? "";
                        if (val != col.Value)
                        {
                            throw new ApplicationException("Значение в ячейке не соответствует файлу");
                        }
                    }
                    catch (AddInException ex)
                    {
                        Debug.WriteLine(ex.Message);
                        // При возникновении ошибки выбрать другой файл маппинга
                        continue;
                    }
                }
                checkedMapping = mapping;
            }
            return checkedMapping;
        }

        private int GetRowStart(Excel.Worksheet sheet)
        {
            Excel.Range findcell = sheet.UsedRange.Find("НАИМЕНОВАНИЕ РАБОТ", LookIn: Excel.XlFindLookIn.xlValues);
            if (findcell is null) throw new AddInException("Лист не соответствует формату");
            int row = findcell.Row + 2;
            return row;
        }
    }
}
