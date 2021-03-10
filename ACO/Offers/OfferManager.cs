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

namespace ACO
{
    /// <summary>
    /// Собирает данные из КП
    /// </summary>
    class OfferManager
    {
        private Excel.Worksheet _sheet;

        public OfferManager() { }
        public OfferManager(Excel.Worksheet sheet)
        {
            _sheet = sheet;
        }

        private List<OfferMapping> _OffersMapping;
        public List<OfferMapping> OffersMapping
        {
            get 
            {                 
                 _OffersMapping = GetOffers();
                
                return _OffersMapping; 
            }
            set { _OffersMapping = value; }
        }

       

        private List<OfferMapping> GetOffers()
        {
          List<OfferMapping> offers  = new List<OfferMapping>();
            string folder = GetFolderSettingsKP();
            string[] files = Directory.GetFiles(folder); 
            foreach(string file in files)
            {
                offers.Add(new OfferMapping(file));
            }
            return offers;
        }
        private static string GetFolderSettingsKP()
        {
            string path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Spectrum",
            "ACO",
            "Offers"
            );
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            //string filename = Path.Combine(path, "settings.xml");
            return path;
        }

    

        //public Offer Offer
        //{
        //    get
        //    {
        //        if (_Offer is null)
        //        {
        //            _Offer = SetOffer(_sheet);
        //        }
        //        return _Offer;
        //    }
        //    private set
        //    {
        //        _Offer = value;
        //    }
        //}
        //private Offer _Offer;

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

        
        public bool ReadOffer()
        {
            bool validation = CheckColumns();
            int rowStart = GetRowStart(_sheet);
            int rowEnd = _sheet.UsedRange.Row + _sheet.UsedRange.Rows.Count - 1;
            for (int row = rowStart; row <= rowEnd; row++)
            {
                try
                {
                    Item rowItem = new Item();
                    /// Сохранение  строки 
                    //rowItem.
                    //Offer.Items.Add(rowItem);
                }
                catch (AddInException ex)
                {
                    validation = ex.StopProcess;
                    if (ex.StopProcess) break;
                }
            }
            return validation;
        }

        /// <summary>
        ///  проверить столбцы КП
        /// </summary>
        /// <returns></returns>
        private bool CheckColumns()
        {
            return false;
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
