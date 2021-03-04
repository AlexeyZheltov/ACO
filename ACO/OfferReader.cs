using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO
{
    /// <summary>
    /// Собирает данные и 
    /// </summary>
    class OfferReader
    {
        Excel.Worksheet _sheet;
        public Offer Offer
        {
            get
            {
                if (_Offer is null)
                {
                    _Offer = SetOffer(_sheet);
                }
                return _Offer;
            }
            private set
            {
                _Offer = value;
            }
        }
        private Offer _Offer;

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


        public OfferReader(Excel.Worksheet sheet)
        {
            _sheet = sheet;
        }
        public bool ReadOffer()
        {
            bool validation = true;
            int rowStart = GetRowStart(_sheet);
            int rowEnd = _sheet.UsedRange.Row + _sheet.UsedRange.Rows.Count - 1;
            for (int row = rowStart; row <= rowEnd; row++)
            {
                try
                {
                    Item rowItem = new Item();
                    Offer.Items.Add(rowItem);
                }
                catch (AddInException ex)
                {
                    validation = ex.StopProcess;
                    if (ex.StopProcess) break;
                }
            }
            return validation;
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
