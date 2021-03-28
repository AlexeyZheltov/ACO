using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACO.ProjectManager;
using ACO.ExcelHelpers;
using ACO.ProjectBook;
using System;

namespace ACO
{
    public class ProjectWorkbook
    {
        Excel.Workbook _ProjectBook = Globals.ThisAddIn.Application.ActiveWorkbook;
        Project _project;
        public Excel.Worksheet AnalisysSheet
        {
            get
            {
                if (_AnalisysSheet is null)
                {
                    _AnalisysSheet = ExcelHelper.GetSheet(_ProjectBook, _project.AnalysisSheetName);
                }
                return _AnalisysSheet;
            }
            set
            {
                _AnalisysSheet = value;
            }
        }
        Excel.Worksheet _AnalisysSheet;

        public List<OfferAddress> OfferAddress
        {
            get
            {
                if (_OfferAddress == null)
                {
                    _OfferAddress = GetAddersses();
                }
                return _OfferAddress;
            }
            set
            {
                _OfferAddress = value;
            }
        }
        List<OfferAddress> _OfferAddress;
              
        public ProjectWorkbook()
        {
            _project = new ProjectManager.ProjectManager().ActiveProject;
        }

        /// <summary>
        ///  Номера столбцов 
        /// </summary>
        /// <returns></returns>
        public List<OfferAddress> GetAddersses()
        {
            List<OfferAddress> addresses = new List<OfferAddress>();
            int lastCol = AnalisysSheet.Cells[1, AnalisysSheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            int columnStart = 0;
            int columnTotal = 0;

            for (int col = 1; col <= lastCol; col++)
            {
                string val = _AnalisysSheet.Cells[1, col].Value?.ToString() ?? "";
                if (val == "offer_start") { columnStart = col; }
                else if (val == Project.ColumnsNames[StaticColumns.CostTotal]) { columnTotal = col; }
                else if (val == "offer_end")
                {
                    OfferAddress address = new OfferAddress
                    {
                        ColStartOffer = columnStart,
                        ColStartOfferComments = col,
                        ColTotalCost = columnTotal,
                        ColPercentTotal = col + 4,
                        ColPercentMaterial = col + 6,
                        ColPercentWorks = col + 7,
                        ColComments = col + 8
                    };
                    addresses.Add(address);
                }
            }
            return addresses;
        }

        public int GetFirstRow()
        {
            return _project.RowStart;
        }
    }
}
