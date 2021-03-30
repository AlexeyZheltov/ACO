using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACO.ProjectManager;
using ACO.ExcelHelpers;
using ACO.ProjectBook;
using System;
using System.Drawing;

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
        Excel.Worksheet _SheetPallet;

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
            _SheetPallet = ExcelHelper.GetSheet(_ProjectBook, "Палитра");
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
        public string GetLetter(StaticColumns column)
        {
           ColumnMapping mapping = _project.Columns.Find(x => x.Name == Project.ColumnsNames[column]);
            if (mapping is null) throw new AddInException($"Не в проекте не указан столбец: {Project.ColumnsNames[column]}");
            return mapping.ColumnSymbol;
        }

        public Excel.Range GetAnalysisRange()
        {
            int lastCol = AnalisysSheet.Cells[1, AnalisysSheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column +8;
            int lastRow = AnalisysSheet.UsedRange.Row + AnalisysSheet.UsedRange.Rows.Count - 1;
            string letterNumber =  GetLetter(StaticColumns.Number);
            //string letter = ExcelHelper.
            Excel.Range cell = AnalisysSheet.Cells[lastRow,lastCol];
            Excel.Range rng = AnalisysSheet.Range[$"{letterNumber}{_project.RowStart}:{cell.Address[ColumnAbsolute: false]}"];
            //Excel.Range rng = AnalisysSheet.Range[AnalisysSheet.Cells[_project.RowStart,2], AnalisysSheet.Cells[lastRow, lastCol]];
            return rng;
        }

        public void ColorCell(Excel.Range cell, string lvl = "defalut")
        {
            string text = cell.Value?.ToString() ?? "";
            if (text != "#НД" || text != "")
            {
                double percent = double.TryParse(text, out double pct) ? pct : 0;
                if (percent > 0.15 || text.Contains("Отс-ет"))
                {//Красный  >0.15
                    cell.Interior.Color = Color.FromArgb(255, 0, 0);
                    cell.Font.Color = Color.FromArgb(255, 255, 255);
                }
                else if (percent < -0.15)
                {// Желтый 
                    cell.Interior.Color = Color.FromArgb(242, 255, 0);
                    cell.Font.Color = Color.FromArgb(242, 0, 0);
                }
                else if (percent > 0.05 && percent < 0.15)
                {
                    /// Светло фиолетовый
                    cell.Interior.Color = Color.FromArgb(255, 176, 197);
                    cell.Font.Color = Color.FromArgb(125, 0, 33);
                }
                else if (percent < -0.05 && percent > -0.15)
                {// Светло желтый
                    cell.Interior.Color = Color.FromArgb(252, 250, 104);
                    cell.Font.Color = Color.FromArgb(0, 0, 0);
                }
                else if (lvl != "defalut")
                {
                    // Формат строки по уровню
                    Dictionary<string, Excel.Range> pallets = ExcelReader.ReadPallet(_SheetPallet);
                    if (pallets.TryGetValue(lvl, out Excel.Range pallet))
                    {
                        pallet.Copy();
                        cell.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    }
                }
                else
                {
                    cell.Interior.Color = Color.FromArgb(232, 242, 238);
                    cell.Font.Color = Color.FromArgb(0, 0, 0);
                }
            }
        }


    }
}
