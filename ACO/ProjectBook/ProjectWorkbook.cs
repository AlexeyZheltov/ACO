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
                    Excel.Range cellName = _AnalisysSheet.Cells[6, columnStart + 1];
                    string name = cellName.Value?.ToString() ?? "";
                    if (string.IsNullOrEmpty(name))
                    {
                        name = $"УЧАСТНИК {addresses.Count + 1}";
                        cellName.Value = name;
                    }

                    OfferAddress address = new OfferAddress
                    {
                        Name = name,
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
                        ExcelHelper.SetCellFormat(cell, pallet);
                        //pallet.Copy();
                        //cell.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                    }
                }
                else
                {
                    cell.Interior.Color = Color.FromArgb(232, 242, 238);
                    cell.Font.Color = Color.FromArgb(0, 0, 0);
                }
            }
        }

        /// <summary>
        ///  Определить столбцы для окрашивания
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        public static List<(string, string)> GetColredColumns(Excel.Worksheet ws)
        {
            List<(string, string)> columns = new List<(string, string)>();
            int lastCol = ws.Cells[1, ws.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            Excel.Range cellStart = null;
            Excel.Range cellEnd = null;

            for (int col = 1; col <= lastCol; col++)
            {
                Excel.Range cell = ws.Cells[1, col];
                string val = cell.Value?.ToString() ?? "";

                if (val == "offer_start")
                {
                    cellStart = cell.Offset[0, 1];
                }
                if (val == "offer_end")
                {
                    cellEnd = cell.Offset[0, -1];
                }
                if (cellStart != null && cellEnd != null && cellStart.Column < cellEnd.Column)
                {
                    string addressStart = cellStart.Address;
                    string letterStart = addressStart.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
                    string addressEnd = cellEnd.Address;
                    string letterEnd = addressEnd.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
                    if (!string.IsNullOrEmpty(letterStart) && !string.IsNullOrEmpty(letterEnd))
                    {
                        columns.Add((letterStart, letterEnd));
                    }
                    cellStart = null;
                    cellEnd = null;
                }
            }
            return columns;
        }

        /// <summary>
        ///  
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        public static List<(string, string)> GetFormatColumns(Excel.Worksheet ws)
        {
            List<(string, string)> columns = new List<(string, string)>();
            int lastCol = ws.Cells[1, ws.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            Excel.Range cellStart = null;
            Excel.Range cellEnd = null;

            for (int col = 1; col <= lastCol; col++)
            {
                Excel.Range cell = ws.Cells[1, col];
                string val = cell.Value?.ToString() ?? "";

                if (val == Project.ColumnsNames[StaticColumns.Amount])
                {
                    cellStart = cell;
                }
                if (val == Project.ColumnsNames[StaticColumns.CostTotal])
                {
                    cellEnd = cell;
                }
                if (cellStart != null && cellEnd != null && cellStart.Column < cellEnd.Column)
                {
                    string addressStart = cellStart.Address;
                    string letterStart = addressStart.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
                    string addressEnd = cellEnd.Address;
                    string letterEnd = addressEnd.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];
                    if (!string.IsNullOrEmpty(letterStart) && !string.IsNullOrEmpty(letterEnd))
                    {
                        columns.Add((letterStart, letterEnd));
                    }
                    cellStart = null;
                    cellEnd = null;
                }
            }
            return columns;
        }
    }
}
