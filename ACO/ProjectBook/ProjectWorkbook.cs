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
    public enum StaticColumnsOffer
    {
        StartOffer,
        Level,
        File,
        Number,
        Cipher,
        Classifier,
        Name,
        Code,
        Material,
        Size,
        Type,
        VendorCode,
        Label,
        Producer,
        Unit,
        Amount,
        ContractorAmount,
        CostMaterialsPerUnit,
        CostMaterialsTotal,
        CostWorksPerUnit,
        CostWorksTotal,
        CostTotalPerUnit,
        CostTotal,
        Comment,               
        ColStartOfferComments,
        ColCommentsDescriptionWorks,
        ColDeviationVolume,
        ColCommentsVolumeWorks,
        ColDeviationCost,
        ColCommentsCostWorks,
        ColDeviationMaterials,
        ColDeviationWorks,
        ColComments
    }

    public class ProjectWorkbook
    {


        public static Dictionary<StaticColumnsOffer, string> ColumnsMarksOffer =
            new Dictionary<StaticColumnsOffer, string>
            {
                { StaticColumnsOffer.StartOffer, "offer_start"  },
                { StaticColumnsOffer.Level, "Уровень" },
                { StaticColumnsOffer.File, "Файл" },
                { StaticColumnsOffer.Number, "№ п/п" },
                { StaticColumnsOffer.Cipher, "Шифр" },
                { StaticColumnsOffer.Classifier, "Классификатор" },
                { StaticColumnsOffer.Name, "Наименование работ" },
                { StaticColumnsOffer.Code, "Маркировка / Обозначение" },
                { StaticColumnsOffer.Material, "Материал" },
                { StaticColumnsOffer.Size, "Формат / Габаритные размеры / Диаметр" },
                { StaticColumnsOffer.Type, "Тип, марка, обозначение" },
                { StaticColumnsOffer.VendorCode, "Артикул" },
                { StaticColumnsOffer.Producer, "Производитель" },
                { StaticColumnsOffer.Label, "Маркировка" },
                { StaticColumnsOffer.Unit, "Ед. изм." },
                { StaticColumnsOffer.Amount, "Кол-во" },
                { StaticColumnsOffer.ContractorAmount, "Кол-во (подрядчик)" },
                { StaticColumnsOffer.CostMaterialsPerUnit, "Цена материалы за ед." },
                { StaticColumnsOffer.CostMaterialsTotal, "Цена материалы всего" },
                { StaticColumnsOffer.CostWorksPerUnit, "Цена работы за ед." },
                { StaticColumnsOffer.CostWorksTotal, "Цена работы всего" },
                { StaticColumnsOffer.CostTotalPerUnit, "Итого за ед." },
                { StaticColumnsOffer.CostTotal, "Итого" },
                { StaticColumnsOffer.Comment, "Примечание" },
                { StaticColumnsOffer.ColStartOfferComments, "offer_end" },
                { StaticColumnsOffer.ColCommentsDescriptionWorks,"Комментарии к описанию работ" },
                { StaticColumnsOffer.ColDeviationVolume, "Отклонение по объемам" },
                { StaticColumnsOffer.ColCommentsVolumeWorks,"Комментарии к объемам работ" },
                { StaticColumnsOffer.ColDeviationCost,"Отклонение по стоимости" },
                { StaticColumnsOffer.ColCommentsCostWorks,"Комментарии к стоимости работ" },
                { StaticColumnsOffer.ColDeviationMaterials,"Отклонение МАТ" },
                { StaticColumnsOffer.ColDeviationWorks, "Отклонение РАБ" },
                { StaticColumnsOffer.ColComments,"Комментарии к строкам" }
            };



        readonly Excel.Workbook _ProjectBook = Globals.ThisAddIn.Application.ActiveWorkbook;
        readonly Project _project;
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
        readonly Excel.Worksheet _SheetPallet;

        //public List<OfferAddress> OfferAddress
        //{
        //    get
        //    {
        //        if (_OfferAddress == null)
        //        {
        //            _OfferAddress = GetAddersses();
        //        }
        //        return _OfferAddress;
        //    }
        //    set
        //    {
        //        _OfferAddress = value;
        //    }
        //}
        //List<OfferAddress> _OfferAddress;



        /// <summary>
        ///  Столбцы 
        /// </summary>
        public List<OfferColumns> OfferColumns
        {
            get
            {
                if (_OfferColumns == null)
                {
                    _OfferColumns = GetOfferColumns();
                }
                return _OfferColumns;
            }
            set
            {
                _OfferColumns = value;
            }
        }
        List<OfferColumns> _OfferColumns;
        

        public ProjectWorkbook()
        {
            _project = new ProjectManager.ProjectManager().ActiveProject;
            _SheetPallet = ExcelHelper.GetSheet(_ProjectBook, "Палитра");
        }

        /// <summary>
        ///  Столбцы КП \\ новый
        /// </summary>
        /// <returns></returns>
        public List<OfferColumns> GetOfferColumns()
        {
            List<OfferColumns> columns = new List<OfferColumns>();
            /// Последний столбец в первой строке
            int lastCol = AnalisysSheet.Cells[1, AnalisysSheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

            OfferColumns offerColumns = default;
            for (int col = 1; col <= lastCol; col++)
            {

                string val = _AnalisysSheet.Cells[1, col].Value?.ToString() ?? "";
                if (val == ColumnsMarksOffer[StaticColumnsOffer.StartOffer])
                {
                    // Первый столбец
                    offerColumns = new OfferColumns
                    {
                        ColStartOffer = col
                    };
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColComments] && offerColumns != null)
                {
                    // Последний столбец
                    offerColumns.ColComments = col;
                    columns.Add(offerColumns);

                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.Level] && offerColumns != null)
                {
                    offerColumns.ColLevelOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.Name] && offerColumns != null)
                {
                    offerColumns.ColNameOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.Unit] && offerColumns != null)
                {
                    offerColumns.ColUnitOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.Amount] && offerColumns != null)
                {
                    offerColumns.ColCountOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.CostMaterialsPerUnit] && offerColumns != null)
                {
                    offerColumns.ColCostMaterialsPerUnitOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.CostMaterialsTotal] && offerColumns != null)
                {
                    offerColumns.ColCostMaterialsTotalOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.CostWorksPerUnit] && offerColumns != null)
                {
                    offerColumns.ColCostWorksPerUnitOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.CostWorksTotal] && offerColumns != null)
                {
                    offerColumns.ColCostWorksTotalOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.CostTotalPerUnit] && offerColumns != null)
                {
                    offerColumns.ColTotalCostPerUnitOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.CostTotal] && offerColumns != null)
                {
                    offerColumns.ColCostTotalOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.Comment] && offerColumns != null)
                {
                    offerColumns.ColCommentOffer = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColStartOfferComments] && offerColumns != null)
                {
                    offerColumns.ColStartOfferComments = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColCommentsDescriptionWorks] && offerColumns != null)
                {
                    offerColumns.ColCommentsDescriptionWorks = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColDeviationVolume] && offerColumns != null)
                {
                    offerColumns.ColDeviationVolume = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColCommentsVolumeWorks] && offerColumns != null)
                {
                    offerColumns.ColCommentsVolumeWorks = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColDeviationCost] && offerColumns != null)
                {
                    offerColumns.ColDeviationCost = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColCommentsCostWorks] && offerColumns != null)
                {
                    offerColumns.ColCommentsCostWorks = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColDeviationMaterials] && offerColumns != null)
                {
                    offerColumns.ColDeviationMaterials = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColDeviationWorks] && offerColumns != null)
                {
                    offerColumns.ColDeviationWorks = col;
                }
                else if (val == ColumnsMarksOffer[StaticColumnsOffer.ColComments] && offerColumns != null)
                {
                    offerColumns.ColComments = col;
                }
              
            }
            return columns;
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
            int lastCol = AnalisysSheet.Cells[1, AnalisysSheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column + 8;
            int lastRow = AnalisysSheet.UsedRange.Row + AnalisysSheet.UsedRange.Rows.Count - 1;
            string letterNumber = GetLetter(StaticColumns.Number);
            Excel.Range cell = AnalisysSheet.Cells[lastRow, lastCol];
            Excel.Range rng = AnalisysSheet.Range[$"{letterNumber}{_project.RowStart}:{cell.Address[ColumnAbsolute: false]}"];

            return rng;
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
