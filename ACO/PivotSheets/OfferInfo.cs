﻿using ACO.ExcelHelpers;
using ACO.ProjectBook;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.PivotSheets
{
    class OfferInfo
    {
        /*
            Описание менялось				                            ПОЗИЦИЯМ
            Объемы завышены				                                ПОЗИЦИЯМ
            Объемы занижены				                                ПОЗИЦИЯМ
            Сумма завышеных работ по разделам				            РУБ БЕЗ НДС
            Разница в стоимости с оценкой СПЕКТРУМ				        РУБ БЕЗ НДС
            НЕ оценено на сумму				                            РУБ БЕЗ НДС
            Выявленные ошибки				                            РУБ БЕЗ НДС
            Итого включая не оцененные работы и корректировку ошибок	РУБ БЕЗ НДС
         */

        readonly Excel.Application _app = Globals.ThisAddIn.Application;
        readonly Excel.Worksheet _SheetUrv12;
        readonly Excel.Worksheet _AnalisysSheet;
        readonly ProjectWorkbook _projectWorkbook;
        readonly ProjectManager.ProjectManager _projectManager;
        readonly ProjectManager.Project _project;

        //Описание менялось
        string _rangeChengedNames;
        string _rangeCostComments;
      //  string _rangeCostTotal;

        public OfferInfo(ProjectWorkbook projectWorkbook)
        {
            Excel.Workbook wb = _app.ActiveWorkbook;
            _projectWorkbook = projectWorkbook;
            _SheetUrv12 = ExcelHelper.GetSheet(wb, "Урв12");
            _projectManager = new ProjectManager.ProjectManager();
            _project = _projectManager.ActiveProject;
            _AnalisysSheet = ExcelHelper.GetSheet(wb, _project.AnalysisSheetName);
        }

        public void SetInfo()
        {
            int ix = 0;
            foreach (OfferAddress address in _projectWorkbook.OfferAddress)
            {
                SetColumns(address);
                PrintInfo(ix);
                ix++;
            }
        }

        private void PrintInfo(int ix)
        {
            int column = 6 + 5 * ix;
            int rowStart = 13;
            //string = "Описание менялось";
            int row = ExcelHelper.FindCell(_SheetUrv12, "Описание менялось").Row;
            _SheetUrv12.Cells[row, column].Formula = $"=IFERROR(COUNTIF({_rangeChengedNames}, \"ЛОЖЬ\"), \"#НД\")";

            row = ExcelHelper.FindCell(_SheetUrv12, "Объемы завышены").Row;
            _SheetUrv12.Cells[row, column].Formula = $"=IFERROR(COUNTIF({_rangeCostComments}, \"Расценки завышены\"), \"#НД\")";
            row = ExcelHelper.FindCell(_SheetUrv12, "Объемы занижены").Row;
            _SheetUrv12.Cells[row, column].Formula = $"=IFERROR(COUNTIF({_rangeCostComments}, \"Расценки занижены\"), \"#НД\")";

            int rowTotalSumm = ExcelHelper.FindCell(_SheetUrv12, "ОБЩАЯ СУММА РАСХОДОВ (без НДС)").Row;
            string cellAddress = _SheetUrv12.Cells[rowTotalSumm, column].Address;

            row = ExcelHelper.FindCell(_SheetUrv12, "Разница в стоимости с оценкой СПЕКТРУМ").Row;
            string addresBaseSumm = _SheetUrv12.Cells[rowTotalSumm, 4].Address;
            _SheetUrv12.Cells[row, column].Formula = $"= {addresBaseSumm} - {cellAddress}";


            row = ExcelHelper.FindCell(_SheetUrv12, "Итого включая не оцененные работы и корректировку ошибок").Row;
            _SheetUrv12.Cells[row, column].Formula = $"= {cellAddress}" +
                                                        $"+{_SheetUrv12.Cells[row - 1, column].Address}" +
                                                        $"+{_SheetUrv12.Cells[row - 2, column].Address}";


            Excel.Range rngOfferSum = _SheetUrv12.Range[_SheetUrv12.Cells[rowStart, column], _SheetUrv12.Cells[rowTotalSumm - 2, column]];
            Excel.Range rngOfferCommentCost = _SheetUrv12.Range[_SheetUrv12.Cells[rowStart, column + 3], _SheetUrv12.Cells[rowTotalSumm - 2, column + 3]];
            Excel.Range rngBasisSum = _SheetUrv12.Range[_SheetUrv12.Cells[rowStart, 4], _SheetUrv12.Cells[rowTotalSumm - 2, 4]];

            row = ExcelHelper.FindCell(_SheetUrv12, "Сумма завышеных работ по разделам").Row;

            _SheetUrv12.Cells[row, column].Formula = $"=SUMIF({rngOfferCommentCost.Address}, \">0.1\",{rngOfferSum.Address}) - {rngOfferCommentCost.Address}, \">0.1\" ,{rngBasisSum.Address})";

            row = ExcelHelper.FindCell(_SheetUrv12, "Выявленные ошибки").Row;
            _SheetUrv12.Cells[row, column].Formula = $"=SUMIF({rngOfferCommentCost.Address},\"#НД\",{rngOfferSum.Address} )";

            row = ExcelHelper.FindCell(_SheetUrv12, "НЕ оценено на сумму").Row;
            _SheetUrv12.Cells[row, column].Formula = $"=SUMIF({rngOfferCommentCost.Address}, \"=\"#НД\"\",{rngBasisSum.Address} )";
        }

        private void SetColumns(OfferAddress address)
        {
            int lastRow = _AnalisysSheet.UsedRange.Row + _AnalisysSheet.UsedRange.Rows.Count - 1;
            int rowStart = _project.RowStart;

            Excel.Range range = _AnalisysSheet.Range[
                    _AnalisysSheet.Cells[rowStart, address.ColStartOfferComments],
                    _AnalisysSheet.Cells[lastRow, address.ColStartOfferComments]];

            _rangeChengedNames = $"'{_AnalisysSheet.Name}'!{range.Address}";

            range = _AnalisysSheet.Range[
                    _AnalisysSheet.Cells[rowStart, address.ColPercentTotal + 1],
                    _AnalisysSheet.Cells[lastRow, address.ColPercentTotal + 1]];

            _rangeCostComments = $"'{_AnalisysSheet.Name}'!{range.Address}";

            //range = _AnalisysSheet.Range[
            //       _AnalisysSheet.Cells[rowStart, address.ColPercentTotal + 1],
            //       _AnalisysSheet.Cells[lastRow, address.ColPercentTotal + 1]];
            //_rangeCostTotal = $"'{_AnalisysSheet.Name}'!{range.Address}";
        }
    }
}
