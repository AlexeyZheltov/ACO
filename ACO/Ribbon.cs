using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System;

namespace ACO
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        /// <summary>
        /// Диалог выбора файлов КП
        /// </summary>
        /// <returns></returns>
        private string[] GetFiles()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Документ Excel|*.xls*|All files|*.*";
            openFileDialog.Title = "Выберите файлы КП";
            openFileDialog.Multiselect = true;
            string[] files = default;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                files = openFileDialog.FileNames;
            }
            return files;
        }

        /// <summary>
        /// Загрузка КП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLoadKP_Click(object sender, RibbonControlEventArgs e)
        {
            string[] files = GetFiles();
            if ( files.Length < 1) { return; }
                List<Offer> offers = new List<Offer>();
                ExcelHelpers.ExcelFile.Init();
                ExcelHelpers.ExcelFile.Acselerate(true);
                foreach (string fileName in files)
                {
                    ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
                    excelBook.Open(fileName);
                    Excel.Worksheet sheet = excelBook.GetSheet(Offer.SheetName);
                    OfferReader reader = new OfferReader(sheet);
                    if (reader.ReadOffer())
                    { offers.Add(reader.Offer); }
                        excelBook.Close();
                }
                WriteOffers(offers);
                ExcelHelpers.ExcelFile.Acselerate(false);
                ExcelHelpers.ExcelFile.Finish();            
        }

        private void WriteOffers(List<Offer> offers)
        {
            Excel.Workbook mainBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            
            foreach(Offer offer in offers)
            {

            }
        }

        /// <summary>
        ///  Создание проекта сравнения КП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCreateProgect_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void BtnAbout_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void BtnLoadLvl12_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void BtnUpdateLvl12_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
