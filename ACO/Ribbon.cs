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
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
                List<Offer> offers = new List<Offer>();
                ExcelHelpers.ExcelFile.Init();
                ExcelHelpers.ExcelFile.Acselerate(true);
            pb.SetMainBarVolum(files.Length);
            pb.Show(new AddinWindow(Globals.ThisAddIn));
                foreach (string fileName in files)
                {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.MainBarTick(fileName);
                    ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
                    excelBook.Open(fileName);
                    Excel.Worksheet sheet = excelBook.GetSheet(Offer.SheetName);
                    OfferReader reader = new OfferReader(sheet);
                    if (reader.ReadOffer())
                    { offers.Add(reader.Offer); }
                        excelBook.Close();
                }
                WriteOffers(offers,pb);
                ExcelHelpers.ExcelFile.Acselerate(false);
                ExcelHelpers.ExcelFile.Finish();            
        }

        private void WriteOffers(List<Offer> offers, IProgressBarWithLogUI pb )
        {
            Excel.Workbook mainBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            foreach(Offer offer in offers)
            {
                pb.SubBarTick();
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
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

        private void BtnProjectManager_Click(object sender, RibbonControlEventArgs e)
        {
            ProjectManager.FormManager manager = new ProjectManager.FormManager();
            manager.Show(new AddinWindow(Globals.ThisAddIn));
        }
    }
}
