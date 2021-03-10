using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;

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


        IProgressBarWithLogUI _pb;

        /// <summary>
        /// Загрузка КП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnLoadKP_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                string[] files = GetFiles();
                if (files.Length < 1) { return; }
                List<Offer> offers = new List<Offer>();
                ExcelHelpers.ExcelFile.Init();
                ExcelHelpers.ExcelFile.Acselerate(true);


                if (_pb is null)
                {
                    _pb = new ProgressBarWithLog();
                    _pb.CloseForm += () => { _pb = null; };
                    _pb.Show(new AddinWindow(Globals.ThisAddIn));
                }
                _pb.ClearMainBar();
                _pb.ClearSubBar();
                _pb.SetMainBarVolum(files.Length);
                // _pb.MainBarTick("Подключение к Excel");
                await Task.Run(() =>
                {
                    ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
                    foreach (string fileName in files)
                    {
                        try
                        {
                            if (_pb.IsAborted) throw new AddInException("Процесс остановлен");
                            _pb.MainBarTick(fileName);

                            excelBook.Open(fileName);
                            OfferManager offerReader = new OfferManager(excelBook);
                            //new OfferManager(sheet);
                            //
                            // Offer offer = offerReader.ReadFromSheet(sheet);
                            if (offerReader.ReadOffer())
                            { offers.Add(offerReader.Offer); }
                        }
                        catch (AddInException ex)
                        {
                            MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        finally
                        {
                            excelBook.Close();
                            ExcelHelpers.ExcelFile.Acselerate(false);
                            ExcelHelpers.ExcelFile.Finish();
                        }
                    }
                    WriteOffers(offers, _pb);

                    if (_pb.IsAborted)
                    {
                        _pb.ClearMainBar();
                        _pb.ClearSubBar();
                        _pb.IsAborted = false;
                        MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    

                });
            }
            catch (AddInException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK);
            }

        }

        private void WriteOffers(List<Offer> offers, IProgressBarWithLogUI pb)
        {
            Excel.Workbook mainBook = Globals.ThisAddIn.Application.ActiveWorkbook;
            pb.SetSubBarVolume(offers.Count);
            ProjectManager.ProjectManager project = new ProjectManager.ProjectManager();
            foreach (Offer offer in offers)
            {
                pb.SubBarTick();
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                project.AddOffer(offer);
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

        private void BtnKP_Click(object sender, RibbonControlEventArgs e)
        {
            new Offers.FormManagerKP().Show(new AddinWindow(Globals.ThisAddIn));
        }
    }
}
