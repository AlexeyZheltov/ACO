using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using System.IO;

namespace ACO
{
    public partial class Ribbon
    {
        Excel.Application _app = null ;
        IProgressBarWithLogUI _pb;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            _app = Globals.ThisAddIn.Application;
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
                    ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();

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
                            {
                                //projectManager.PrintOffer(offerReader.Offer);
                                offers.Add(offerReader.Offer);
                            }
                        }
                        catch (AddInException ex)
                        {
                            TextBox tb = _pb.GetLogTextBox();
                            tb.Text += "Ошибка:" + ex.Message + " (" + ex.InnerException.Message + ")" + Environment.NewLine;
                            //  MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void WriteOffers(List<Offer> offers, IProgressBarWithLogUI pb)
        {
            Excel.Workbook mainBook = _app.ActiveWorkbook;
            pb.SetSubBarVolume(offers.Count);
            //  ProjectManager.ProjectManager project = new ProjectManager.ProjectManager();
            //ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            Offers.OfferWriter offerWriter = new Offers.OfferWriter();

            foreach (Offer offer in offers)
            {
                pb.SubBarTick();
                offerWriter.PrintOffer(offer);
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                //    project.AddOffer(offer);
            }
        }

        /// <summary>
        ///  Создание проекта сравнения КП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCreateProgect_Click(object sender, RibbonControlEventArgs e)
        {
            string pathTamplate = Properties.Settings.Default.TamplateProgectPath;
            string path = default;
            if (!File.Exists(pathTamplate))
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Документ Excel|*.xls*|All files|*.*";
                openFileDialog.Title = "Выберите файл шаблона проекта";
                openFileDialog.Multiselect = false;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    path = openFileDialog.FileName;
                    if (!File.Exists(path)) return;
                }
            }
            else { path = pathTamplate; }
            Excel.Workbook newProjectBook = _app.Workbooks.Open(path);
            newProjectBook.Activate();
            _app.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();


            /*  Dim varResult As Variant
            Dim ActBook As Workbook
            'displays the save file dialog
            varResult = Application.GetSaveAsFilename(FileFilter:= _
                     "Excel Files (*.xlsx), *.xlsx", Title:="Save PO", _
                    InitialFileName:="\\showdog\service\Service_job_PO\")
            'checks to make sure the user hasn't canceled the dialog
            If varResult <> False Then
                ActiveWorkbook.SaveAs Filename:=varResult, _
                FileFormat:=xlWorkbookNormal
                Exit Sub
            End If*/
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
