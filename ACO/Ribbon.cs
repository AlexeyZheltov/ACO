using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using System.IO;
using ACO.Offers;
using ACO.Settings;
using ACO.ExcelHelpers;
using ACO.ProjectManager;

namespace ACO
{
    public partial class Ribbon
    {
        Excel.Application _app = null;
        IProgressBarWithLogUI _pb;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            _app = Globals.ThisAddIn.Application;
        }

        /// <summary>
        /// Загрузка КП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnLoadKP_Click(object sender, RibbonControlEventArgs e)
        {
            string[] files = GetFiles();
            if ((files?.Length ?? 0) < 1) { return; }

            string offerSettingsName = GetOfferSettings();


            ExcelHelpers.ExcelFile.Init();
            //ExcelHelpers.ExcelFile.Acselerate(true);
            if (_pb is null)
            {
                _pb = new ProgressBarWithLog();
                _pb.CloseForm += () => { _pb = null; };
                _pb.Show();
                // _pb.Show(new AddinWindow(Globals.ThisAddIn));
            }
            _pb.ClearMainBar();
            _pb.ClearSubBar();
            _pb.SetMainBarVolum(files.Length);

            ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            foreach (string fileName in files)
            {
                try
                {
                    if (_pb.IsAborted) throw new AddInException("Процесс остановлен");
                    _pb.MainBarTick(fileName);
                    excelBook.Open(fileName);
                    OfferWriter offerWriter = new OfferWriter(excelBook);
                    await Task.Run(() =>
                    {
                        offerWriter.Print(_pb, offerSettingsName);
                    });
                }
                catch (AddInException ex)
                {
                    TextBox tb = _pb.GetLogTextBox();
                    string message = $"Ошибка:{ex.Message }";
                    if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                    message += Environment.NewLine;
                    tb.Text += message;
                }
                finally
                {
                    if (_pb.IsAborted)
                    {
                        _pb.ClearMainBar();
                        _pb.ClearSubBar();
                        _pb.IsAborted = false;
                        MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    _pb.CloseFrm();
                    excelBook.Close();
                    ExcelHelpers.ExcelFile.Acselerate(false);
                    ExcelHelpers.ExcelFile.Finish();

                }


            }

        }

        private async void BtnSpectrum_Click(object sender, RibbonControlEventArgs e)
        {
            string file = GetFile();
            if (!File.Exists(file)) { return; }

            ExcelHelpers.ExcelFile.Init();
            //ExcelHelpers.ExcelFile.Acselerate(true);
            if (_pb is null)
            {
                _pb = new ProgressBarWithLog();
                _pb.CloseForm += () => { _pb = null; };
                _pb.Show(new AddinWindow(Globals.ThisAddIn));
            }
            _pb.ClearMainBar();
            _pb.ClearSubBar();
            _pb.SetMainBarVolum(1);
            //_pb.SetMainBarVolum(file);

            ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();

            try
            {
                if (_pb.IsAborted) throw new AddInException("Процесс остановлен");
                _pb.MainBarTick(file);
                excelBook.Open(file);
                await Task.Run(() =>
                {
                    OfferWriter offerWriter = new OfferWriter(excelBook);
                    offerWriter.PrintSpectrum(_pb);
                });
            }
            catch (AddInException ex)
            {
                TextBox tb = _pb.GetLogTextBox();
                string message = $"Ошибка:{ex.Message }";
                if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                message += Environment.NewLine;
                tb.Text += message; //"Ошибка:" + ex.Message + " (" + ex?.InnerException.Message + ")" + Environment.NewLine;
            }
            finally
            {
                if (_pb.IsAborted)
                {
                    _pb.ClearSubBar();
                    _pb.ClearMainBar();
                    _pb.IsAborted = false;
                    MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                _pb.CloseFrm();
                excelBook.Close();
                ExcelHelpers.ExcelFile.Acselerate(false);
                ExcelHelpers.ExcelFile.Finish();
            }
        }

        private string GetOfferSettings()
        {
            string settingsFile = "";
            FormSelectOfferSettings form = new FormSelectOfferSettings();
            if (form.ShowDialog(new AddinWindow(Globals.ThisAddIn)) == DialogResult.OK)
            {
                settingsFile = form.OfferSettingsName ?? "";
            }
            return settingsFile;
        }


        /// <summary>
        /// Диалог выбора файлов КП
        /// </summary>
        /// <returns></returns>
        private string[] GetFiles()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Документ Excel|*.xls*|All files|*.*";
            openFileDialog.Title = "Выберите файлы";
            openFileDialog.Multiselect = true;
            string[] files = default;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                files = openFileDialog.FileNames;
            }
            return files;
        }

        /// <summary>
        /// Диалог выбора файла Шаблона
        /// </summary>
        /// <returns></returns>
        private string GetFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel|*.xl*|All files|*.*";
            openFileDialog.Title = "Выберите файл";
            openFileDialog.Multiselect = false;
            string file = default;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog.FileName;
            }
            return file;
        }

        /// <summary>
        ///  Создание проекта сравнения КП. Открыть Шаблон. Сохранить как 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCreateProgect_Click(object sender, RibbonControlEventArgs e)
        {
            string pathTamplate = Properties.Settings.Default.TamplateProgectPath;
            string path = default;
            if (!File.Exists(pathTamplate))
            {
                path = GetFile();
                if (!File.Exists(path)) return;
                Properties.Settings.Default.TamplateProgectPath = path;
                Properties.Settings.Default.Save();
            }
            else { path = pathTamplate; }
            Excel.Workbook newProjectBook = _app.Workbooks.Open(path);
            newProjectBook.Activate();
            _app.Dialogs[Excel.XlBuiltInDialog.xlDialogSaveAs].Show();
        }

        private void BtnProjectManager_Click(object sender, RibbonControlEventArgs e)
        {
            ProjectManager.FormManager manager = new ProjectManager.FormManager();
            manager.Show(new AddinWindow(Globals.ThisAddIn));
        }

        /// <summary>
        ///  Диспетчер КП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnKP_Click(object sender, RibbonControlEventArgs e)
        {
            new Offers.FormManagerKP().Show(new AddinWindow(Globals.ThisAddIn));
        }

        /// <summary>
        ///  Настройки
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnSettings_Click(object sender, RibbonControlEventArgs e)
        {
            FormSettings form = new FormSettings();
            //new AddinWindow(Globals.ThisAddIn)
            form.ShowDialog();
        }

        private void BtnUpdateFormuls_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = wb.ActiveSheet;
            Excel.Worksheet pws = wb.Sheets["Палитра"];

            ExcelHelper.UnGroup(ws);

            HItem root = new HItem();
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            ProjectManager.Project project = projectManager.ActiveProject;
            string letter = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            foreach (var (Row, Level) in ExcelReader.ReadSourceItems(ws, letter, project.RowStart))
                root.Add(new HItem()
                {
                    Level = Level,
                    Row = Row
                });

            root.Numeric(new Numberer(), null); //pb.SubBarCount(root.AllCount)

            letter = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            ExcelHelper.Write(ws, root, null, letter); //pb.SubBarCount(root.AllCount)

            //раз
            FMapping mappin = new FMapping()
            {
                Amount = "",
                MaterialPerUnit = "",
                MaterialTotal = "",
                WorkPerUnit = "",
                WorkTotal = "",
                PricePerUnit = "",
                Total = ""
            };

            //два
            ExcelHelper.SetFormulas(ws, mappin, root, null); //Прогресс бар только для отмены
            ExcelHelper.SetFormulas(ws, mappin.Shift(ws, 10), root, null);

            //три
            var pallet = ExcelReader.ReadPallet(pws);
            letter = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;

            //четыре
            ExcelHelper.Repaint(ws, pallet, project.RowStart, letter, null, ("A", "B"), ("AA", "AB"));//pb.SubBarCount(root.AllCount)

            ExcelHelper.Group(ws, null, letter); //Этот метод сам установит Max для прогрессбара
        }
    }
}
