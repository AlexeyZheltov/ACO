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
using System.Text.RegularExpressions;

namespace ACO
{
    public partial class Ribbon
    {
        Excel.Application _app = null;


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
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            if (pb is null)
            {
                pb.CloseForm += () => { pb = null; };
                pb.Show();
                // _pb.Show(new AddinWindow(Globals.ThisAddIn));
            }
            pb.ClearMainBar();
            pb.ClearSubBar();
            pb.SetMainBarVolum(files.Length);

            ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            foreach (string fileName in files)
            {
                try
                {
                    if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                    pb.MainBarTick(fileName);
                    excelBook.Open(fileName);
                    OfferWriter offerWriter = new OfferWriter(excelBook);
                    await Task.Run(() =>
                    {
                        offerWriter.Print(pb, offerSettingsName);
                    });
                }
                catch (AddInException ex)
                {
                    TextBox tb = pb.GetLogTextBox();
                    string message = $"Ошибка:{ex.Message }";
                    if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                    message += Environment.NewLine;
                    tb.Text += message;
                }
                finally
                {
                    if (pb.IsAborted)
                    {
                        pb.ClearMainBar();
                        pb.ClearSubBar();
                        pb.IsAborted = false;
                        MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    pb.CloseFrm();
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
            ExcelHelpers.ExcelFile.Acselerate(true);

            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            // if (pb is null)
            //{
            pb.CloseForm += () => { pb = null; };
            pb.Show(new AddinWindow(Globals.ThisAddIn));
            //}
            pb.ClearMainBar();
            pb.ClearSubBar();
            pb.SetMainBarVolum(1);


            ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();

            try
            {
                if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                pb.MainBarTick(file);
                excelBook.Open(file);
                await Task.Run(() =>
                {
                    OfferWriter offerWriter = new OfferWriter(excelBook);
                    offerWriter.PrintSpectrum(pb);
                });
            }
            catch (AddInException ex)
            {
                TextBox tb = pb.GetLogTextBox();
                string message = $"Ошибка:{ex.Message }";
                if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                message += Environment.NewLine;
                tb.Text += message; //"Ошибка:" + ex.Message + " (" + ex?.InnerException.Message + ")" + Environment.NewLine;
            }
            finally
            {
                if (pb.IsAborted)
                {
                    pb.ClearSubBar();
                    pb.ClearMainBar();
                    pb.IsAborted = false;
                    MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                pb.CloseFrm();
                excelBook.Close();
                ExcelHelpers.ExcelFile.Acselerate(false);
                ExcelHelpers.ExcelFile.Finish();
            }
        }

        /// <summary>
        ///  Форма вы бора настроек КП
        /// </summary>
        /// <returns></returns>
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

        /// <summary>
        ///  Обновление формул, Окраска уровней списка
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnUpdateFormuls_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelpers.ExcelFile.Init();
            ExcelHelpers.ExcelFile.Acselerate(true);
            try
            {
                UpdateFormate();
            }
            catch (AddInException addInEx)
            {
                MessageBox.Show(addInEx.Message, "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ExcelFile.Acselerate(false);
                ExcelFile.Finish();
            }
        }

        /// <summary>
        /// Обновление формул, Окраска уровней списка
        /// </summary>
        private async void UpdateFormate()
        {
            #region Подготовка прогресс бара
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.CloseForm += () => { pb = null; };
            pb.Show(new AddinWindow(Globals.ThisAddIn));
            pb.ClearMainBar();
            pb.ClearSubBar();
            pb.SetMainBarVolum(5);
            #endregion 

            pb.MainBarTick("Подготвка");
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

            PbAbortedFinish(pb);
            pb.MainBarTick("Подготовка списка");
            pb.SetSubBarVolume(root.AllCount());
            await Task.Run(() =>
            {
                root.Numeric(new Numberer(), pb); //pb.SubBarCount(root.AllCount)
                letter = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            });
            pb.ClearSubBar();
            pb.SetSubBarVolume(root.AllCount());
            await Task.Run(() =>
            {
                ExcelHelper.Write(ws, root, pb, letter); //pb.SubBarCount(root.AllCount)
            });

            PbAbortedFinish(pb);
            pb.MainBarTick("Запись формул");
            string letterAmount = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Count]).ColumnSymbol;
            string letterMaterialPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostMaterialsPerUnit]).ColumnSymbol;
            string letterMaterialTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostMaterialsTotal]).ColumnSymbol;
            string letterWorkPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostWorksPerUnit]).ColumnSymbol;
            string letterWorkTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostWorksTotal]).ColumnSymbol;
            string letterPricePerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotalPerUnit]).ColumnSymbol;
            string letterTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol;
            await Task.Run(() =>
            {
                //раз
                FMapping mappin = new FMapping()
                {
                    Amount = letterAmount,
                    MaterialPerUnit = letterMaterialPerUnit,
                    MaterialTotal = letterMaterialTotal,
                    WorkPerUnit = letterWorkPerUnit,
                    WorkTotal = letterWorkTotal,
                    PricePerUnit = letterPricePerUnit,
                    Total = letterTotal
                };
                //два
                ExcelHelper.SetFormulas(ws, mappin, root, pb); //Прогресс бар только для отмены
            });
            // ExcelHelper.SetFormulas(ws, mappin.Shift(ws, 10), root, null);

            PbAbortedFinish(pb);
            pb.MainBarTick("Форматирование списка");

            //три
            pb.ClearSubBar();
            int count = root.AllCount();
            pb.SetSubBarVolume(count);
            var pallet = ExcelReader.ReadPallet(pws);
            letter = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            //четыре
            List<(string, string)> colored_columns = GetColoredColumns(ws);
            colored_columns.Add(("A", letterTotal));
            (string, string)[] columns = colored_columns.ToArray();
            //columns.add
            ExcelHelper.Repaint(ws, pallet, project.RowStart, letter, pb, columns);//pb.SubBarCount(root.AllCount)

            //await Task.Run(() =>
            //{
            //});

            PbAbortedFinish(pb);
            pb.MainBarTick("Группировка списка");
            pb.ClearSubBar();
            await Task.Run(() =>
            {
                ExcelHelper.Group(ws, pb, letter); //Этот метод сам установит Max для прогрессбара
            });

            pb.ClearSubBar();
            pb.ClearMainBar();
            pb.CloseFrm();
        }

        private List<(string, string)> GetColoredColumns(Excel.Worksheet ws)
        {
            List<(string, string)> colored_columns = new List<(string, string)>();
            //int lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count;
            int lastCol = ws.Cells[1, ws.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;

             Excel.Range cellEnd = null;
            Excel.Range cellStart = null;
            for (int col = 1; col <= lastCol; col++)
            {
                Excel.Range cell = ws.Cells[1, col];
                string val = cell.Value?.ToString() ?? "";
                if (val == "offer_start")
                {
                    cellStart = cell;
                }
                else if (val == "offer_end")
                {
                    cellEnd = cell;
                }
                if (cellStart != null && cellEnd != null)
                {
                    if (cellStart.Column < cellEnd.Column)
                    {
                        string address = cellStart.Address;
                        string startOfferLetter = Regex.Match(address, @"[A-Z]+").Value ?? "";
                        address = cellEnd.Address;
                        string endOfferLetter = Regex.Match(address, @"[A-Z]+").Value ?? "";

                       if (!string.IsNullOrEmpty(startOfferLetter) && !string.IsNullOrEmpty(endOfferLetter)) colored_columns.Add((startOfferLetter, endOfferLetter));
                    }
                    cellStart = null;
                    cellEnd = null;
                }
            }
            return colored_columns;
        }

        private void PbAbortedFinish(IProgressBarWithLogUI pb)
        {
            if (pb.IsAborted)
            {
                pb.ClearMainBar();
                pb.ClearSubBar();
                pb.IsAborted = false;
                pb.CloseFrm();
                throw new AddInException("Выполнение было прервано.");
            }
        }
    }
}
