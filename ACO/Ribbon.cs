using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using System.IO;
using ACO.Offers;
using ACO.ExcelHelpers;
using ACO.ProjectManager;
using ACO.ProjectBook;
using System.Drawing;
using Microsoft.Office.Interop.Excel;


namespace ACO
{
    public partial class Ribbon
    {
        Excel.Application _app = null;
        private void ExcelAcselerate(bool mode)
        {
            _app.Calculation = mode ? Excel.XlCalculation.xlCalculationManual : Excel.XlCalculation.xlCalculationAutomatic;
            _app.ScreenUpdating = !mode;
            _app.DisplayAlerts = !mode;
        }

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
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Документ Excel|*.xls*|All files|*.*",
                Title = "Выберите файлы",
                Multiselect = true
            };
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
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel|*.xl*|All files|*.*",
                Title = "Выберите файл",
                Multiselect = false
            };
            string file = default;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                file = openFileDialog.FileName;
            }
            return file;
        }

        /// <summary>
        /// Загрузка КП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnLoadKP_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return; // ячейка редактируется
            string[] files = GetFiles();
            if ((files?.Length ?? 0) < 1) return;

            string offerSettingsName = GetOfferSettings();
            if (string.IsNullOrEmpty(offerSettingsName)) return;

            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.Show(new AddinWindow(Globals.ThisAddIn));
            await Task.Run(() =>
            {
                int count = files.Length;
                pb.SetMainBarVolum(count);

                ExcelHelpers.ExcelFile.Init();
                ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();

                pb.Writeline("Инициализацция диспетчера проектов.");
                ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
                foreach (string fileName in files)
                {
                    try
                    {
                        if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                        pb.Writeline("Открытие файла");
                        pb.MainBarTick(fileName);
                        excelBook.Open(fileName);
                        pb.Writeline("Инициализация загрузчика");
                        OfferWriter offerWriter = new OfferWriter(excelBook);
                        pb.Writeline("Заполнение листа Анализ\n");

                        //   await Task.Run(() =>
                        //    {
                        ExcelAcselerate(true);
                        offerWriter.Print(pb, offerSettingsName);
                        pb.Writeline("Завершение");
                       // ExcelAcselerate(false);
                        pb.CloseFrm();
                        //  });
                    }
                    catch (AddInException addInEx)
                    {
                        string message = $"Ошибка:{addInEx.Message }";
                        if (addInEx.InnerException != null) message += $"{addInEx.InnerException.Message}";
                        pb.Writeline(message);
                    }
                    catch (Exception ex)
                    {
                        string message = $"Ошибка:{ex.Message }";
                        if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                        MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        excelBook.Close();
                        ExcelHelpers.ExcelFile.Finish();
                        ExcelAcselerate(false);
                    }
                }
            });
        }

        private async void BtnSpectrum_Click(object sender, RibbonControlEventArgs e)
        {
            string file = GetFile();
                    if (!File.Exists(file)) { return; }
                IProgressBarWithLogUI pb = new ProgressBarWithLog();
                    pb.Show(new AddinWindow(Globals.ThisAddIn));
            await Task.Run(() =>
            {
                try
                {
                    pb.SetMainBarVolum(1);
                    // PrintSpectrum(pb, file);
                    ExcelHelpers.ExcelFile.Init();
                    ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
                    if (pb.IsAborted) throw new AddInException("Процесс остановлен");

                    pb.Writeline($"Открытие файла :");
                    pb.MainBarTick(file);
                    excelBook.Open(file);

                    pb.Writeline("Инициализация загрузчика.");
                    OfferWriter offerWriter = new OfferWriter(excelBook);

                    if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                    pb.Writeline("Заполнение листа Анализ.");
                    //await Task.Run(() =>
                    //{
                    ExcelAcselerate(true);
                    offerWriter.PrintSpectrum(pb);
                    pb.Writeline("Завершение.");
                   // ExcelAcselerate(false);
                    pb.CloseFrm();
                }
                catch (AddInException ex)
                {
                    string message = $"Ошибка:{ex.Message }";
                    if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                    pb.Writeline(message);
                }
                catch (Exception ex)
                {
                    string message = $"Ошибка:{ex.Message }";
                    if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                    MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    ExcelAcselerate(false);
                }
            });
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
        ///  Создание проекта сравнения КП. Открыть Шаблон. Сохранить как 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnCreateProgect_Click(object sender, RibbonControlEventArgs e)
        {
            string pathTamplate = Properties.Settings.Default.TamplateProgectPath;
            string path;
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
            form.ShowDialog(new AddinWindow(Globals.ThisAddIn));
        }

        private async void BtnUpdateFormuls_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return; // ячейка редактируется
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.Show();
            await Task.Run(() =>
            {
                try
                {
                    ExcelAcselerate(true);
                    UpdateFormuls(pb);
                    pb.CloseFrm();
                }
                catch (AddInException addInEx)
                {
                    string message = $"Ошибка:{addInEx.Message }";
                    if (addInEx.InnerException != null) message += $"{addInEx.InnerException.Message}";
                    pb.Writeline(message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    ExcelAcselerate(false);
                }
            });
        }

        /// <summary>
        /// Формулы, окраска уровней
        /// </summary>
        private void UpdateFormuls(IProgressBarWithLogUI pb)
        {
            pb.SetMainBarVolum(6);
            pb.MainBarTick("Подготвка");
            // ExcelAcselerate(true);
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            ProjectManager.Project project = projectManager.ActiveProject;

            Excel.Worksheet ws = ExcelHelper.GetSheet(wb, project.AnalysisSheetName);
            Excel.Worksheet pws = ExcelHelper.GetSheet(wb, "Палитра");

            //======1=======
            pb.MainBarTick("Разгруппировать список");
            ExcelHelper.UnGroup(ws);
            PbAbortedStopProcess(pb);

            HItem root = new HItem();
            string letterLevel = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;

            foreach (var (Row, Level) in ExcelReader.ReadSourceItems(ws, letterLevel, project.RowStart))
                root.Add(new HItem()
                {
                    Level = Level,
                    Row = Row
                });
            PbAbortedStopProcess(pb);

            //======2=======
            pb.MainBarTick("Подготовка списка");
            pb.SetSubBarVolume(root.AllCount());
            root.Numeric(new Numberer(), pb);
            string letterNumber = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            pb.ClearSubBar();
            pb.SetSubBarVolume(root.AllCount());

            ExcelHelper.Write(ws, root, pb, letterNumber);
            PbAbortedStopProcess(pb);
            //======3=======
            pb.MainBarTick("Запись формул");
            string letterAmount = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Amount]).ColumnSymbol;
            string letterMaterialPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostMaterialsPerUnit]).ColumnSymbol;
            string letterMaterialTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostMaterialsTotal]).ColumnSymbol;
            string letterWorkPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostWorksPerUnit]).ColumnSymbol;
            string letterWorkTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostWorksTotal]).ColumnSymbol;
            string letterPricePerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotalPerUnit]).ColumnSymbol;
            string letterTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol;
            string letterComment = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Comment]).ColumnSymbol;

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
                                                           // Обновление формул КП
            HashSet<int> columnsAmount = GetNumbersCoumnsOfCount(ws);
            foreach (int col in columnsAmount)
            {
                ExcelHelper.SetFormulas(ws, mappin.Shift(ws, col), root, pb);
            }

            pb.ClearSubBar();
            PbAbortedStopProcess(pb);

            //======4=======
            pb.MainBarTick("Форматирование списка");
            //три
            var pallet = ExcelReader.ReadPallet(pws);
            int count = ws.UsedRange.Rows.Count;
            pb.SetSubBarVolume(count);
            List<(string, string)> colored_columns = ProjectWorkbook.GetColredColumns(ws);
            colored_columns.Add(("A", letterComment));
            (string, string)[] columns = colored_columns.ToArray();

            //четыре
            ExcelHelper.Repaint(ws, pallet, project.RowStart, letterLevel, pb, columns);

            List<(string, string)> columns_format = ProjectWorkbook.GetFormatColumns(ws);
            ExcelHelper.SetNumberFormat(ws, project.RowStart, columns_format.ToArray());
            ExcelHelper.SetNumberFormat(ws, project.RowStart, letterAmount);
            ExcelHelper.SetNumberFormat(ws, project.RowStart, letterMaterialPerUnit);
            ExcelHelper.SetNumberFormat(ws, project.RowStart, letterMaterialTotal);
            ExcelHelper.SetNumberFormat(ws, project.RowStart, letterWorkPerUnit);
            ExcelHelper.SetNumberFormat(ws, project.RowStart, letterWorkTotal);
            ExcelHelper.SetNumberFormat(ws, project.RowStart, letterPricePerUnit);
            ExcelHelper.SetNumberFormat(ws, project.RowStart, letterTotal);

            PbAbortedStopProcess(pb);
            pb.MainBarTick("Группировка списка");
            pb.ClearSubBar();


            ExcelHelper.Group(ws, pb, letterLevel); //Этот метод сам установит Max для прогрессбара
                                                    // ExcelAcselerate(false);
            pb.ClearMainBar();
        }



        /// <summary>
        ///  Определить номера столбцов с ко-ом для загруженных П
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        private HashSet<int> GetNumbersCoumnsOfCount(Excel.Worksheet ws)
        {
            HashSet<int> columnsAmount = new HashSet<int>();
            int lastCol = ws.Cells[1, ws.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
            for (int col = 1; col <= lastCol; col++)
            {
                Excel.Range cell = ws.Cells[1, col];
                string val = cell.Value?.ToString() ?? "";

                if (val == Project.ColumnsNames[StaticColumns.Amount])
                {
                    columnsAmount.Add(cell.Column);
                }
            }
            return columnsAmount;
        }

        /// <summary>
        ///  Прогресс бар. нажата кнопка прервать
        /// </summary>
        /// <param name="pb"></param>
        private void PbAbortedStopProcess(IProgressBarWithLogUI pb)
        {
            if (pb.IsAborted)
            {
                pb.ClearMainBar();
                pb.ClearSubBar();
                pb.IsAborted = false;
                pb.CloseFrm();
                throw new AddInException("Выполнение было прервано");
            }
        }

        /// <summary>
        ///  Окраска ячеек 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnColorComments_Click(object sender, RibbonControlEventArgs e)
        {
            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            int lastRow = projectWorkbook.AnalisysSheet.UsedRange.Row + projectWorkbook.AnalisysSheet.UsedRange.Rows.Count - 1;
            int startRow = projectWorkbook.GetFirstRow();
            int count = lastRow - startRow + 1;
            if (count > 0)
            {
                IProgressBarWithLogUI pb = new ProgressBarWithLog();
                pb.Show();
                await Task.Run(() =>
                {
                    try
                    {
                        pb.SetMainBarVolum(1);
                        pb.MainBarTick("Уловное форматирование ячеек комментариев");
                        pb.SetSubBarVolume(count);
                        ExcelAcselerate(true);
                        int k = 0;
                        for (int row = startRow; row <= lastRow; row++)
                        {
                            pb.SubBarTick();
                            PbAbortedStopProcess(pb);
                            foreach (OfferAddress offeraddress in projectWorkbook.OfferAddress)
                            {
                                projectWorkbook.ColorCell(projectWorkbook.AnalisysSheet.Cells[row, offeraddress.ColPercentWorks]);
                                projectWorkbook.ColorCell(projectWorkbook.AnalisysSheet.Cells[row, offeraddress.ColPercentMaterial]);
                                projectWorkbook.ColorCell(projectWorkbook.AnalisysSheet.Cells[row, offeraddress.ColPercentTotal]);
                            }
                        }
                      //  ExcelAcselerate(false);
                        pb.CloseFrm();

                    }
                    catch (AddInException addinEx)
                    {
                        pb.Writeline(addinEx.Message);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        ExcelAcselerate(false);
                    }
                });
            }
        }

        /// <summary>
        ///  Запись формул на уровень
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnLoadLvl11_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return; // ячейка редактируется
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.Show();
            await Task.Run(() =>
            {
                ExcelAcselerate(true);
                try
                {
                    new PivotSheets.Pivot().LoadUrv11(pb);
                    pb.CloseFrm();
                }
                catch (AddInException addinEx)
                {
                    pb.Writeline(addinEx.Message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    ExcelAcselerate(false);
                }
            });
        }

        /// <summary>
        ///  Запись формул на уровень 12
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BtnLoadLvl12_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return; // ячейка редактируется
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.Show();

            await Task.Run(() =>
            {
                try
                {
                    ExcelAcselerate(true);
                    new PivotSheets.Pivot().LoadUrv12(pb);
                    pb.CloseFrm();
                }
                catch (AddInException addinEx)
                {
                    pb.Writeline(addinEx.Message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    ExcelAcselerate(false);
                }
            });
        }


        private async void BtnUpdateLvl11_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return; // ячейка редактируется
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            await Task.Run(() =>
            {
                try
                {
                    ExcelAcselerate(true);
                    pb.Show();
                    new PivotSheets.Pivot().UpdateUrv11(pb);
                    pb.CloseFrm();
                }
                catch (AddInException addinEx)
                {
                    pb.Writeline(addinEx.Message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    ExcelAcselerate(false);
                }
            });
        }

        private async void BtnUpdateLvl12_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return; // ячейка редактируется
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.Show();
            await Task.Run(() =>
            {
                try
                {
                    ExcelAcselerate(true);
                    new PivotSheets.Pivot().UpdateUrv12(pb);
                    pb.CloseFrm();
                }
                catch (AddInException addinEx)
                {
                    pb.Writeline(addinEx.Message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    ExcelAcselerate(false);
                }
            });
        }

        private void BtnExcelScreenUpdating_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelAcselerate(false);
        }
    }
}
