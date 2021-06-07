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
using ACO.Settings;

namespace ACO
{
    public partial class Ribbon
    {
        readonly ACO.Properties.Settings settings = ACO.Properties.Settings.Default;
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
                        ExcelAcselerate(true);
                        offerWriter.Print(pb, offerSettingsName);
                        pb.Writeline("Формулы анализа");
                        SetAnalysis();
                        pb.Writeline("Фильтр");
                        SetDataFilter();
                        pb.Writeline("Группировка столбцов");
                        GroupColumns();
                        pb.Writeline("Завершение");
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

        private async void BtnBaseEstimate_Click(object sender, RibbonControlEventArgs e)
        {
            string file = GetFile();
            if (!File.Exists(file)) { return; }
            string offerSettingsName = GetOfferSettings();
            if (string.IsNullOrEmpty(offerSettingsName)) return;
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();
            pb.Show(new AddinWindow(Globals.ThisAddIn));
            await Task.Run(() =>
            {
                try
                {
                    pb.SetMainBarVolum(1);
                    ExcelHelpers.ExcelFile.Init();
                    if (pb.IsAborted) throw new AddInException("Процесс остановлен");

                    pb.Writeline($"Открытие файла :");
                    pb.MainBarTick(file);
                    excelBook.Open(file);

                    pb.Writeline("Инициализация загрузчика.");
                    OfferWriter offerWriter = new OfferWriter(excelBook);

                    if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                    pb.Writeline("Заполнение листа Анализ.");

                    ExcelAcselerate(true);
                    offerWriter.PrintBaseEstimate(pb, offerSettingsName);
                    pb.Writeline("Завершение.");
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
                    excelBook.Close();
                    ExcelHelpers.ExcelFile.Finish();
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
        private void BtnCreateProject_Click(object sender, RibbonControlEventArgs e)
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
            if (ExcelHelper.IsEditing()) return;
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

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            ProjectManager.Project project = projectManager.ActiveProject;

            Excel.Worksheet ws = ExcelHelper.GetSheet(wb, project.AnalysisSheetName);
            Excel.Worksheet pws = ExcelHelper.GetSheet(wb, "Палитра");

            ExcelHelper.CollapseColumns(ws);
            //======1=======
            pb.MainBarTick("Разгруппировать список");
            ExcelHelper.UnGroupRows(ws);
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

            Excel.Range celEndBasis = ws.Cells[1, ws.Range[$"{ letterComment}1"].Column + 8];
            string letterEndBasis = celEndBasis.Address.Split(new char[] { '$' }, StringSplitOptions.RemoveEmptyEntries)[0];

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
            colored_columns.Add(("A", letterEndBasis));
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
                    PivotSheets.Pivot pivot = new PivotSheets.Pivot(pb);
                    pivot.LoadUrv11();
                    pivot.SheetUrv11.Activate();
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
                    PivotSheets.Pivot pivot = new PivotSheets.Pivot(pb);
                    pivot.LoadUrv12();
                    pivot.SheetUrv12.Activate();

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

        private async void SptBtn_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return;
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.Show();

            await Task.Run(() =>
            {
                try
                {
                    ExcelAcselerate(true);
                    PivotSheets.Pivot pivot = new PivotSheets.Pivot(pb);
                    pivot.LoadUrv12();
                    pb.ClearSubBar();
                    pb.ClearMainBar();
                    pivot.LoadUrv11();
                    pivot.SheetUrv12.Activate();
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
                    new PivotSheets.Pivot(pb).UpdateUrv11();
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
                    new PivotSheets.Pivot(pb).UpdateUrv12();
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

        private void BtnFormatComments_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (new FrmColorCommentsFomat().ShowDialog() == DialogResult.OK)
                {
                    SetAnalysis();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ExcelAcselerate(false);
            }
        }

        private void SptBtnFormatComments_Click(object sender, RibbonControlEventArgs e)
        {
            SetAnalysis();
        }

        /// <summary>
        ///  Окраска комментариев
        ///  Формулы Анализа
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetAnalysis()
        {

            if (ExcelHelper.IsEditing()) return;
            try
            {
                ExcelAcselerate(true);
                ProjectWorkbook projectWorkbook = new ProjectWorkbook();
                SetAnalisysFormuls(projectWorkbook);

                // очистить условное форматирование
                projectWorkbook.AnalisysSheet.UsedRange.FormatConditions.Delete();
                ConditonsFormatManager formatManager = new ConditonsFormatManager();
                int lastRow = projectWorkbook.AnalisysSheet.UsedRange.Row + projectWorkbook.AnalisysSheet.UsedRange.Rows.Count + 1;
                int firstRow = projectWorkbook.GetFirstRow();
                foreach (OfferColumns offeraddress in projectWorkbook.OfferColumns)
                {
                    /// Works
                    List<ConditionFormat> conditions = formatManager.ListConditionFormats.FindAll(x => x.ColumnName ==
                                                     ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationWorks]);
                    Excel.Range rng = projectWorkbook.AnalisysSheet.Range[
                                projectWorkbook.AnalisysSheet.Cells[firstRow, offeraddress.ColDeviationWorks],
                               projectWorkbook.AnalisysSheet.Cells[lastRow, offeraddress.ColDeviationWorks]];
                    conditions.ForEach(x => x.SetCondition(rng));
                    /// Materials
                    conditions = formatManager.ListConditionFormats.FindAll(x => x.ColumnName ==
                                                ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationMat]);
                    rng = projectWorkbook.AnalisysSheet.Range[
                                projectWorkbook.AnalisysSheet.Cells[firstRow, offeraddress.ColDeviationMaterials],
                               projectWorkbook.AnalisysSheet.Cells[lastRow, offeraddress.ColDeviationMaterials]];
                    conditions.ForEach(x => x.SetCondition(rng));

                    /// Стоимость
                    conditions = formatManager.ListConditionFormats.FindAll(x => x.ColumnName ==
                                                ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationCost]);
                    rng = projectWorkbook.AnalisysSheet.Range[
                                projectWorkbook.AnalisysSheet.Cells[firstRow, offeraddress.ColDeviationCost],
                               projectWorkbook.AnalisysSheet.Cells[lastRow, offeraddress.ColDeviationCost]];
                    conditions.ForEach(x => x.SetCondition(rng));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                ExcelAcselerate(false);
            }
        }


        /// <summary>
        ///  Формулы анализа 
        /// </summary>
        private void SetAnalisysFormuls(ProjectWorkbook projectWorkbook)
        {
            int firstRow = projectWorkbook.GetFirstRow();
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            ProjectManager.Project project = projectManager.ActiveProject;
            Excel.Worksheet ws = projectWorkbook.AnalisysSheet;
            // Литеры стрлбцов базовой оценки
            string letterName = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Name]).ColumnSymbol;

            string letterAmount = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Amount]).ColumnSymbol;
            string letterCostPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotalPerUnit]).ColumnSymbol;

            string letterWorkPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostWorksPerUnit]).ColumnSymbol;
            string letterMaterialPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostMaterialsPerUnit]).ColumnSymbol;

            string letterEndBasis = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Comment]).ColumnSymbol;



            string top = settings.TopBoundAnalysis.ToString().Replace(',', '.');
            string bottom = settings.BottomBoundAnalysis.ToString().Replace(',', '.');

            int lastRow = ws.Range[$"{letterCostPerUnit}{ws.Rows.Count}"].End[Excel.XlDirection.xlUp].Row;


            // Ячейка  стоимость
            Excel.Range cellBasisCost = ws.Range[$"{letterCostPerUnit}{firstRow}"];
            string addressBasisCostPerUnit = cellBasisCost.Address[RowAbsolute: false, ColumnAbsolute: true];

            // базовая стоимость работ
            Excel.Range cellBasisWorks = ws.Range[$"{letterWorkPerUnit}{firstRow}"];
            string addressBasisWorks = cellBasisWorks.Address[RowAbsolute: false, ColumnAbsolute: true];

            // базовая стоимость материалов
            Excel.Range cellBasisMaterials = ws.Range[$"{letterMaterialPerUnit}{firstRow}"];
            string addressBasisMaterials = cellBasisMaterials.Address[RowAbsolute: false, ColumnAbsolute: true];

            Excel.Range cellBasisAmount = ws.Range[$"{letterAmount}{firstRow}"];
            string addressAmount = cellBasisAmount.Address[RowAbsolute: false, ColumnAbsolute: true];

            Excel.Range cellEndBasis = ws.Range[$"{letterEndBasis}{firstRow}"];

            /// ячейки Среднее. медиана
            Excel.Range CellAvgAmount = ws.Cells[cellEndBasis.Row, cellEndBasis.Column + 1];

            Excel.Range CellAvgCostMaterials = ws.Cells[cellEndBasis.Row, cellEndBasis.Column + 2];
            Excel.Range CellAvgCostWorks = ws.Cells[cellEndBasis.Row, cellEndBasis.Column + 3];
            Excel.Range CellAvgTotalCost = ws.Cells[cellEndBasis.Row, cellEndBasis.Column + 4];

            Excel.Range CellMedianCostMaterials = ws.Cells[cellEndBasis.Row, cellEndBasis.Column + 5];
            Excel.Range CellMedianCostWorks = ws.Cells[cellEndBasis.Row, cellEndBasis.Column + 6];
            Excel.Range CellMedianTotalCost = ws.Cells[cellEndBasis.Row, cellEndBasis.Column + 7];

            string addressAvgAmount = CellAvgAmount.Address[RowAbsolute: false, ColumnAbsolute: true];
            string addressAvgCostMaterials = CellAvgCostMaterials.Address[RowAbsolute: false, ColumnAbsolute: true];
            string addressAvgCostWorks = CellAvgCostWorks.Address[RowAbsolute: false, ColumnAbsolute: true];
            string addressAvgTotalCost = CellAvgTotalCost.Address[RowAbsolute: false, ColumnAbsolute: true];

            string addressMedianCostMaterials = CellMedianCostMaterials.Address[RowAbsolute: false, ColumnAbsolute: true];
            string addressMedianCostWorks = CellMedianCostWorks.Address[RowAbsolute: false, ColumnAbsolute: true];
            string addressMedianTotalCost = CellMedianTotalCost.Address[RowAbsolute: false, ColumnAbsolute: true];

            // Аргументы функции
            string argumentsCommonCost = addressBasisCostPerUnit;
            string argumentsCommonWorks = addressBasisWorks;
            string argumentsCommonMaterials = addressBasisMaterials;
            string argumentsAmount = addressAmount;

            foreach (OfferColumns offerColumns in projectWorkbook.OfferColumns)
            {
                Excel.Range cellCostAddress = ws.Cells[firstRow, offerColumns.ColTotalCostPerUnitOffer];
                Excel.Range cellWorksAddress = ws.Cells[firstRow, offerColumns.ColCostWorksPerUnitOffer];
                Excel.Range cellMaterialsAddress = ws.Cells[firstRow, offerColumns.ColCostMaterialsPerUnitOffer];
                Excel.Range cellAmountAddress = ws.Cells[firstRow, offerColumns.ColCountOffer];

                //Аргументы функции
                argumentsCommonCost += "," + cellCostAddress.Address[RowAbsolute: false, ColumnAbsolute: true];
                argumentsCommonWorks += "," + cellWorksAddress.Address[RowAbsolute: false, ColumnAbsolute: true];
                argumentsCommonMaterials += "," + cellMaterialsAddress.Address[RowAbsolute: false, ColumnAbsolute: true];
                argumentsAmount += "," + cellAmountAddress.Address[RowAbsolute: false, ColumnAbsolute: true];
            }


            // Формалы среденее медиана
            CellAvgAmount.Formula = $"=AVERAGE({argumentsAmount})";
            CellAvgCostMaterials.Formula = $"=AVERAGE({argumentsCommonMaterials})";
            CellAvgCostWorks.Formula = $"=AVERAGE({argumentsCommonWorks})";
            CellAvgTotalCost.Formula = $"=AVERAGE({argumentsCommonCost})";
          
            CellMedianCostMaterials.Formula = $"=MEDIAN({argumentsCommonMaterials})";
            CellMedianCostWorks.Formula = $"=MEDIAN({argumentsCommonWorks})";
            CellMedianTotalCost.Formula = $"=MEDIAN({argumentsCommonCost})"; ///$"= {addressMedianCostMaterials} + " +
                                                                             //  $"{addressMedianCostWorks}"; 

            /// Для каждого диапазона КП
            foreach (OfferColumns offerColumns in projectWorkbook.OfferColumns)
            {
                string formulaDeviationCost = "";
                string formulaDviationWorks = "";
                string formulaDviationMaterials = "";
                string formulaDviationAmount = "";

                ///Отклонение по стоимости // Адрес ячейки
                Excel.Range CellOfferName = ws.Cells[firstRow, offerColumns.ColNameOffer];
                Excel.Range CellOfferAmount = ws.Cells[firstRow, offerColumns.ColCountOffer];
                Excel.Range CellOfferCost = ws.Cells[firstRow, offerColumns.ColTotalCostPerUnitOffer];
                Excel.Range CellOfferWorks = ws.Cells[firstRow, offerColumns.ColCostWorksPerUnitOffer];
                Excel.Range CellOfferMaterials = ws.Cells[firstRow, offerColumns.ColCostMaterialsPerUnitOffer];

                Excel.Range CellOfferDeviationVolume = ws.Cells[firstRow, offerColumns.ColDeviationVolume];
                Excel.Range CellOfferDeviationCost = ws.Cells[firstRow, offerColumns.ColDeviationCost];
                Excel.Range CellOfferDeviationWorks = ws.Cells[firstRow, offerColumns.ColDeviationWorks];
                Excel.Range CellOfferDeviationMaterials = ws.Cells[firstRow, offerColumns.ColDeviationMaterials];


                string AddressOfferName = CellOfferName.Address[RowAbsolute: false, ColumnAbsolute: true];
                string AddressOfferAmount = CellOfferAmount.Address[RowAbsolute: false, ColumnAbsolute: true];
                string AddressOfferCost = CellOfferCost.Address[RowAbsolute: false, ColumnAbsolute: true];
                string AddressOfferWorks = CellOfferWorks.Address[RowAbsolute: false, ColumnAbsolute: true];
                string AddressOfferMaterials = CellOfferMaterials.Address[RowAbsolute: false, ColumnAbsolute: true];

                string AddressOfferDeviationCost = CellOfferDeviationCost.Address[RowAbsolute: false, ColumnAbsolute: true];
                string AddressDeviationVolume = CellOfferDeviationVolume.Address[RowAbsolute: false, ColumnAbsolute: true];
                string AddressDeviationWorks = CellOfferDeviationWorks.Address[RowAbsolute: false, ColumnAbsolute: true];

                if (Properties.Settings.Default.AnalysisFormulaCost == (byte)FormulaAnalysis.DeviationBasis)
                {
                    // Отклонение  от базовой оценки
                    //Отклонение по объемам                    
                    formulaDviationAmount = $"=IFERROR({addressAmount}/{AddressOfferAmount}-1,\"#НД\")";
                    formulaDeviationCost = $"=IFERROR(IF({addressBasisCostPerUnit}<>0," +
                                           $"{AddressOfferCost}/{addressBasisCostPerUnit}-1,0),\"#НД\")";
                    // по стоимости работ
                    formulaDviationWorks =
                        $"=IFERROR(IF({addressBasisWorks}<>0," +
                       $"{AddressOfferWorks }/{addressBasisWorks}-1,\"Отс-ет ст-ть работ\"),\"#НД\")";
                    // по стоимости материалов
                    formulaDviationMaterials =
                        $"=IFERROR(IF({addressBasisMaterials}<>0," +
                       $"{AddressOfferMaterials}/{addressBasisMaterials}-1,\"Отс-ет ст-ть мат.\"),\"#НД\")";
                }
                else if (Properties.Settings.Default.AnalysisFormulaCost == (byte)FormulaAnalysis.Avarage)
                {
                    // Отклонение от среднего
                    // по стоимости
                    formulaDeviationCost = $"=IFERROR({AddressOfferCost}/{addressAvgTotalCost}-1,\"#НД\")";
                    // по стоимости работ
                    formulaDviationWorks = $"=IFERROR(IF({addressAvgCostWorks}<>0," +
                       $"{AddressOfferWorks }/ {addressAvgCostWorks}-1,\"Отс-ет ст-ть работ\"),\"#НД\")";
                    // по стоимости материалов
                    formulaDviationMaterials = $"=IFERROR(IF({addressAvgCostMaterials}<>0," +
                       $"{AddressOfferMaterials }/ {addressAvgCostMaterials} -1 , \"Отс-ет ст-ть мат.\"),\"#НД\")";


                }
                else if (Properties.Settings.Default.AnalysisFormulaCost == (byte)FormulaAnalysis.Median)
                {
                    // Отклонение от медианы
                    // по стоимости
                    formulaDeviationCost = $"=IFERROR({AddressOfferCost }/{addressMedianTotalCost}-1,\"#НД\")";
                    // по стоимости работ
                    formulaDviationWorks = $"=IFERROR(IF({addressMedianCostWorks}<>0," +
                      $"{AddressOfferWorks }/ {addressMedianCostWorks} -1 ,\"Отс-ет ст-ть работ\"),\"#НД\")";
                    // по стоимости материалов
                    formulaDviationMaterials = $"=IFERROR(IF({addressMedianCostMaterials}<>0," +
                      $"{AddressOfferMaterials }/ {addressMedianCostMaterials} -1 ,\"Отс-ет ст-ть мат.\"),\"#НД\")";
                }

                //-----------------------------------------------------
                //Отклонение по объемам 
                if (Properties.Settings.Default.AnalysisFormulaCount == (byte)FormulaAnalysis.DeviationBasis)
                {
                    // Отклонение  от базовой оценки                                      
                    formulaDviationAmount = $"=IFERROR({AddressOfferAmount}/{addressAmount}-1,\"#НД\")";
                }
                else if (Properties.Settings.Default.AnalysisFormulaCount == (byte)FormulaAnalysis.Avarage)
                {
                    // Отклонение от среднего
                    formulaDviationAmount = $"=IFERROR({AddressOfferAmount}/{addressAvgAmount}-1,\"#НД\")";
                }
                CellOfferDeviationCost.Formula = formulaDeviationCost;
                CellOfferDeviationWorks.Formula = formulaDviationWorks;
                CellOfferDeviationMaterials.Formula = formulaDviationMaterials;
                //Отклонение по объемам
                CellOfferDeviationVolume.Formula = formulaDviationAmount;

                //Наименование вида работ
                Excel.Range cellChekName = ws.Cells[firstRow, offerColumns.ColStartOfferComments];
                string AddressChekName = cellChekName.Address[RowAbsolute: false, ColumnAbsolute: true];
                cellChekName.Formula = $"=${letterName}{firstRow}={AddressOfferName}";
                //Комментарии Спектрум к описанию работ
                ws.Cells[firstRow, offerColumns.ColCommentsDescriptionWorks].Formula = $"=IF({AddressChekName}=TRUE,\".\",Комментарии!$A$2)";

                //Комментарии Спектрум к объемам работ
                ws.Cells[firstRow, offerColumns.ColCommentsVolume].Formula = $"=IF({AddressDeviationVolume}=\"#НД\",\"#НД\", IF({AddressDeviationVolume}>{top}%,Комментарии!$A$5,IF({AddressDeviationVolume}<{bottom}%,Комментарии!$A$6,\".\")))";

                //Комментарии к строкам "0"
                ws.Cells[firstRow, offerColumns.ColComments].Formula =
                           $"=IF({AddressOfferDeviationCost}=-1,\"Указать стоимость единичной расценки и посчитать итог\",\".\")";

                //Комментарии Спектрум к стоимости работ
                ws.Cells[firstRow, offerColumns.ColCommentsCostWorks].Formula =
                    $"=IF({AddressOfferDeviationCost}=\"#НД\",\"#НД\", IF({AddressOfferDeviationCost}>{top}%,Комментарии!$A$9,IF({AddressOfferDeviationCost}<{bottom}%,Комментарии!$A$10,\".\")))";

                // Протянуть формулы до конца листа
                Excel.Range rng = ws.Range[ws.Cells[firstRow, offerColumns.ColStartOfferComments],
                                                    ws.Cells[firstRow, offerColumns.ColComments]];
                if (lastRow > firstRow)
                {
                    Excel.Range destination = ws.Range[ws.Cells[firstRow, offerColumns.ColStartOfferComments], ws.Cells[lastRow, offerColumns.ColComments]];
                    rng.AutoFill(destination);
                    destination.Interior.Color = Color.FromArgb(232, 242, 238);
                    destination.Columns[3].NumberFormat = "0%";
                    destination.Columns[5].NumberFormat = "0%";
                    destination.Columns[7].NumberFormat = "0%";
                    destination.Columns[8].NumberFormat = "0%";

                    Excel.Range rangeAvgAnalysis = ws.Range[CellAvgAmount, ws.Cells[lastRow, CellMedianTotalCost.Column]];
                    rangeAvgAnalysis.Rows[1].AutoFill(rangeAvgAnalysis);
                    rangeAvgAnalysis.NumberFormat = "#,##0.00";
                }
            }
        }

        private void BtnClearFormateContions_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rng = _app.Selection;
            // Удалить условное форматирование            
            rng.FormatConditions.Delete();
        }

        private void BtnCol_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range rng = _app.Selection;
            // Удалить условное форматирование
            rng.FormatConditions.Delete();
            ConditonsFormatManager formatManager = new ConditonsFormatManager();
            List<ConditionFormat> conditions = formatManager.ListConditionFormats.FindAll(a => a.ColumnName == "Выделение");
            conditions.ForEach(x => x.SetCondition(rng));
        }

        private void GroupColumns()
        {
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            Project project = projectManager.ActiveProject;
            Excel.Workbook wb = _app.ActiveWorkbook;
            Excel.Worksheet sh = ExcelHelper.GetSheet(wb, project.AnalysisSheetName);
            new ListAnalysis(sh, project).GroupColumn();
        }

        /// <summary>
        /// Группировка строк
        /// </summary>
        private void GroupRows()
        {
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.SetMainBarVolum(1);
            try
            {
                ExcelAcselerate(true);
                ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
                Project project = projectManager.ActiveProject;
                Excel.Workbook wb = _app.ActiveWorkbook;
                Excel.Worksheet sh = ExcelHelper.GetSheet(wb, project.AnalysisSheetName);

                string letterLevel = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
                pb.Writeline("Группировка строк");
                ExcelHelper.Group(sh, pb, letterLevel);
            }
            catch (Exception ex)
            {
                string message = $"Ошибка:{ex.Message }";
                if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                pb.ClearMainBar();
                pb.CloseFrm();
                ExcelAcselerate(false);
            }
        }

        private void BtnGroupColumns_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return; // ячейка редактируется
            try
            {
                ExcelAcselerate(true);
                GroupColumns();
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
        }

        private void BtnGroupRows_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return; // ячейка редактируется
            try
            {
                ExcelAcselerate(true);
                GroupRows();
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
        }

        private void BtnUngroupColumns_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return;
            try
            {
                ExcelHelper.UnGroupColumns(_app.ActiveSheet);
            }
            catch (Exception ex)
            {
                string message = $"Ошибка:{ex.Message }";
                if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnUngroupRows_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelper.UnGroupRows(_app.ActiveSheet);
        }

        private void BtnFormatNumber_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
                ProjectManager.Project project = projectManager.ActiveProject;
                Excel.Worksheet ws = ExcelHelper.GetSheet(wb, project.AnalysisSheetName);

                string letterAmount = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Amount]).ColumnSymbol;
                string letterMaterialPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostMaterialsPerUnit]).ColumnSymbol;
                string letterMaterialTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostMaterialsTotal]).ColumnSymbol;
                string letterWorkPerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostWorksPerUnit]).ColumnSymbol;
                string letterWorkTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostWorksTotal]).ColumnSymbol;
                string letterPricePerUnit = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotalPerUnit]).ColumnSymbol;
                string letterTotal = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.CostTotal]).ColumnSymbol;
                string letterComment = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Comment]).ColumnSymbol;

                List<(string, string)> columns_format = ProjectWorkbook.GetFormatColumns(ws);
                ExcelHelper.SetNumberFormat(ws, project.RowStart, columns_format.ToArray());
                ExcelHelper.SetNumberFormat(ws, project.RowStart, letterAmount);
                ExcelHelper.SetNumberFormat(ws, project.RowStart, letterMaterialPerUnit);
                ExcelHelper.SetNumberFormat(ws, project.RowStart, letterMaterialTotal);
                ExcelHelper.SetNumberFormat(ws, project.RowStart, letterWorkPerUnit);
                ExcelHelper.SetNumberFormat(ws, project.RowStart, letterWorkTotal);
                ExcelHelper.SetNumberFormat(ws, project.RowStart, letterPricePerUnit);
                ExcelHelper.SetNumberFormat(ws, project.RowStart, letterTotal);
            }
            catch (Exception ex)
            {
                string message = $"Ошибка:{ex.Message }";
                if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSetFormul_Click(object sender, RibbonControlEventArgs e)
        {
            FormSettingFormuls form = new FormSettingFormuls();
            if (form.ShowDialog() == DialogResult.OK)
            {
                SetAnalysis();
            }
        }

        private void BtnDataFilter_Click(object sender, RibbonControlEventArgs e)
        {
            if (ExcelHelper.IsEditing()) return;
            try
            {
                SetDataFilter();
            }
            catch (Exception ex)
            {
                string message = $"Ошибка:{ex.Message }";
                if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                MessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetDataFilter()
        {
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            ProjectManager.Project project = projectManager.ActiveProject;
            SetFilter(project);
        }

        private void SetFilter(Project project)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = ExcelHelper.GetSheet(wb, project.AnalysisSheetName);

            int lastRow = ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1;
            int lastCol = ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1;

            Excel.Range rng = ws.Range[ws.Cells[project.RowStart - 1, 1], ws.Cells[lastRow, lastCol]];
            rng.AutoFilter(1);
        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = _app.ActiveWorkbook;
            string without = "БЕЗ НДС"; //РУБ БЕЗ НДС
            string with = "C НДС"; //РУБ БЕЗ НДС

            string find = default;
            string replacement = default;

            if (comboBoxLvlCost.Text  == "Без НДС")
            {
                find = with;
                replacement = without;
            }
            else if (comboBoxLvlCost.Text == "С НДС")
            {
                find = without;
                replacement = with;
            }

            foreach (Excel.Worksheet sh in workbook.Worksheets)
            {
                sh.UsedRange.Replace(What: find, Replacement: replacement, MatchCase: false);
            }
        }
    }
}
