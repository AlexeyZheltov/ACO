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
using ACO.ProjectBook;
using System.Drawing;

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
        /// Загрузка КП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLoadKP_Click(object sender, RibbonControlEventArgs e)
        {
            string[] files = GetFiles();
            if ((files?.Length ?? 0) < 1) return;

            string offerSettingsName = GetOfferSettings();
            if (string.IsNullOrEmpty(offerSettingsName)) return;

            ExcelHelpers.ExcelFile.Init();
            ExcelAcselerate(true);

            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            pb.CloseForm += () => { pb = null; };
            pb.Show(new AddinWindow(Globals.ThisAddIn));

            pb.SetMainBarVolum(files.Length);
            ExcelHelpers.ExcelFile excelBook = new ExcelHelpers.ExcelFile();

            pb.Writeline("Инициализацция диспетчера проектов.");
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            foreach (string fileName in files)
            {
                try
                {
                    if (pb.IsAborted) throw new AddInException("Процесс остановлен");
                    pb.MainBarTick(fileName);
                    pb.Writeline("Открытие файла");
                    excelBook.Open(fileName);
                    pb.Writeline("Инициализация загрузчика");
                    OfferWriter offerWriter = new OfferWriter(excelBook);
                    pb.Writeline("Заполнение листа Анализ\n");

                    //await Task.Run(() =>
                    //{
                    //});
                    offerWriter.Print(pb, offerSettingsName);
                    pb.CloseFrm();
                }
                catch (AddInException ex)
                {
                    string message = $"Ошибка:{ex.Message }";
                    if (ex.InnerException != null) message += $"{ex.InnerException.Message}";
                    pb.Writeline(message);
                }
                finally
                {
                    if (pb?.IsAborted ?? false)
                    {
                        pb.ClearMainBar();
                        pb.ClearSubBar();
                        pb.IsAborted = false;
                        MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    excelBook.Close();
                    ExcelHelpers.ExcelFile.Finish();
                    ExcelAcselerate(false);
                }
            }
        }

        private void BtnSpectrum_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelpers.ExcelFile.Init();
            ExcelAcselerate(true);
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            try
            {
                string file = GetFile();
                if (!File.Exists(file)) { return; }
                pb.CloseForm += () => { pb = null; };
                pb.Show(new AddinWindow(Globals.ThisAddIn));
                PrintSpectrum(pb, file);
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
                ExcelHelpers.ExcelFile.Finish();
                ExcelAcselerate(false);
            }
        }
        private async void PrintSpectrum(IProgressBarWithLogUI pb, string file)
        {
            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            pb.SetMainBarVolum(1);
            pb.MainBarTick(file);
            if (pb.IsAborted) throw new AddInException("Процесс остановлен");
            await Task.Run(() =>
            {
                OfferWriter offerWriter = new OfferWriter(file);
                offerWriter.PrintSpectrum(pb);

                pb.CloseFrm();
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

        private void BtnUpdateFormuls_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelAcselerate(true);
            try
            {
                UpdateFormuls();
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
                ExcelAcselerate(false);
            }
        }

        /// <summary>
        /// Формулы, окраска уровней
        /// </summary>
        private async void UpdateFormuls()
        {
            IProgressBarWithLogUI pb = new ProgressBarWithLog();

            pb.CloseForm += () => { pb = null; };
            pb.Show(new AddinWindow(Globals.ThisAddIn));
            pb.ClearMainBar();
            pb.ClearSubBar();
            pb.SetMainBarVolum(5);

            pb.MainBarTick("Подготвка");
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            ProjectManager.ProjectManager projectManager = new ProjectManager.ProjectManager();
            ProjectManager.Project project = projectManager.ActiveProject;
            
            Excel.Worksheet ws = ExcelHelper.GetSheet(wb, project.AnalysisSheetName); 
            Excel.Worksheet pws = ExcelHelper.GetSheet(wb, "Палитра"); 
            ExcelHelper.UnGroup(ws);
            HItem root = new HItem();
            string letterLevel = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;

            foreach (var (Row, Level) in ExcelReader.ReadSourceItems(ws, letterLevel, project.RowStart))
                root.Add(new HItem()
                {
                    Level = Level,
                    Row = Row
                });
            PbAbortedStopProcess(pb);

            pb.MainBarTick("Подготовка списка");
            pb.SetSubBarVolume(root.AllCount());
            await Task.Run(() =>
            {
                root.Numeric(new Numberer(), pb); 
            });
            string letterNumber = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Number]).ColumnSymbol;
            pb.ClearSubBar();
            pb.SetSubBarVolume(root.AllCount());
            await Task.Run(() =>
            {
                ExcelHelper.Write(ws, root, pb, letterNumber);
            });

            PbAbortedStopProcess(pb);

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
            await Task.Run(() =>
            {
                //два
                ExcelHelper.SetFormulas(ws, mappin, root, pb); //Прогресс бар только для отмены
            });
            // Обновление формул КП
            HashSet<int> columnsAmount = GetNumbersCoumnsOfCount(ws);
            foreach (int col in columnsAmount)
            {
                ExcelHelper.SetFormulas(ws, mappin.Shift(ws, col), root, pb);
            }

            PbAbortedStopProcess(pb);
            pb.MainBarTick("Форматирование списка");
            //три
            pb.ClearSubBar();

            var pallet = ExcelReader.ReadPallet(pws);
            int count = ws.UsedRange.Rows.Count;//root.AllCount();
            pb.SetSubBarVolume(count);
            //letterLevel = project.Columns.Find(x => x.Name == Project.ColumnsNames[StaticColumns.Level]).ColumnSymbol;
            List<(string, string)> colored_columns = GetColredcolumns(ws);
            colored_columns.Add(("A", letterComment));
            (string, string)[] columns = colored_columns.ToArray();

            //четыре
            await Task.Run(() =>
            {
                ExcelHelper.Repaint(ws, pallet, project.RowStart, letterLevel, pb, columns); 
            });

            PbAbortedStopProcess(pb);
            pb.MainBarTick("Группировка списка");
            pb.ClearSubBar();
            await Task.Run(() =>
            {
                ExcelHelper.Group(ws, pb, letterLevel); //Этот метод сам установит Max для прогрессбара
            });
            pb.ClearMainBar();
            pb.CloseFrm();
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
        ///  Определить столбцы для окрашивания
        /// </summary>
        /// <param name="ws"></param>
        /// <returns></returns>
        private List<(string, string)> GetColredcolumns(Excel.Worksheet ws)
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
        ///  Окраска ячеек 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnColorComments_Click(object sender, RibbonControlEventArgs e)
        {
            ProjectWorkbook projectWorkbook = new ProjectWorkbook();
            int lastRow = projectWorkbook.AnalisysSheet.UsedRange.Row + projectWorkbook.AnalisysSheet.UsedRange.Rows.Count;
            int startRow = projectWorkbook.GetFirstRow();
            for (int row = startRow; row <= lastRow; row++)
            {
                foreach (OfferAddress offeraddress in projectWorkbook.OfferAddress)
                {
                    ColorCell(projectWorkbook.AnalisysSheet.Cells[row, offeraddress.ColPercentWorks]);
                    ColorCell(projectWorkbook.AnalisysSheet.Cells[row, offeraddress.ColPercentMaterial]);
                    ColorCell(projectWorkbook.AnalisysSheet.Cells[row, offeraddress.ColPercentTotal]);
                }
            }
        }

        private void ColorCell(Excel.Range cell)
        {
            string text = cell.Value?.ToString() ?? "";
            if (text != "#НД" || text != "")
            {
                double percent = double.TryParse(text, out double pct) ? pct : 0;
                if (percent > 0.15)
                {//Красный  >0.15
                    cell.Interior.Color = Color.FromArgb(255, 0, 0);
                    cell.Font.Color = Color.FromArgb(255, 255, 255);
                }
                else if (percent < -0.15)
                {// Желтый 
                    cell.Interior.Color = Color.FromArgb(242, 255, 0);
                    cell.Font.Color = Color.FromArgb(242, 0, 0);
                }
                else if (percent < -0.05 && percent > -0.15 || percent > 0.05 && percent < 0.15)
                {// Светло желтый
                    cell.Interior.Color = Color.FromArgb(252, 250, 104);
                    cell.Font.Color = Color.FromArgb(0, 0, 0);
                }
                else if (percent < -0.05 && percent > -0.15 || percent > 0.05 && percent < 0.15)
                {//  0.05 < percent < -0.05,  0.15 < percent > -0.15 
                    /// Светло фиолетовый
                    cell.Interior.Color = Color.FromArgb(255, 176, 197);
                    cell.Font.Color = Color.FromArgb(125, 0, 33);
                }
                else
                {
                    cell.Interior.Color = Color.FromArgb(255, 255, 255);
                    cell.Font.Color = Color.FromArgb(0, 0, 0);
                }

            }
        }



        /// <summary>
        ///  Запись формул на уровень
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLoadLvl11_Click(object sender, RibbonControlEventArgs e)
        {
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
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
        }
        /// <summary>
        ///  Запись формул на уровень 12
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLoadLvl12_Click(object sender, RibbonControlEventArgs e)
        {
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            ExcelAcselerate(true);
            try
            {
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
        }


        private void BtnUpdateLvl11_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range cell = Globals.ThisAddIn.Application.ActiveCell;
            string text = ExcelHelper.GetText(cell);
            MessageBox.Show(text);
        }

        private void BtnUpdateLvl12_Click(object sender, RibbonControlEventArgs e)
        {
            IProgressBarWithLogUI pb = new ProgressBarWithLog();
            ExcelAcselerate(true);
            try
            {
                pb.Show();
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
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
        }
    }
}
