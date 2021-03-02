using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spectrum.SpLoader.XMLSetting;
using Spectrum.SpLoader.ExcelHelpers;
using System.IO;
using OCD = Ookii.Dialogs.WinForms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Spectrum.SpLoader
{
    public partial class MainForm : Form
    {
        static readonly char[] _allowLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
        bool _mappingChanged = false;
        string _collectFolderPath;
        string[] _collectFiles;
        IProgressBarWithLogUI _pb = null;

        public MainForm()
        {
            InitializeComponent();
        }

        private async void CreateVORButton_Click(object sender, EventArgs e)
        {
            if (!Validator.IsValid())
            {
                MessageBox.Show("Исправте все ошибки.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if ((_collectFiles?.Length ?? 0) == 0)
            {
                MessageBox.Show("Нет выбранных файлов", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_mappingChanged)
            {
                MessageBox.Show("Маппинг изменен. Для продолжения сохраните маппинг.", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (_pb is null)
            {
                _pb = new ProgressBarWithLog();
                _pb.CloseForm += () => { _pb = null; };
                _pb.Show();
            }

            _pb.ClearMainBar();
            _pb.ClearSubBar();
            _pb.SetMainBarVolum(8);
            _pb.MainBarTick("Подключение к Excel");

            await Task.Run(() =>
            {
                ExcelFile.Init();
            });

            ExcelFile efile = new ExcelFile();
            if (!efile.Open(SettingManager.OmniPath))
            {
                MessageBox.Show("Не удалось открыть Omni файл", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Excel.Worksheet ws = efile.GetSheet("Table");

            if(ws == null)
            {
                MessageBox.Show($"В файле {efile.WorkBook.Name} отсутствует таблица Table", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExcelFile.Acselerate(true);
            _pb.MainBarTick("Обработка Omni файла");

            Dictionary<string, string[]> omni = null;
            await Task.Run(() =>
            {
                omni = ExcelReader.ReadOmni(ws, _pb);
            });

            if (_pb.IsAborted)
            {
                efile.Close();
                _pb.ClearMainBar();
                _pb.ClearSubBar();
                _pb.IsAborted = false;
                ExcelFile.Acselerate(false);
                ExcelFile.Finish();
                MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            efile.Close();
            
            HierarchyDictionary root = new HierarchyDictionary();
            string[] OmniKeys = omni.Keys.ToArray();
            Mapping CurrentMapping = MappingManager.Current;

            int temp = 1;
            _pb.MainBarTick("", 1);
            await Task.Run(() =>
            {
                foreach (var file in _collectFiles)
                {
                    if (_pb.IsAborted) break;
                    if (efile.Open(file))
                    {
                        _pb.MainBarTick($"Обработка файла {temp++} из {_collectFiles.Length}", 0);
                        _pb.ClearSubBar();
                        Excel.Worksheet worksheet = efile.WorkBook.Sheets[1];
                        _pb.SetSubBarVolume(worksheet.UsedRange.Rows.Count);
                        string topLevelName = ExcelReader.GetTopLevel(worksheet);
                        foreach (TargetItem item in ExcelReader.ReadSourceItems(worksheet, CurrentMapping, OmniKeys))
                        {
                            if (_pb.IsAborted) break;
                            _pb.SubBarTick();
                            if (item is null) continue;
                            Stack<string> stack = omni[item.OmniClass].ToStack();
                            stack.Push(topLevelName);
                            root.Add(new Package
                            {
                                Item = item,
                                Path = stack
                            });
                        }
                        efile.Close();
                    }
                }
            });

            if (_pb.IsAborted)
            {
                efile.Close();
                _pb.ClearMainBar();
                _pb.ClearSubBar();
                _pb.IsAborted = false;
                ExcelFile.Acselerate(false);
                ExcelFile.Finish();
                MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _pb.MainBarTick("Нумерация");
            _pb.ClearSubBar();
            temp = root.AllCount();
            _pb.SetSubBarVolume(temp);
            await Task.Run(() =>
            {
                root.Numeric(new Numberer(), _pb);
            });

            if (_pb.IsAborted)
            {
                _pb.ClearMainBar();
                _pb.ClearSubBar();
                _pb.IsAborted = false;
                ExcelFile.Acselerate(false);
                ExcelFile.Finish();
                MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _pb.MainBarTick("Выгрузка данных в итоговый файл");
            _pb.ClearSubBar();
            _pb.SetSubBarVolume(temp);
            efile.Open(SettingManager.TemplatePath);

            await Task.Run(() =>
            {
                ExcelHelper.WriteResult(efile.GetSheet("Рсч-П"), root, _pb);
            });

            if (_pb.IsAborted)
            {
                efile.Close();
                _pb.ClearMainBar();
                _pb.ClearSubBar();
                _pb.IsAborted = false;
                ExcelFile.Acselerate(false);
                ExcelFile.Finish();
                MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _pb.MainBarTick("Группировка данных");
            _pb.ClearSubBar();

            await Task.Run(() =>
            {
                ExcelHelper.Group(efile.GetSheet("Рсч-П"), _pb);
            });

            if (_pb.IsAborted)
            {
                efile.Close();
                _pb.ClearMainBar();
                _pb.ClearSubBar();
                _pb.IsAborted = false;
                ExcelFile.Acselerate(false);
                ExcelFile.Finish();
                MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _pb.MainBarTick("Покраска таблицы");
            _pb.ClearSubBar();

            await Task.Run(() =>
            {
                ExcelHelper.Repaint(efile.GetSheet("Рсч-П"), ExcelReader.ReadPallet(efile.GetSheet("Палитра")), _pb);
            });

            if (_pb.IsAborted)
            {
                efile.Close();
                _pb.ClearMainBar();
                _pb.ClearSubBar();
                _pb.IsAborted = false;
                ExcelFile.Acselerate(false);
                ExcelFile.Finish();
                MessageBox.Show("Выполнение было прервано", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _pb.MainBarTick("Сохранение файла");
            _pb.ClearSubBar();

            await Task.Run(() =>
            {
                efile.WorkBook.SaveAs(Path.Combine(_collectFolderPath, "CollectedData.xlsm"));
            });
            
            efile.Close();
            _pb.ClearMainBar();
            ExcelFile.Acselerate(false);
            ExcelFile.Finish();

            MessageBox.Show("Сбор данных завершен", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            Validator.SetErrorProvider(errorProvider);
            SettingManager.Load();

            Validator.Add(TemplatePathTextBox, Validator.Type.Path);
            TemplatePathTextBox.BackColor = Color.White;
            TemplatePathTextBox.SetText(SettingManager.TemplatePath);

            Validator.Add(OmniPathTextBox, Validator.Type.Path);
            OmniPathTextBox.BackColor = Color.White;
            OmniPathTextBox.SetText(SettingManager.OmniPath);

            MappingManager.Load();
            MappingComboBox.Items.Clear();
            MappingComboBox.Items.AddRange(MappingManager.GetMappingList());

            if (MappingManager.Current is Mapping mapping)
            {
                MappingComboBox.SelectedItem = mapping.Name;
                MappingToTextBoxs(mapping);
            }

            foreach (Control control in ColumnsGroupBox.Controls)
                if (control is TextBox textBox)
                    Validator.Add(textBox, Validator.Type.Column);

            Validator.Validate();

            if (MappingComboBox.Text == "")
                errorProvider.SetError(MappingComboBox, "Не выбран маппинг");
            else
                errorProvider.SetError(MappingComboBox, "");
        }

        private void MappingToTextBoxs(Mapping mapping)
        {
            TypeColumnTextBox.Text = mapping.Type;
            FormatColumnTextBox.Text = mapping.Format;
            MarkingColumnTextBox.Text = mapping.Marking;
            WorkNameColumnTextBox.Text = mapping.WorkName;
            AmountColumnTextBox.Text = mapping.Amount;
            OmniClassColumnTextBox.Text = mapping.Omni;
            MakerColumnTextBox.Text = mapping.Maker;
            NoteColumnTextBox.Text = mapping.Note;
            MaterialColumnTextBox.Text = mapping.Material;
            ArticleColumnTextBox.Text = mapping.Article;
            UnitColumnTextBox.Text = mapping.Unit;
            _mappingChanged = false;
        }

        private void TextBoxsToMapping(Mapping mapping)
        {
            mapping.Type = TypeColumnTextBox.Text;
            mapping.Format = FormatColumnTextBox.Text;
            mapping.Marking = MarkingColumnTextBox.Text;
            mapping.WorkName = WorkNameColumnTextBox.Text;
            mapping.Amount = AmountColumnTextBox.Text;
            mapping.Omni = OmniClassColumnTextBox.Text;
            mapping.Maker = MakerColumnTextBox.Text;
            mapping.Note = NoteColumnTextBox.Text;
            mapping.Material = MaterialColumnTextBox.Text;
            mapping.Article = ArticleColumnTextBox.Text;
            mapping.Unit = UnitColumnTextBox.Text;
        }

        private void TemplatePathButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                CheckFileExists = true,
                CheckPathExists = true,
                Filter = "Файл шаблона|*.xlsm",
                FilterIndex = 0,
                Multiselect = false,
                RestoreDirectory = true,
                Title = "Выбор файла шаблона"
            };

            if(dialog.ShowDialog() == DialogResult.OK)
            {
                TemplatePathTextBox.SetText(dialog.FileName);
                SettingManager.Load();
                SettingManager.TemplatePath = dialog.FileName;
                SettingManager.Save();
                Validator.Validate();
            }
        }

        private void OmniPathButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog()
            {
                CheckFileExists = true,
                CheckPathExists = true,
                Filter = "Файл OmniClass|*.xlsx",
                FilterIndex = 0,
                Multiselect = false,
                RestoreDirectory = true,
                Title = "Выбор файла OmniClass"
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                OmniPathTextBox.SetText(dialog.FileName);
                SettingManager.Load();
                SettingManager.OmniPath = dialog.FileName;
                SettingManager.Save();
                Validator.Validate();
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            errorProvider.Dispose();
        }

        private void AddMappingButton_Click(object sender, EventArgs e)
        {
            using(InputBox dialog = new InputBox()
            {
                Title = "Новый маппинг",
                Caption = "Введите имя новго маппинга"
            })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    MappingManager.Add(dialog.Value);
                    ReloadMappingList();
                }
            }
        }

        private void RenameMappingButton_Click(object sender, EventArgs e)
        {
            if(MappingManager.Current is Mapping mapping)
            {
                using (InputBox dialog = new InputBox()
                {
                    Title = "Переименовывание маппинга",
                    Caption = "Введите новое имя маппинга"
                })
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        MappingManager.Rename(mapping.Name, dialog.Value);
                        ReloadMappingList();
                    }
                }
            }
        }

        private void ReloadMappingList()
        {
            MappingComboBox.Items.Clear();
            MappingComboBox.Items.AddRange(MappingManager.GetMappingList());
            MappingComboBox.SelectedItem = MappingManager.Current.Name;
            MappingToTextBoxs(MappingManager.Current);
            Validator.Validate();
        }

        private void SaveMappingButton_Click(object sender, EventArgs e)
        {
            if (Validator.IsValidMapping() && MappingManager.Current is Mapping mapping)
            {
                TextBoxsToMapping(mapping);
                MappingManager.Save();
                _mappingChanged = false;
            }
            else
                MessageBox.Show("Исправте все ошибки маппинга", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DeleteMappingButton_Click(object sender, EventArgs e)
        {
            //Загружать сохранять и ваще че делать
            MappingManager.Delete();
        }

        private void ColumnTextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char keyChar = e.KeyChar;

            if (Char.IsControl(keyChar))
                return;

            keyChar = Char.ToUpper(keyChar);

            if ((sender as TextBox).TextLength == 3 || !_allowLetters.Contains(keyChar))
            {
                e.Handled = true;
                return;
            }

            e.KeyChar = keyChar;
        }

        private void ColumnTextBox_TextChanged(object sender, EventArgs e)
        {
            Validator.Validate();
            _mappingChanged = true;
        }

        private void MappingComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            MappingManager.Select(MappingComboBox.Text);
            MappingToTextBoxs(MappingManager.Current);
            errorProvider.SetError(MappingComboBox, "");
        }

        private void SelectFilesButton_Click(object sender, EventArgs e)
        {
            //Выбор папки с файлами
            OCD.VistaFolderBrowserDialog dialog = new OCD.VistaFolderBrowserDialog()
            {
                Description = "Выбор папки",
                UseDescriptionForTitle = true
            };

            if(dialog.ShowDialog() == DialogResult.OK)
            {
                _collectFolderPath = dialog.SelectedPath;
                _collectFiles = (from path in Directory.GetFiles(_collectFolderPath, "*.xls*")
                                 where !Path.GetFileName(path).StartsWith("~")
                                 select path).ToArray();
                SelectedFilesCountLable.Text = $"Выбранно файлов: {_collectFiles.Length} шт.";
            }
        }
    }

    static class TextBoxExtension
    {
        public static void SetText(this TextBox textBox, string text)
        {
            textBox.Text = text;
            textBox.SelectionStart = textBox.TextLength;
        } 
    }
}
