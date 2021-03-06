﻿using ACO.ExcelHelpers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ACO.ProjectManager
{
    public partial class FormManager : Form
    {
        private readonly Excel.Application _app = Globals.ThisAddIn.Application;
        private ProjectManager _projectManager;
        private List<ColumnMapping> _mappingColumns;
        public FormManager()
        {
            InitializeComponent();
            TableColumns.ReadOnly = false;
            TableProjects.ReadOnly = false;
            TableProjects.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            LoadData();
        }
        private void LoadData()
        {
            LoadProjects();
            LoadColumns();
            LoadRangeValues();
        }

        /// <summary>
        ///  Данные таблицы проектов
        /// </summary>
        private void LoadProjects()
        {
            _projectManager = new ProjectManager();
            if (_projectManager.Projects.Count > 0)
            {
                UpdateTableProject();
                TableProjects.Columns[0].HeaderText = "Текущий";
                TableProjects.Columns[1].HeaderText = "Проект";
                TableProjects.Columns[2].HeaderText = "Путь";
                TableProjects.Columns[2].Visible = false;
                TableProjects.Columns[3].Visible = false;
                TableProjects.Columns[4].Visible = false;
                TableProjects.Columns[0].Width = 70;              
                TableProjects.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                for (int i = 0; i < TableProjects.Rows.Count; i++)
                {
                    if (_projectManager.Projects[i].Name == Properties.Settings.Default.ActiveProjectName)
                    {
                        TableProjects.Rows[i].Selected = true;
                    }
                }
            }
            else
            {
                TableProjects.Rows.Clear();
                TableProjects.ColumnHeadersVisible = false;
            }
        }

        /// <summary>
        /// Данные таблицы столбцов
        /// </summary>
        private void LoadColumns()
        {
            if (_projectManager.ActiveProject != null)
            {
                VewActiveProject(_projectManager.ActiveProject);

                _mappingColumns = _projectManager.ActiveProject.Columns;
                if ((_mappingColumns?.Count ?? 0) > 0)
                {
                    UpdateTableColumns();
                    TableColumns.Columns[0].Width = 70;
                    TableColumns.Columns[1].Width = 90;
                    TableColumns.Columns[3].Width = 60;
                    TableColumns.Columns[2].ReadOnly = true;
                    TableColumns.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    TableColumns.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    TableColumns.Columns[0].HeaderText = "Проверять";
                    TableColumns.Columns[1].HeaderText = "Обязательный";
                    TableColumns.Columns[2].HeaderText = "Название";
                    TableColumns.Columns[3].HeaderText = "Столбец";
                    TableColumns.Columns[4].Visible = false;
                }
                else
                {
                    TableColumns.Rows.Clear();
                    TableColumns.ColumnHeadersVisible = false;
                }
            }
        }

        /// <summary>
        ///  Лист \ первая строка
        /// </summary>
        private void LoadRangeValues()
        {
            if (_projectManager.ActiveProject != null)
            {
                TBoxFirstRowRangeValues.Text =
                    _projectManager.ActiveProject.RowStart.ToString();
                TBoxSheetName.Text = _projectManager.ActiveProject.AnalysisSheetName;
            }
            else
            {
                TBoxFirstRowRangeValues.Text = "";
                TBoxSheetName.Text = "";
            }
        }

        private void BtnAddProject_Click(object sender, EventArgs e)
        {
            string name = TboxProjectName.Text;
            if (!string.IsNullOrWhiteSpace(name))
            {
                ProjectManager projectManager = new ProjectManager();
                projectManager.CreateProject(name);
                LoadData();
            }
        }

        private void UpdateTableProject()
        {
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _projectManager.Projects.Count; i++)
            {
                Source.Add(_projectManager.Projects[i]);
            };

            TableProjects.DataSource = Source;
        }

        private void UpdateTableColumns()
        {
            BindingSource Source = new BindingSource();
            for (int i = 0; i < _mappingColumns.Count; i++)
            {
                Source.Add(_mappingColumns[i]);
            };
            TableColumns.DataSource = Source;
        }

        /// <summary>
        ///  Удалить строку
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TableColumns_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (e.RowIndex > 0)
            {
                DataGridViewRow row = TableColumns.Rows[e.RowIndex];
                string name = row.Cells[2].Value?.ToString() ?? "";
                ColumnMapping columnMapping = _mappingColumns.Find(x => x.Name == name);
                if (columnMapping != null)
                {
                    _mappingColumns.Remove(columnMapping);
                }
            }
        }
             
        private void TableProjects_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;
            if (e.ColumnIndex == 1 && e.RowIndex >= 0)
            {
                string name = TableProjects.Rows[e.RowIndex].Cells[1].Value?.ToString() ?? "";
                string filePath = TableProjects.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? "";
                Project project = _projectManager.Projects.Find(x => x.FileName == filePath);

                if (!string.IsNullOrEmpty(name) && project != null)
                {
                    project.Name = name;
                    project.Save();
                }
            }
        }

        private void TableProjects_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int row = e.RowIndex;
            int col = e.ColumnIndex;

            if (row >= 0)
            {
                string name = TableProjects.Rows[e.RowIndex].Cells[1].Value?.ToString() ?? "";
                string filePath = TableProjects.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? "";
                Project project = _projectManager.Projects.Find(x => x.FileName == filePath);
                if (e.ColumnIndex == 0)
                {
                    var active = TableProjects.Rows[e.RowIndex].Cells[0].Value;
                    if ((bool)active || string.IsNullOrWhiteSpace(name)) return;
                    TableProjects.Rows[e.RowIndex].Cells[0].Value = true;
                    Properties.Settings.Default.ActiveProjectName = name;
                    _projectManager.SetActiveProject();
                    LoadData();
                }
            }
        }

        /// <summary>
        ///  Сохранение активного проекта
        /// </summary>
        private void Save()
        {
            _projectManager.ActiveProject.Columns = _mappingColumns;
            _projectManager.ActiveProject.AnalysisSheetName = TBoxSheetName.Text;
            _projectManager.ActiveProject.RowStart = int.TryParse(TBoxFirstRowRangeValues.Text, out int fr) ? fr : 0;
            _projectManager.ActiveProject.Save();
        }

        /// <summary>
        ///  Кнопка сохранить
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnAccept_Click(object sender, EventArgs e)
        {
            Save();
            Close();
        }

        private void BtnOpenFolder_Click(object sender, EventArgs e)
        {
            string folder = ProjectManager.GetFolderProjects();
            System.Diagnostics.Process.Start(folder);
        }

        /// <summary>
        ///  Уддалить выделенный файл проекта 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (TableProjects.SelectedRows.Count > 0)
            {
                DataGridViewRow row = TableProjects.SelectedRows[0];
                string nameProject = row.Cells[1].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(nameProject))
                {
                    Project project = _projectManager.Projects.Find(x => x.Name == nameProject);
                    if (project != null) project.Delete();
                    TableProjects.Rows.Remove(row);
                    _projectManager.Projects.Remove(project);
                    LoadData();
                }
            }
        }
        private void BtnSelect_Click(object sender, EventArgs e)
        {
            if (TableProjects.SelectedRows.Count > 0)
            {
                DataGridViewRow row = TableProjects.SelectedRows[0];
                string name = row.Cells[1].Value.ToString() ?? "";
                Project newActiveProject = _projectManager.Projects.Find(p => p.Name == name);
                if (newActiveProject != null)
                {
                    row.Cells[0].Value = true;
                    _projectManager.ActiveProject = newActiveProject;
                    VewActiveProject(newActiveProject);
                }
            }
        }

        /// <summary>
        ///  Показать активный проект в заголовке формы.
        /// </summary>
        /// <param name="project"></param>
        private void VewActiveProject(Project project)
        {
            this.Text = $"Диспетчер проектов [{project.Name}]";
        }


        private void BtnSetCurrentSheet_Click(object sender, EventArgs e)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            if (ws != null)
            {
                TBoxSheetName.Text = ws.Name;
                int firstrow = ws.UsedRange.Row + ws.UsedRange.Rows.Count;
                TBoxFirstRowRangeValues.Text = firstrow.ToString();
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
             
        static readonly char[] _allowLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
        private void TableColumns_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridViewTextBoxEditingControl tb = (DataGridViewTextBoxEditingControl)e.Control;
            tb.KeyPress += new KeyPressEventHandler(dataGridViewTextBox_KeyPress);
            e.Control.KeyPress += new KeyPressEventHandler(dataGridViewTextBox_KeyPress);
        }

        private void dataGridViewTextBox_KeyPress(object sender, KeyPressEventArgs e)
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
    }
}


    //private void Closing(object sender, FormClosingEventArgs e)
    //{
    //    if (e.CloseReason == CloseReason.UserClosing) DialogResult = DialogResult.Cancel;
    //}