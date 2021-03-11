using Excel = Microsoft.Office.Interop.Excel;
using ACO.Offers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ACO.ProjectManager
{
    class ProjectManager
    {
        public Project ActiveProject
        {
            get
            {
                foreach (Project project in Projects)
                {
                    if (project.Active)
                    {
                        if (_ActiveProject == null)
                        {
                            _ActiveProject = project;
                        }
                        else
                        {
                            project.Active = false;
                            project.Save();
                        }
                    }
                }
                if (_ActiveProject is null && Projects.Count > 0)
                    _ActiveProject = Projects[0];

                return _ActiveProject;
            }
            set
            {
                _ActiveProject = value;


                foreach (Project p in Projects)
                {
                    p.Active = p.Name == _ActiveProject.Name;
                    p.Save();
                }

            }
        }
        private Project _ActiveProject;

        /// <summary>
        ///  Коллекция всех проектов
        /// </summary>
        public List<Project> Projects
        {
            get
            {
                if (_Projects is null)
                {
                    _Projects = new List<Project>();
                    string folder = GetFolderProjects();
                    string[] files = Directory.GetFiles(folder);
                    foreach (string file in files)
                    {
                        if (new FileInfo(file).Extension == ".xml")
                        {
                            Project project = Project.GetFromXML(file);
                            _Projects.Add(project);
                        }
                    }
                }
                return _Projects;
            }
            private set
            {
                _Projects = value;
            }
        }

        private List<Project> _Projects;
        public void CreateProject(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) { return; }
            string path = GetFolderProjects();
            string filename = Path.Combine(path, name + ".xml");
            if (!File.Exists(filename))
            {
                CreateNewProjectXML(name, filename);
            }
            else
            {
                if (MessageBox.Show("Удалить старый файл?", "Файл уже существует!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    File.Delete(filename);
                    CreateNewProjectXML(name, filename);
                }
            }
        }

        /// <summary>
        ///  Создать новый файл проекта
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="path"></param>
        public void CreateNewProjectXML(string projectname, string path)
        {
            foreach (Project project in Projects)
            {
                project.Active = false;
                project.Save();
            }
            XElement root = new XElement("project");
            root.Add(new XAttribute("ProjectName", projectname));
            root.Add(new XAttribute("Active", true));
            XElement xeColumns = new XElement("Columns");

            /// Скопировать настройки столбцов из активного проекта
            if ((ActiveProject?.Columns?.Count ?? 0) > 0)
            {
                foreach (ColumnMapping column in ActiveProject.Columns)
                {
                    xeColumns.Add(column.GetXElement());
                }
            }
            root.Add(xeColumns);
            XDocument xdoc = new XDocument(root);
            xdoc.Save(path);
        }

        /// <summary>
        /// Генерирует путь к файлу
        /// </summary>
        /// <param name="file">Имя файла</param>
        /// <returns>Путь к файлу в AppData</returns>
        private static string GetPathTo(string file)
        {
            string path = GetFolderProjects();
            return Path.Combine(path, file);
        }
        public static string GetFolderProjects()
        {
            string path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Spectrum",
            "ACO",
            "Projects");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }


        private static string GetApplicationSettingsFilename()
        {
            string path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Spectrum",
            "ACO");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            string filename = Path.Combine(path, "settings.xml");
            return filename;
        }

        /// <summary>
        ///  
        /// </summary>
        /// <param name="offer"></param>
        internal void AddOffer(Offer offer)
        {
            //foreach (Item itm in offer.Items)
            //{
            //    int row = itm.Row;
            //    int col = ActiveProject.Columns.Find(c => c.Value == itm.Header)?.Column ?? 0;

            //    foreach (ColumnMapping column in ActiveProject.Columns)
            //    {
            //        //column.Column
            //        //cell
            //    }
            //}
        }

        internal void PrintOffer(Offer offer)
        {
          Excel.Worksheet sh = Globals.ThisAddIn.Application.ActiveSheet;
            //
            //TODO Определить место вставки  
            List<ColumnMapping> columnsMapping = ActiveProject.Columns;


            foreach (Record record in offer.Records)
            {
            int rowPrint = GetRow(record.Number);
            if (rowPrint == 0) throw new AddInException("Не удалось определить строку вставки. Номер перечня: "+ record.Number);
                foreach (ColumnMapping col in columnsMapping)
                {
                    int columnPrint = 0; //TODO определить столбец вставки 
                    if (record.Values.ContainsKey(col.Value ))
                    {
                       object val = record.Values[col.Value];
                        Excel.Range cellPrint = sh.Cells[rowPrint, columnPrint];
                        cellPrint.Value = val;
                    }

                }
            }
        }

        private int GetRow(string number)
        {
                //TODO определить строку вставки 
            int row = 0;
            if (row == 0) row = InsertRow(number);

            return row;
        }

        private int InsertRow(string number)
        {
                //TODO если такого пункта нет вставить строку
            int row = 0;

            
            return row;
        }
    }
}
