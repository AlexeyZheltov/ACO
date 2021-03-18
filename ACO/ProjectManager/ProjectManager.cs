
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
        /// <summary>
        ///  Проект отмеченный как текущий
        /// </summary>
        public Project ActiveProject
        {           
            get
            {
                if (_ActiveProject == null)
                {
                     SetActiveProject();
                }
                return _ActiveProject;
            }
            set
            {
                _ActiveProject = value;
                Properties.Settings.Default.ActiveProjectName = _ActiveProject?.Name ?? "";
                Projects.ForEach(x => x.Active = x.Name == _ActiveProject.Name);
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
                           string activeProjectName  = Properties.Settings.Default.ActiveProjectName ;
                            if (project.Name == activeProjectName) project.Active = true;
                            _Projects.Add(project);
                        }
                    }
                    if (_Projects.Count == 0) _Projects.Add( DefaultProject.Get());
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
                CreateNewProject(name, filename);
            }
            else
            {
                if (MessageBox.Show("Удалить старый файл?", "Файл уже существует!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    File.Delete(filename);
                    CreateNewProject(name, filename);
                }
            }
        }

        public void SetActiveProject()
        {            
            Project activeProject = null;
            string activeProjectName= Properties.Settings.Default.ActiveProjectName;            
            activeProject = Projects.Find(x => x.Name == activeProjectName);
            if (activeProject is null && Projects.Count > 0)
            {
               activeProject= Projects[0];           
            }
            ActiveProject = activeProject;            
        }

        /// <summary>
        ///  Создать новый файл проекта
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="path"></param>
        public void CreateNewProject(string projectname, string path)
        {
            Project newProject = DefaultProject.Get();
            newProject.FileName = path;
            newProject.Name = projectname;

            //XElement root = new XElement("project");
            //root.Add(new XAttribute("ProjectName", projectname));
            //root.Add(new XAttribute("Active", true));
            //XElement xeColumns = new XElement("Columns");

            ///// Скопировать настройки столбцов из активного проекта
            //if ((ActiveProject?.Columns?.Count ?? 0) > 0)
            //{
            //    foreach (ColumnMapping column in ActiveProject.Columns)
            //    {
            //        xeColumns.Add(column.GetXElement());
            //    }
            //}
            //root.Add(xeColumns);
            //XDocument xdoc = new XDocument(root);
            newProject.Save();
            ActiveProject = newProject;
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
    }
}
