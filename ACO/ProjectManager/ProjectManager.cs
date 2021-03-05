using ACO.ExcelHelpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ACO.ProjectManager
{
    class ProjectManager
    {

       public Project ActiveProject {
            get {
               if(_ActiveProject is null)
                {
                    foreach (Project project in Projects)
                    {
                        if (project.Active)
                        {
                            _ActiveProject = project;
                            break;
                        }
                    }
                    if (_ActiveProject is null && Projects.Count>0)
                        _ActiveProject = Projects[0];
                }
                return _ActiveProject;
            }
            set
            {
                _ActiveProject = value;
            }
        }
        private Project _ActiveProject;
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
                            Project project = LoadProject(file);
                            _Projects.Add(project);
                        }
                    }
                }
                return _Projects;
            }
            set
            {
                _Projects = value;
            }
        }



        private List<Project> _Projects;
        public void CreateProject(string name)
        {
            if (!string.IsNullOrEmpty(name))
            {
                string filename = GetPathTo(name+".xml") ;
                if (!File.Exists(filename))
                {
                    CreateNewProjectXML(name, filename);
                }
            }
        }

        public  void CreateNewProjectXML(string projectname, string path)
        {
            Projects?.ForEach(x => x.Active = false);
            XElement root = new XElement("ProjectName", projectname);
            XElement xe = new XElement("Active", true);
            //root.Add(xe);
            XDocument xdoc = new XDocument(root);
            xdoc.Add(xe);
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
            return Path.Combine(path, file );
        }
        private static string GetFolderProjects()
        {
            string path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Spectrum",
            "ACO");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path ;
        }

        /// <summary>
        /// Сохраняет в файл без кеширования
        /// </summary>
        /// <param name="data">Словарь ключ: имя меппинга, значение: сам меппинг</param>
        //public static void Save(Dictionary<string, string> data, string selectedMapping)
        //{
        //    //  XElement root = new XElement(MappingConsts.Root, new XAttribute(MappingConsts.Selected, selectedMapping));
        //    XElement root = new XElement("ProjectName", "");
        //    ///  foreach (var item in data.Values)
        //    //root.Add(new XElement(MappingConsts.ElementName,
        //    //            new XAttribute(MappingConsts.Name, item.Name),
        //    //            new XElement(MappingConsts.Omni, item.Omni),
        //    //            new XElement(MappingConsts.WorkName, item.WorkName),
        //    //            new XElement(MappingConsts.Marking, item.Marking),
        //    //            new XElement(MappingConsts.Material, item.Material),
        //    //            new XElement(MappingConsts.Format, item.Format),
        //    //            new XElement(MappingConsts.Type, item.Type),
        //    //            new XElement(MappingConsts.Article, item.Article),
        //    //            new XElement(MappingConsts.Maker, item.Maker),
        //    //            new XElement(MappingConsts.Unit, item.Unit),
        //    //            new XElement(MappingConsts.Amount, item.Amount),
        //    //            new XElement(MappingConsts.Note, item.Note)));

        //    XDocument xdoc = new XDocument(root);
        //    xdoc.Save("");

        //}

        private Project LoadProject(string file)
        {
            Project project = new Project();

            XDocument xdoc = XDocument.Load(file);
            XElement root = xdoc.Root;
            project.FileName = file;
            project.Name = root.Value;
            project.Active =bool.Parse(xdoc.Element("Active").Value);
            project.Columns = LoadColumnsFromXElement(xdoc.Element("Columns"));
            //Dictionary<string, Mapping> buffer = new Dictionary<string, Mapping>();
            //(from xe in root.Elements(ProjectName)
            // select new Mapping()
            // {
            //   project.Name = xe.Attribute(ProjectName).Value
            // })
            // .ToList()
            // .ForEach(i => buffer.Add(i.Name, i));

                //return (buffer, root.Attribute(MappingConsts.Selected).Value);
                //project.Name = 
            return project;
        }

        private List<ColumnMapping> LoadColumnsFromXElement(XElement xElement)
        {
            List<ColumnMapping> columns = new List<ColumnMapping>();

            foreach (XElement xcol in xElement.Elements("Column"))
            {
                ColumnMapping mapping = new ColumnMapping()
                {
                    Name = xcol.Element("Name").Value,
                    Cell = GetCellFromXElement(xcol.Element("Cell"))
                };
                columns.Add(mapping);
            }
            return columns;
        }

        private Cell GetCellFromXElement(XElement xElement)
        {
            return new Cell()
            {
                Value = xElement.Element("Value").Value,
                Row =int.Parse( xElement.Element("Row").Value),
                Column =int.Parse(xElement.Element("Column").Value),
                ColumnString = xElement.Element("ColumnString").Value,
                Address = xElement.Element("Address").Value
            };

        }
    }
}
