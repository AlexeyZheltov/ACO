using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO.ProjectManager
{
    /// <summary>
    ///  Создает стандартные настройки проекта 
    /// </summary>
    class DefaultProject : Project
    {

        private const string name = "default";

        private DefaultProject() { }
        public static Project Get()
        {
            Project defProject = default;
            string filename = Path.Combine(ProjectManager.GetFolderProjects(), name + ".xml");
            if (File.Exists(filename))
            {
                defProject = GetFromXML(filename);
            }
            else
            {

                List<ColumnMapping> columns = new List<ColumnMapping>();

                columns.Add(new ColumnMapping()
                {
                    Value = "№ п/п",
                    Check = false,
                    Obligatory = true,
                    Address = "$J$3",
                    Row = 3,
                    Column = 10
                });
                columns.Add(new ColumnMapping()
                {
                    Value = "Наименование работ и затрат",
                    Check = true,
                    Obligatory = false,
                    Address = "$P$3",
                    Row = 3,
                    Column = 16
                });
                columns.Add(new ColumnMapping()
                {
                    Value = "Ед. изм.",
                    Check = true,
                    Obligatory = false,
                    Address = "$Q$3",
                    Row = 3,
                    Column = 17
                });
                columns.Add(new ColumnMapping()
                {
                    Value = "Кол-во по проекту",
                    Check = true,
                    Obligatory = false,
                    Address = "$R$3",
                    Row = 3,
                    Column = 18
                });
                columns.Add(new ColumnMapping()
                {
                    Value = "Кол-во СХ",
                    Check = true,
                    Obligatory = false,
                    Address = "$S$3",
                    Row = 3,
                    Column = 19
                });

                defProject = new Project()
                {
                    FileName = Path.Combine(ProjectManager.GetFolderProjects(), name + ".xml"),
                    AnalysisSheetName = "Анализ",
                    Name = name,
                    RowStart = 7,
                    RangeValuesStart = 20,
                    RangeValuesEnd = 25,
                    Columns = columns
                };
                defProject.Save();
            }
            return defProject;
        }

       
    }

        //public static Project Get()
        //{
        //    ProjectManager manager = new ProjectManager();
        //    Project project = manager.Projects.Find(x => x.Name == name);
        //    if (project is null) project = Create();
        //    return project;
        //}

        //public static Project Create()
        //{
        //    List<ColumnMapping> columns = new List<ColumnMapping>();

        //    columns.Add(new ColumnMapping()
        //    {
        //        Value = "№ п/п",
        //        Check = false,
        //        Obligatory = true,
        //        Address = "$J$3",
        //        Row = 3,
        //        Column = 10
        //    });
        //    columns.Add(new ColumnMapping()
        //    {
        //        Value = "Наименование работ и затрат",
        //        Check = true,
        //        Obligatory = false,
        //        Address = "$P$3",
        //        Row = 3,
        //        Column = 16
        //    });
        //    columns.Add(new ColumnMapping()
        //    {
        //        Value = "Ед. изм.",
        //        Check = true,
        //        Obligatory = false,
        //        Address = "$Q$3",
        //        Row = 3,
        //        Column = 17
        //    });
        //    columns.Add(new ColumnMapping()
        //    {
        //        Value = "Кол-во по проекту",
        //        Check = true,
        //        Obligatory = false,
        //        Address = "$R$3",
        //        Row = 3,
        //        Column = 18
        //    });
        //    columns.Add(new ColumnMapping()
        //    {
        //        Value = "Кол-во СХ",
        //        Check = true,
        //        Obligatory = false,
        //        Address = "$S$3",
        //        Row = 3,
        //        Column = 19
        //    });

        //    Project project = new Project()
        //    {
        //        FileName = Path.Combine(ProjectManager.GetFolderProjects(), name + ".xml"),
        //        AnalysisSheetName = "Анализ NEBO",
        //        Name = name,
        //        RowStart = 7,
        //        RangeValuesStart = 20,
        //        RangeValuesEnd = 25,
        //        Columns = columns
        //    };
        //    project.Save();
        //    return project;
        //}
    
}
