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
                    Name = "",
                    Value = "№ п/п",
                    Check = false,
                    Obligatory = true,
                    Address = "$J$3",
                    Row = 3,
                    Column = 10
                });
                columns.Add(new ColumnMapping()
                {
                    Name = "Наименование работ и затрат",
                    Value = "Наименование работ и затрат",
                    Check = true,
                    Obligatory = false,
                    Address = "$P$3",
                    Row = 3,
                    Column = 16
                });
                columns.Add(new ColumnMapping()
                {
                    Name = "Ед. изм.",
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
                    Name = "Кол-во по проекту",
                    Check = true,
                    Obligatory = false,
                    Address = "$R$3",
                    Row = 3,
                    Column = 18
                });
                columns.Add(new ColumnMapping()
                {
                    Value = "Кол-во СХ",
                    Name = "Кол-во СХ",
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
                    FirstColumnOffer = 27,
                    LastColumnOffer = 47,
                    Columns = columns
                };
                defProject.Save();
            }
            return defProject;
        }
    }
}
