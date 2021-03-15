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
                    Name =Project.ColumnsNames[StaticColumns.Number] ,
                    //Value = "№ п/п",
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "J"
                    //Address = "$J$3",
                    //Row = 3,
                    //Column = 10
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.NoEstimatesAndCalculations],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "K"
                }) ;
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.NameVOR],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "L"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Code],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "M"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.ProductCode],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "N"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Producer],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "O"
                });
                
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Name],
                    //Value = "Наименование работ и затрат",
                    Check = true,
                    Obligatory = false,
                    ColumnSymbol = "P"
                    //Address = "$P$3",
                    //Row = 3,
                    //Column = 16
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Unit],
                    //Value = "Ед. изм.",
                    Check = true,
                    Obligatory = false,
                    ColumnSymbol = "Q"
                    //Address = "$Q$3",
                    //Row = 3,
                    //Column = 17
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CountProject], //"Кол-во по проекту",
                    //Name = "Кол-во по проекту",
                    Check = true,
                    Obligatory = false,
                    ColumnSymbol = "R"
                    //Address = "$R$3",
                    //Row = 3,
                    //Column = 18
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CountSH],//"Кол-во СХ",
                    //Value = "Кол-во СХ",
                    Check = true,
                    Obligatory = false,
                    ColumnSymbol = "S"
                    //Address = "$S$3",
                    //Row = 3,
                    //Column = 19
                });

                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostMaterialsPerUnit],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "T"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostMaterialsTotal],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "U"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostWorksPerUnit],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "V"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostWorksTotal],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "W"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostTotalPerUnit],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "X"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostTotal],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "Y"
                });

                defProject = new Project()
                {
                    FileName = Path.Combine(ProjectManager.GetFolderProjects(), name + ".xml"),
                    AnalysisSheetName = "Анализ",
                    Name = name,
                    RowStart = 7,
                    
                    //RangeValuesStart = 20,
                    //RangeValuesEnd = 25,
                    //FirstColumnOffer = 27,
                    //LastColumnOffer = 47,
                    Columns = columns
                };
                defProject.Save();
            }
            return defProject;
        }
    }
}
