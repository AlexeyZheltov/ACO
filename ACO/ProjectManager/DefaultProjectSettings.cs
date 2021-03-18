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
                    Name = Project.ColumnsNames[StaticColumns.Level],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "A"
                });
                columns.Add(new ColumnMapping()
                {
                    Name =Project.ColumnsNames[StaticColumns.Number] ,
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "B"                   
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Cipher],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "D"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Classifier],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "E"
                });

                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Name],                  
                    Check = true,
                    Obligatory = true,
                    ColumnSymbol = "F"                  
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Material],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "H"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Size],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "I"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Type],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "J"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.VendorCode],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "K"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Producer],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "L"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Unit],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "M"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Count],
                    Check = false,
                    Obligatory = false,
                    ColumnSymbol = "N"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostMaterialsPerUnit],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "O"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostMaterialsTotal],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "P"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostWorksPerUnit],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "Q"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostWorksTotal ],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "R"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostTotalPerUnit],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "S"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.CostTotal],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "T"
                });
                columns.Add(new ColumnMapping()
                {
                    Name = Project.ColumnsNames[StaticColumns.Comment],
                    Check = false,
                    Obligatory = true,
                    ColumnSymbol = "U"
                });
   
                defProject = new Project()
                {
                    FileName = Path.Combine(ProjectManager.GetFolderProjects(), name + ".xml"),
                    AnalysisSheetName = "Рсч-П",
                    Name = name,
                    RowStart = 10,
                    
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
