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
            string filename = Path.Combine(ProjectManager.GetFolderProjects(), name + ".xml");
            Project defProject;
            if (File.Exists(filename))
            {
                defProject = GetFromXML(filename);
            }
            else
            {

                List<ColumnMapping> columns = new List<ColumnMapping>
                {
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Level],
                        Check = false,
                        Obligatory = false,
                        ColumnSymbol = "A"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Number],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "C"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Cipher],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "D"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Classifier],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "E"
                    },

                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Name],
                        Check = true,
                        Obligatory = true,
                        ColumnSymbol = "F"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Material],
                        Check = false,
                        Obligatory = false,
                        ColumnSymbol = "H"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Size],
                        Check = false,
                        Obligatory = false,
                        ColumnSymbol = "I"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Type],
                        Check = false,
                        Obligatory = false,
                        ColumnSymbol = "J"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.VendorCode],
                        Check = false,
                        Obligatory = false,
                        ColumnSymbol = "K"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Producer],
                        Check = false,
                        Obligatory = false,
                        ColumnSymbol = "L"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Unit],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "M"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Amount],
                        Check = true,
                        Obligatory = true,
                        ColumnSymbol = "N"
                    },
                    // new ColumnMapping()
                    //{
                    //    Name = Project.ColumnsNames[StaticColumns.ContractorAmountAmount],
                    //    Check = true,
                    //    Obligatory = true,
                    //    ColumnSymbol = ""
                    //},
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.CostMaterialsPerUnit],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "O"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.CostMaterialsTotal],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "P"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.CostWorksPerUnit],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "Q"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.CostWorksTotal],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "R"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.CostTotalPerUnit],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "S"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.CostTotal],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "T"
                    },
                    new ColumnMapping()
                    {
                        Name = Project.ColumnsNames[StaticColumns.Comment],
                        Check = false,
                        Obligatory = true,
                        ColumnSymbol = "U"
                    }
                };

                defProject = new Project()
                {
                    FileName = Path.Combine(ProjectManager.GetFolderProjects(), name + ".xml"),
                    AnalysisSheetName = "Рсч-П",
                    Name = name,
                    RowStart = 10,
                    Columns = columns
                };
                defProject.Save();
            }
            return defProject;
        }
    }
}
