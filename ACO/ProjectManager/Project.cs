using ACO.ExcelHelpers;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ACO.ProjectManager
{
    public enum StaticColumns
    {
        Level,
        File,
        Number,
        Cipher,
        Classifier,
        Name,
        Code,
        Material,
        Size,
        Type,
        VendorCode,
        Label,
        Producer,
        Unit,
        Amount,
        ContractorAmount,
        CostMaterialsPerUnit,
        CostMaterialsTotal,
        CostWorksPerUnit,
        CostWorksTotal,
        CostTotalPerUnit,
        CostTotal,
        Comment
    }
       

    class Project
    {
        public  bool Active {get;set;}

        /// <summary>
        ///  Название проекта
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        ///  Путь к файлу
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        ///  Название листа 
        /// </summary>
        public string AnalysisSheetName { get; set; }
        
        /// <summary>
        /// Строка начала данных
        /// </summary>
        public int RowStart { get; set; } = 10;

        /// <summary>
        ///  Адреса ячеек шапки используемых столбцов
        /// </summary>
        public List<ColumnMapping> Columns { get; set; }

        public static Dictionary<StaticColumns, string> ColumnsNames =
            new Dictionary<StaticColumns, string>
            {
                { StaticColumns.Level, "Уровень" },
                { StaticColumns.File, "Файл" },
                { StaticColumns.Number, "№ п/п" },
                { StaticColumns.Cipher, "Шифр" },
                { StaticColumns.Classifier, "Классификатор" },
                { StaticColumns.Name, "Наименование работ" },
                { StaticColumns.Code, "Маркировка / Обозначение" },
                { StaticColumns.Material, "Материал" },
                { StaticColumns.Size, "Формат / Габаритные размеры / Диаметр" },
                { StaticColumns.Type, "Тип, марка, обозначение" },
                { StaticColumns.VendorCode, "Артикул" },
                { StaticColumns.Producer, "Производитель" },
                { StaticColumns.Label, "Маркировка" },
                { StaticColumns.Unit, "Ед. изм." },
                { StaticColumns.Amount, "Кол-во" },
                { StaticColumns.ContractorAmount, "Кол-во (подрядчик)" },
                { StaticColumns.CostMaterialsPerUnit, "Цена материалы за ед." },
                { StaticColumns.CostMaterialsTotal, "Цена материалы всего" },
                { StaticColumns.CostWorksPerUnit, "Цена работы за ед." },
                { StaticColumns.CostWorksTotal, "Цена работы всего" },
                { StaticColumns.CostTotalPerUnit, "Итого за ед." },
                { StaticColumns.CostTotal, "Итого" },
                { StaticColumns.Comment, "Примечание" }
            };
            

        public Project() { }

        /// <summary>
        ///  Сохранить XML - Файл
        /// </summary>
        public void Save()
        {
            XElement root = new XElement("project");
            XAttribute xaName = new XAttribute("ProjectName", Name);
            root.Add(xaName);

            XElement xeSheets = new XElement("Sheets");
            XElement xeAnalysisSheet = new XElement("AnalysisSheet");
            xeAnalysisSheet.Add(new XAttribute("Name", AnalysisSheetName));
            XElement xeRows = new XElement("Rows");
            XElement xeRowStart = new XElement("RowStart");
            xeRowStart.Add(new XAttribute("Row", RowStart.ToString()));
            xeRows.Add(xeRowStart);
            xeAnalysisSheet.Add(xeRows);

            XElement xeColumns = new XElement("Columns");
            /// Диапазон значения
            XElement xeRangeValues = new XElement("RangeValues");
            xeAnalysisSheet.Add(xeRangeValues);

            foreach (ColumnMapping cell in Columns)
            {
                XElement xeColumn = cell.GetXElement();
                xeColumns.Add(xeColumn);
            }
            xeAnalysisSheet.Add(xeColumns);
            /// Диапазон предложения 
            XElement xeRangeOffer = new XElement("RangeOffer");
            xeAnalysisSheet.Add(xeRangeOffer);

            xeSheets.Add(xeAnalysisSheet);
            root.Add(xeSheets);
            XDocument xdoc = new XDocument(root);
            xdoc.Save(FileName);
        }

        /// <summary>
        ///  Загрузить Project из XML файла
        /// </summary>
        public static Project GetFromXML(string filename)
        {
            Project project = new Project();
            XDocument xdoc = XDocument.Load(filename);
            XElement root = xdoc.Root;
            project.FileName = filename;
            project.Name = root.Attribute("ProjectName").Value?.ToString() ?? "";

            XElement xeSheets = root.Element("Sheets");
            /// Лист Анализ
            XElement xeAnalysisSheet = xeSheets.Element("AnalysisSheet");
            project.AnalysisSheetName = xeAnalysisSheet.Attribute("Name").Value?.ToString() ?? "";

            /// Строки
            XElement xeRows = xeAnalysisSheet.Element("Rows");
            XElement xeRowStart = xeRows.Element("RowStart");
            project.RowStart = int.TryParse(xeRowStart.Attribute("Row").Value, out int r) ? r : 0;
            /// Столбцы
            project.Columns = LoadColumnsFromXElement(xeAnalysisSheet.Element("Columns"));
            return project;
        }

        /// <summary>
        ///  Прочитать столбцы 
        /// </summary>
        /// <param name="xElement"></param>
        /// <returns></returns>
        private static List<ColumnMapping> LoadColumnsFromXElement(XElement xElement)
        {
            List<ColumnMapping> columns = new List<ColumnMapping>();
            if (xElement != null)
            {
                foreach (XElement xcol in xElement.Elements())
                {
                    columns.Add(ColumnMapping.GetCellFromXElement(xcol));
                }
            }
            return columns;
        }

        /// <summary>
        /// Установить цифровые момера столбцов
        /// </summary>
        /// <param name="sheetProject"></param>
        internal void SetColumnNumbers(Worksheet sheetProject)
        {
            foreach (ColumnMapping mapping in Columns)
            {
                if (!string.IsNullOrWhiteSpace(mapping.ColumnSymbol))
                {
                    mapping.Column = sheetProject.Range[$"{mapping.ColumnSymbol}1"].Column;
                }
            }
        }

        public ColumnMapping GetColumn(StaticColumns name)
        {
            ColumnMapping mapping = Columns.Find(x => x.Name == ColumnsNames[name]);
            if (mapping is null) throw new AddInException($"Маппинг столбца \"{ColumnsNames[name]}\" не найден!");
            return mapping;
        }

        internal void Delete()
        {
            if (File.Exists(FileName)) File.Delete(FileName);
        }
    }
}
