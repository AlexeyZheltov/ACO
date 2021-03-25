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
        Producer,
        Unit,
        Amount,
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
                { StaticColumns.Level, "Маркер иерархии 1-2-3-4" },
                { StaticColumns.File, "Файл" },
                { StaticColumns.Number, "№ п/п" },
                { StaticColumns.Cipher, "ШИФР" },
                { StaticColumns.Classifier, "Классификатор" },
                { StaticColumns.Name, " НАИМЕНОВАНИЕ РАБОТ" },
                { StaticColumns.Code, "МАРКИРОВКА/ ОБОЗНАЧЕНИЕ" },
                { StaticColumns.Material, "МАТЕРИАЛ" },
                { StaticColumns.Size, "ФОРМАТ / ГАБАРИТНЫЕ РАЗМЕРЫ / ДИАМЕТР (Ф) ММ" },
                { StaticColumns.Type, "ТИП, МАРКА, ОБОЗНАЧЕНИЕ ДОКУМЕНТА, ОПРОСНОГО ЛИСТА" },
                { StaticColumns.VendorCode, "АРТИКУЛ" },
                { StaticColumns.Producer, "ПРОИЗВОДИТЕЛЬ" },
                { StaticColumns.Unit, "ЕД.ИЗМ." },
                { StaticColumns.Amount, "КОЛ-ВО" },
                { StaticColumns.CostMaterialsPerUnit, "ЦЕНА МАТЕРИАЛОВ, РУБ БЕЗ НДС. ЗА ЕДИНИЦУ" },
                { StaticColumns.CostMaterialsTotal, "ЦЕНА МАТЕРИАЛОВ, РУБ БЕЗ НДС. ВСЕГО" },
                { StaticColumns.CostWorksPerUnit, "ЦЕНА РАБОТ, РУБ БЕЗ НДС. ЗА ЕДИНИЦУ" },
                { StaticColumns.CostWorksTotal, "ЦЕНА РАБОТ, РУБ БЕЗ НДС. ВСЕГО" },
                { StaticColumns.CostTotalPerUnit, "ЦЕНА ЗА ЕДИНИЦУ. РУБ БЕЗ НДС" },
                { StaticColumns.CostTotal, "ИТОГО. РУБ БЕЗ НДС" },
                { StaticColumns.Comment, "ПРИМЕЧАНИЕ" }
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
