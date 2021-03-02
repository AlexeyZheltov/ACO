using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Spectrum.SpLoader.XMLSetting
{
    /// <summary>
    /// Служит для сохранения и извлечения данных о маппингах
    /// </summary>
    /// <remarks>
    /// Сохраняет данные в AppData/Local/Spectrum/SaLoader/mappings.xml и
    /// AppData/Local/Spectrum/SaLoader/setting.xml
    /// </remarks>
    static class XMLManager
    {
        static class MappingConsts
        {
            public const string FileName = "mappings.xml";
            public const string Root = "Mappings";
            public const string ElementName = "Mapping";
            public const string Selected = "Selected";
            public const string Name = "Name";
            public const string Omni = "Omni";
            public const string WorkName = "WorkName";
            public const string Marking = "Marking";
            public const string Material = "Material";
            public const string Format = "Format";
            public const string Type = "Type";
            public const string Article = "Article";
            public const string Maker = "Maker";
            public const string Unit = "Unit";
            public const string Amount = "Name";
            public const string Note = "Note";
        }
        
        static class SettingConsts
        {
            public const string FileName = "settings.xml";
            public const string Root = "Settings";
            public const string Parameter = "Parameter";
        }


        
        /// <summary>
        /// Сохраняет в файл без кеширования
        /// </summary>
        /// <param name="data">Словарь ключ: имя меппинга, значение: сам меппинг</param>
        public static void Save(Dictionary<string, Mapping> data, string selectedMapping)
        {
            XElement root = new XElement(MappingConsts.Root, new XAttribute(MappingConsts.Selected, selectedMapping));
            
            foreach(var item in data.Values)
                root.Add(new XElement(MappingConsts.ElementName,
                            new XAttribute(MappingConsts.Name, item.Name),
                            new XElement(MappingConsts.Omni, item.Omni),
                            new XElement(MappingConsts.WorkName, item.WorkName),
                            new XElement(MappingConsts.Marking, item.Marking),
                            new XElement(MappingConsts.Material, item.Material),
                            new XElement(MappingConsts.Format, item.Format),
                            new XElement(MappingConsts.Type, item.Type),
                            new XElement(MappingConsts.Article, item.Article),
                            new XElement(MappingConsts.Maker, item.Maker),
                            new XElement(MappingConsts.Unit, item.Unit),
                            new XElement(MappingConsts.Amount, item.Amount),
                            new XElement(MappingConsts.Note, item.Note)));

            XDocument xdoc = new XDocument(root);
            xdoc.Save(GetPathTo(MappingConsts.FileName));
            
        }

        /// <summary>
        /// Сохраняет в файл без кеширования
        /// </summary>
        /// <param name="data">Список параметров</param>
        public static void Save(params string[] data)
        {
            XElement root = new XElement(SettingConsts.Root);
            foreach (string item in data)
                root.Add(new XElement(SettingConsts.Parameter, item));


            XDocument xdoc = new XDocument();
            xdoc.Add(root);
            xdoc.Save(GetPathTo(SettingConsts.FileName));
        }

        /// <summary>
        /// Считывает сохраненные мэппинги
        /// </summary>
        /// <returns>Словарь ключ: имя меппинга, значение: сам меппинг</returns>
        public static (Dictionary<string, Mapping>, string selectedMapping) ReadMapping()
        {
            string path = GetPathTo(MappingConsts.FileName);

            if (!File.Exists(path)) return (new Dictionary<string, Mapping>(), "");

            XDocument xdoc = XDocument.Load(path);
            XElement root = xdoc.Root;
            Dictionary<string, Mapping> buffer = new Dictionary<string, Mapping>();

            (from xe in root.Elements(MappingConsts.ElementName)
             select new Mapping()
             {
                 Name = xe.Attribute(MappingConsts.Name).Value,
                 Omni = xe.Element(MappingConsts.Omni).Value,
                 WorkName = xe.Element(MappingConsts.WorkName).Value,
                 Marking = xe.Element(MappingConsts.Marking).Value,
                 Material = xe.Element(MappingConsts.Material).Value,
                 Format = xe.Element(MappingConsts.Format).Value,
                 Type = xe.Element(MappingConsts.Type).Value,
                 Article = xe.Element(MappingConsts.Article).Value,
                 Maker = xe.Element(MappingConsts.Maker).Value,
                 Unit = xe.Element(MappingConsts.Unit).Value,
                 Amount = xe.Element(MappingConsts.Amount).Value,
                 Note = xe.Element(MappingConsts.Note).Value
             })
             .ToList()
             .ForEach(i => buffer.Add(i.Name, i));

            return (buffer, root.Attribute(MappingConsts.Selected).Value);
        }

        /// <summary>
        /// Считывает сохраненные пути к шаблону и omni файлу
        /// </summary>
        /// <returns></returns>
        public static List<string> ReadSetting()
        {
            string path = GetPathTo(SettingConsts.FileName);

            if (!File.Exists(path)) return null;

            XDocument xdoc = XDocument.Load(path);

            return (from item in xdoc.Root.Elements(SettingConsts.Parameter)
                    select item.Value).ToList();
        }

        /// <summary>
        /// Генерирует путь к файлу
        /// </summary>
        /// <param name="file">Имя файла</param>
        /// <returns>Путь к файлу в AppData</returns>
        private static string GetPathTo(string file)
        {
            string path = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "Spectrum",
                        "SpLoader");

            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            return Path.Combine(path, file);
        }
    }
}
