using System.Collections.Generic;

namespace Spectrum.SpLoader.XMLSetting
{
    /// <summary>
    /// Класс работы с набором мэппингов.
    /// </summary>
    /// <remarks>Необходимо вызвать метод Load для загрузки данных из файла.</remarks>
    static class MappingManager
    {
        static Dictionary<string, Mapping> _mappings = new Dictionary<string, Mapping>();

        /// <summary>
        /// Только для чтения. Возвращает текущий мэппинг
        /// </summary>
        public static Mapping Current { get; set; }

        /// <summary>
        /// Загружает набор мэппингов из файла
        /// </summary>
        public static void Load()
        {
            string current;
            (_mappings, current) = XMLManager.ReadMapping();

            if (_mappings.TryGetValue(current, out Mapping mapping))
                Current = mapping;  
        }

        /// <summary>
        /// Получает список имен всех маппингов
        /// </summary>
        /// <returns></returns>
        public static string[] GetMappingList()
        {
            List<string> buffer = new List<string>();
            
            foreach (string key in _mappings.Keys)
                buffer.Add(key);
            
            return buffer.ToArray();
        }

        /// <summary>
        /// Сохраняет набор мэппингов в файл
        /// </summary>
        public static void Save() => XMLManager.Save(_mappings, Current.Name);
        
        public static void Delete()
        {
            //Перключение на следующий маппинг если есть
            if (_mappings.ContainsKey(Current.Name))
                _mappings.Remove(Current.Name);
        }

        /// <summary>
        /// Переименовывает мэппинг
        /// </summary>
        /// <param name="oldName">Старое имя набора</param>
        /// <param name="newName">Новое имя набора</param>
        /// <returns>true - если удачно</returns>
        public static bool Rename(string oldName, string newName)
        {
            if (_mappings.TryGetValue(oldName, out Mapping mapping))
            {
                _mappings.Remove(oldName);
                _mappings.Add(newName, mapping);
                Current.Name = newName;
                return true;
            }
            else return false;
        }

        /// <summary>
        /// Устанавливает мэппинг как текущий
        /// </summary>
        /// <param name="name">Имя мэппинга</param>
        /// <returns>Возвращает мэппинг, либо null если мэппинг с таким именем не выбран</returns>
        public static Mapping Select(string name)
        {
            if (_mappings.TryGetValue(name, out Mapping mapping))
            {
                Current = mapping;
                return mapping;
            }
            else return null;
        }

        /// <summary>
        /// Добавляет новый мэппинг к набору
        /// </summary>
        /// <param name="name">Имя нового мэппинга</param>
        public static void Add(string name)
        {
            Mapping mapping = new Mapping() { Name = name };
            _mappings.Add(name, mapping);
            Current = mapping;
        }

        /// <summary>
        /// Удаляет мэппинг из набора
        /// </summary>
        /// <param name="name">Имя удаляемого мэппинга</param>
        /// <returns>true если удаление успешное</returns>
        public static bool Remove(string name)
        {
            if (_mappings.ContainsKey(name))
                _mappings.Remove(name);
            else
                return false;

            return true;
        }
    }
}
