using System.Collections.Generic;

namespace Spectrum.SpLoader.XMLSetting
{
    /// <summary>
    /// Файл работы настройками программы
    /// </summary>
    static class SettingManager
    {
        /// <summary>
        /// Путь к файлу шаблона
        /// </summary>
        public static string TemplatePath { get; set; }

        /// <summary>
        /// Путь к omni файлу
        /// </summary>
        public static string OmniPath { get; set; }

        /// <summary>
        /// Текущий номер лог файла
        /// </summary>
        public static int CurrentLogNumber { get; set; }

        /// <summary>
        /// Загружает настройки из файла
        /// </summary>
        public static void Load()
        {
            List<string> buffer = XMLManager.ReadSetting();
            if ((buffer?.Count ?? 0) == 0) return;
            TemplatePath = buffer[0];
            OmniPath = buffer[1];
            int.TryParse(buffer[2], out int result);
            CurrentLogNumber = result;
        }

        /// <summary>
        /// Сохраняет настройки в файл
        /// </summary>
        public static void Save()
        {
            XMLManager.Save(TemplatePath, OmniPath, CurrentLogNumber.ToString());
        }
    }
}
