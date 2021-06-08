using System.Collections.Generic;
using System.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ACO.ExcelHelpers
{
    /// <summary>
    /// Класс представляющий рабочую книгу
    /// </summary>
    /// <remarks>Основное назначение класса предоставить безопасный к исключениям фасад для работы с книгами Excel.
    /// Кеширования имен листов.
    /// Обязательно вызывать метод Init перед началом работы с классом и метод Finish в конце
    /// </remarks>
    class ExcelFile
    {
        static Excel.Application _application = null;

        readonly static string[] _excelExtensions =
        {
            ".xlsx",
            ".xlsm",
            ".xls"
        };

        Dictionary<string, Excel.Worksheet> _wsCache = null;

        /// <summary>
        /// Ссылка на открытую книгу
        /// </summary>
        public Excel.Workbook WorkBook { get; private set; }

        /// <summary>
        /// Открывает книгу
        /// </summary>
        /// <remarks>Так же выполняет проверки на существование книги. Обрабатывает исключения открытия</remarks>
        /// <returns>Книга</returns>
        public bool Open(string path)
        {
            if (_application == null)
                return false;

            if (!File.Exists(path))
                return false;

            if (!_excelExtensions.Contains(Path.GetExtension(path)))
                return false;

            bool result = false;
            try
            {
                WorkBook = _application.Workbooks.Open(path, UpdateLinks: false);

                // TODO Убрать
                //_application.Visible = true;
                result = true;
            }
            catch { }
            return result;
        }

        /// <summary>
        /// Закрывает книгу без сохранения
        /// </summary>
        public void Close()
        {
            WorkBook?.Close(false);
            WorkBook = null;
            _wsCache = null;
        }
        /// <summary>
        /// Инициализирует статические поля.
        /// </summary>
        public static void Init() => _application =new Excel.Application();

        /// <summary>
        /// Высвобождает статические поля
        /// </summary>
        public static void Finish()
        {
            _application.Quit();
            Marshal.ReleaseComObject(_application);
            _application = null;
        }
    }
}
