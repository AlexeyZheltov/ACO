using System.Collections.Generic;
using System.Linq;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

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
                WorkBook = _application.Workbooks.Open(path);

                // TODO Убрать
                _application.Visible = true;
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
            WorkBook?.Close();
            WorkBook = null;
            _wsCache = null;
        }

        /// <summary>
        /// Получает лист книги
        /// </summary>
        /// <remarks>Должен быть безопасен к исключениям.</remarks>
        /// <param name="sheetName">Имя листа</param>
        /// <returns>Лист книги</returns>
        public Excel.Worksheet GetSheet(string sheetName)
        {
            if (_wsCache == null)
                RefreshWSCache();

            if (_wsCache.TryGetValue(sheetName, out Excel.Worksheet worksheet))
                return worksheet;

            return null;
        }

        public Excel.Worksheet GetSheet(int index)
        {
            if (index <= WorkBook.Worksheets.Count && index > 0)
            {
                Excel.Worksheet worksheet = WorkBook.Worksheets[index];
                return worksheet;
            }
            throw new AddInException($"Лист {index}. Отсутствует");
        }

        /// <summary>
        /// Обновляет кэш страниц
        /// </summary>
        public void RefreshWSCache()
        {
            _wsCache = new Dictionary<string, Excel.Worksheet>();
            foreach (Excel.Worksheet ws in WorkBook.Sheets)
                _wsCache.Add(ws.Name, ws);
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
            _application = null;
        }

        /// <summary>
        /// Ускорение работы Excel за счет управления парметрами ScreenUpdating и DisplayAllert
        /// </summary>
        /// <param name="mode">true - включает ускорение</param>
        public static void Acselerate(bool mode)
        {
            _application.ScreenUpdating = !mode;
            _application.DisplayAlerts = !mode;
        }
    }
}
