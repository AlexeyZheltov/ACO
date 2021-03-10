using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Spectrum.SpLoader
{
    /// <summary>
    /// Вспомогательный класс для валидации главной формы
    /// </summary>
    static class Validator
    {
        public enum Type
        {
            /// <summary>
            /// Валидация пути к файлу
            /// </summary>
            Path,
            /// <summary>
            /// Валидация имени столбца
            /// </summary>
            Column
        }

        class ValidItem
        {
            public TextBox TextBox { get; set; }
            public Type Type { get; set; }
            public bool Valid { get; set; }
        }

        readonly static List<ValidItem> _validatingList = new List<ValidItem>();
        static ErrorProvider _errorProvider;

        
        /// <summary>
        /// Установить провайдер ошибок для формы
        /// </summary>
        /// <param name="errorProvider"></param>
        public static void SetErrorProvider(ErrorProvider errorProvider) => _errorProvider = errorProvider;

        /// <summary>
        /// Добавить TExtBox в валидацию
        /// </summary>
        /// <param name="textBox"></param>
        /// <param name="type">Тип валидации</param>
        public static void Add(TextBox textBox, Type type) => _validatingList.Add(new ValidItem()
        { 
            TextBox = textBox,
            Type = type,
            Valid = false
        });

        /// <summary>
        /// Провести валидацию всех добавленных TextBox
        /// </summary>
        public static void Validate()
        {
            if (_errorProvider == null) return;

            foreach(ValidItem item in _validatingList)
            {
                if (String.IsNullOrEmpty(item.TextBox.Text))
                {
                    _errorProvider.SetError(item.TextBox, "Не может быть пустым");
                    item.Valid = false;
                }
                else if (item.Type == Type.Path && !File.Exists(item.TextBox.Text))
                {
                    _errorProvider.SetError(item.TextBox, "Файл не существует");
                    item.Valid = false;
                }
                else
                {
                    _errorProvider.SetError(item.TextBox, "");
                    item.Valid = true;
                }

            }
        }

        /// <summary>
        /// Проверить валидная ли форма
        /// </summary>
        /// <returns>true если форма валидная</returns>
        public static bool IsValid() => _validatingList.All(item => item.Valid);

        /// <summary>
        /// Проверить валиден ли маппинг
        /// </summary>
        /// <returns>true если маппинг валиден</returns>
        public static bool IsValidMapping() => _validatingList.FindAll(item => item.Type == Type.Column)
                                                              .All(item => item.Valid);
    }
}
