using System;
using System.Windows.Forms;

namespace ACO
{
    /// <summary>
    /// Интерфейс взаимодействия с окном прогрессбара
    /// </summary>
    interface IProgressBarWithLogUI
    {

        /// <summary>
        /// Возникает когда форма закрывается
        /// </summary>
        event Action CloseForm;

        /// <summary>
        /// Указывает была ли нажат кнопка Отмена
        /// </summary>
        bool IsAborted { get; set; }

        /// <summary>
        /// Устанавливает максимальное значение для главного PB
        /// </summary>
        /// <param name="volume">Максимальное значение PB</param>
        void SetMainBarVolum(int volume);

        /// <summary>
        /// Устанавливает максимальное значение для вспомогательного PB
        /// </summary>
        /// <param name="volume">Максимальное значение PB</param>
        void SetSubBarVolume(int volume);

        /// <summary>
        /// Сбрасывает значениея главного PB в ноль
        /// </summary>
        void ClearMainBar();

        /// <summary>
        /// Сбрасывает значениея вспомогательного PB в ноль
        /// </summary>
        void ClearSubBar();

        /// <summary>
        /// Делает приращение главного PB
        /// </summary>
        /// <param name="amount">Величина приращения, по умолчанию 1</param>
        void MainBarTick(string text, int amount = 1);

        /// <summary>
        /// Делает приращение вспомогательного PB
        /// </summary>
        /// <param name="amount">Величина приращения, по умолчанию 1</param>
        void SubBarTick(int amount = 1);

        /// <summary>
        /// Получить TextBox для вывода лога
        /// </summary>
        /// <returns></returns>
        TextBox GetLogTextBox();

        void CloseFrm();
        void ShowDialog();
        void Show();
        void Show(IWin32Window window);
    }
}
