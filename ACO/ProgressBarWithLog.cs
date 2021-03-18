using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ACO
{
    public partial class ProgressBarWithLog : Form, IProgressBarWithLogUI
    {
        /// <summary>
        /// Возникает когда форма закрывается
        /// </summary>
        public event Action CloseForm;

        /// <summary>
        /// Указывает была ли нажат кнопка Отмена
        /// </summary>
        public bool IsAborted { get; set; } = false;

        public ProgressBarWithLog()
        {
            InitializeComponent();
        }

        private void OpenLogFolderButton_Click(object sender, EventArgs e)
        {
            string directoryPath = Path.Combine(
                                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                "Spectrum",
                                "ACO",
                                "Logs");

            if (!Directory.Exists(directoryPath)) Directory.CreateDirectory(directoryPath);

            Process.Start(directoryPath);
        }

        private void AbortButton_Click(object sender, EventArgs e) => IsAborted = true;

        private void ProgressBarWithLog_FormClosed(object sender, FormClosedEventArgs e) => CloseForm();

        /// <summary>
        /// Устанавливает максимальное значение для главного PB
        /// </summary>
        /// <param name="volume">Максимальное значение PB</param>
        public void SetMainBarVolum(int volume) => MainProgressBar.Maximum = volume;

        /// <summary>
        /// Устанавливает максимальное значение для вспомогательного PB
        /// </summary>
        /// <param name="volume">Максимальное значение PB</param>
        public void SetSubBarVolume(int volume)
        {
            Action action = () =>
            {
                SubProgressBar.Maximum = volume;
            };

            if (InvokeRequired)
                Invoke(action);
            else
                SubProgressBar.Maximum = volume;

        }

        /// <summary>
        /// Сбрасывает значениея главного PB в ноль
        /// </summary>
        public void ClearMainBar()
        {
            MainProgressBar.Value = 0;
            MainLabel.Text = "";
        }

        /// <summary>
        /// Сбрасывает значениея вспомогательного PB в ноль
        /// </summary>
        public void ClearSubBar()
        {
            Action action = () =>
            {
                SubProgressBar.Value = 0;
                SubLabel.Text = "";
            };

            if (InvokeRequired)
                Invoke(action);
            else
                action();
        }

        /// <summary>
        /// Делает приращение главного PB
        /// </summary>
        /// <param name="amount">Величина приращения, по умолчанию 1</param>
        public void MainBarTick(string text, int amount = 1)
        {
            Action action = () =>
            {
                MainProgressBar.Value += amount;
                MainLabel.Text = $"Этап {MainProgressBar.Value} из {MainProgressBar.Maximum}: {text}";
            };

            if (InvokeRequired)
                Invoke(action);
            else
                action();
        }

        /// <summary>
        /// Делает приращение вспомогательного PB
        /// </summary>
        /// <param name="amount">Величина приращения, по умолчанию 1</param>
        public void SubBarTick(int amount = 1)
        {
            Action action = () =>
            {
                SubProgressBar.Value += amount;
                SubLabel.Text = $"Обрабатывается строка: {SubProgressBar.Value} из {SubProgressBar.Maximum}";
            };

            if (InvokeRequired)
                Invoke(action);
            else
                action();
            //Invoke((Action)(() =>
            //{
            //    SubProgressBar.Value += amount;
            //    SubLabel.Text = $"{_subBarText}: {SubProgressBar.Value} из {SubProgressBar.Maximum}";
            //}));
        }

        /// <summary>
        /// Получить TextBox для вывода лога
        /// </summary>
        /// <returns></returns>
        public TextBox GetLogTextBox() => LogTextBox;

        void IProgressBarWithLogUI.ShowDialog()
        {
            Action action = () =>
            {
                ShowDialog();
            };
        }
    }
}
