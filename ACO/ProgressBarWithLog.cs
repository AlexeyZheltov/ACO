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
        //public event Action CloseForm;

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

        //  private void ProgressBarWithLog_FormClosed(object sender, FormClosedEventArgs e) => CloseForm();
        //private void ProgressBarWithLog_FormClosed(object sender, FormClosedEventArgs e)
        //{
        //    Action action = () =>
        //    {                
        //    CloseForm();
        //    };

        //    if (InvokeRequired)
        //        Invoke(action);
        //    else
        //        action();
        //}

        /// <summary>
        /// Устанавливает максимальное значение для главного PB
        /// </summary>
        /// <param name="volume">Максимальное значение PB</param>
        public void SetMainBarVolum(int volume)
        {
            Action action = () =>
            {
                MainProgressBar.Maximum = volume;
            };

            if (InvokeRequired)
                Invoke(action);
            else
                SubProgressBar.Maximum = volume;
        }
        /// <summary>
        /// Устанавливает максимальное значение для вспомогательного PB
        /// </summary>
        /// <param name="volume">Максимальное значение PB</param>
        public void SetSubBarVolume(int volume)
        {
            Action action = () =>
            {
                SetCount(volume);
                SubProgressBar.Maximum = volume;// SubProgressBar.Maximum = volume;
            };

            if (InvokeRequired)
                Invoke(action);
            else
                SetCount(volume);
            SubProgressBar.Maximum = volume; //SubProgressBar.Maximum = volume;
        }

        /// <summary>
        /// Сбрасывает значениея главного PB в ноль
        /// </summary>
        public void ClearMainBar()
        {
            Action action = () =>
            {
                MainProgressBar.Value = 0;
                MainLabel.Text = "";
            };
            if (InvokeRequired)
                Invoke(action);
            else
                action();
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

                LogTextBox.Text += text;
                LogTextBox.Text += Environment.NewLine;
            };

            if (InvokeRequired)
                Invoke(action);
            else
                action();
        }

        int _value = 0;
        int _k = 1;
        private void Tick(int amount = 1)
        {
            _value += amount;
            if (_k == 1 || _value == _Count)
            {
                SubProgressBar.Value = _value;
                SubLabel.Text = $"Этап {_value} из {_Count}";
                if (_value == 1)
                {
                    _k = -_Step + 1;

                }
                else
                {
                    _k = -_Step;
                }
            }
            _k++;
        }

        /// <summary>
        /// Делает приращение вспомогательного PB
        /// </summary>
        /// <param name="amount">Величина приращения, по умолчанию 1</param>
        public void SubBarTick(int amount = 1)
        {
            Action action = () =>
            {
                Tick(amount);
                // SubProgressBar.Value += amount;
                // SubLabel.Text = $"Обрабатывается строка: {SubProgressBar.Value} из {SubProgressBar.Maximum}";
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
        int _Step = 1;
        int _Count = 0;
        private void SetCount(int count)
        {
            _value = 0;
            _Count = count;
            if (count > 1000)
            {
                _Step = 99;
            }
            else if (count > 100)
            {
                _Step = 9;
            }
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
            if (InvokeRequired)
                Invoke(action);
            else
                action();
        }

        public void CloseFrm()
        {
            Action action = () =>
            {
                Close();
            };

            if (InvokeRequired)
                Invoke(action);
            else
                action();
        }

        public void Writeline(string message)
        {
            Action action = () =>
            {
                LogTextBox.Text += message;
                LogTextBox.Text += Environment.NewLine;
            };

            if (InvokeRequired)
                Invoke(action);
            else
                action();
        }
    }
}
