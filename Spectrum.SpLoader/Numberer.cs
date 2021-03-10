using System.Collections.Generic;
using System.Text;

namespace Spectrum.SpLoader
{
    /// <summary>
    /// Класс нумерации иерархии данных
    /// </summary>
    class Numberer
    {
        readonly List<int> _numbers;

        public Numberer() => _numbers = new List<int>() { 0 };

        private Numberer(List<int> numbers) => _numbers = new List<int>(numbers);

        /// <summary>
        /// Возвращает номер в виде строки
        /// </summary>
        public string GetNumber()
        {
            _numbers[_numbers.Count - 1]++;
            StringBuilder builder = new StringBuilder();
            foreach (int item in _numbers)
                builder.Append($"{item}.");

            builder.Remove(builder.Length - 1, 1);

            return builder.ToString();
        }

        /// <summary>
        /// Возвращет копию себя с добавленным уровнем нумерации
        /// </summary>
        public Numberer NextLevel()
        {
            Numberer nextLevel = new Numberer(_numbers);
            nextLevel._numbers.Add(0);
            return nextLevel;
        }
    }
}
