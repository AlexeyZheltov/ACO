using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO
{
    /// <summary>
    /// Иеерархический элемент
    /// </summary>
    class HItem
    {
        readonly List<HItem> _items = new List<HItem>();
        HItem _last;
        int _sub_level;

        /// <summary>
        /// Уровень
        /// </summary>
        public int Level { get; set; } = 0;

        /// <summary>
        /// Строка на листе с которой был считан эллемент
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// Номер
        /// </summary>
        public string Number { get; set; }

        /// <summary>
        /// Добавляет эллемент в иерархию
        /// </summary>
        /// <param name="item"></param>
        public void Add(HItem item)
        {
            if(_items.Count == 0)
                _sub_level = item.Level;

            if (_sub_level == item.Level)
            {
                _last = item;
                _items.Add(_last);
            }
            else
                _last.Add(item);
        }

        public List<HItem> GetSubLevel() => _items;

        /// <summary>
        /// Нумерация
        /// </summary>
        /// <param name="numberer"></param>
        /// <param name="pb"></param>
        public void Numeric(Numberer numberer, IProgressBarWithLogUI pb)
        {
            foreach(var item in _items)
            {
                if (pb?.IsAborted ?? false) break;
                pb?.SubBarTick();
                item.Number = numberer.GetNextNumber();
                item.Numeric(numberer.GetNextLevel(), pb);
            }
        }

        /// <summary>
        /// Общее число эллементов в иерархии
        /// </summary>
        /// <returns></returns>
        public int AllCount()
        {
            int count = _items.Count;
            foreach (var item in _items)
                count += item.AllCount();

            return count;
        }

        /// <summary>
        /// Итератор по иерархии
        /// </summary>
        /// <returns></returns>
        public IEnumerable<HItem> Items()
        {
            foreach(var item in _items)
            {
                yield return item;

                foreach (var sub_item in item.Items())
                    yield return sub_item;
            }
        }
    }

    static class ListExtension
    {
        /// <summary>
        /// Является ли последовательность непрерывной, т.е. (item).Row == (item + 1).Row + 1
        /// </summary>
        /// <param name="items"></param>
        /// <returns></returns>
        public static bool IsSolid(this List<HItem> items)
        {
            var s_items = items.OrderBy(x => x.Row).ToList();

            for(int ptr = 0; ptr < s_items.Count - 1; ptr++)
                if (s_items[ptr].Row != s_items[ptr + 1].Row + 1)
                    return false;

            return true;
        }
    }
}
