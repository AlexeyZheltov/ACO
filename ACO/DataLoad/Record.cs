using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO
{
    /// <summary>
    ///  Запсись КП
    /// </summary>
    class Record
    {
        /// <summary>
        /// Уровень номера
        /// </summary>
        public int Level
        {
            get
            {
                if (_Level == 0)
                {                  
                    _Level = Numbers.Length ;
                }
                return _Level;
            }
            private set
            {
                _Level = value;
            }
        }
        int _Level;


        /// <summary>
        ///  Массив номеров пункта.
        /// </summary>
        public string[] Numbers
        {
            get
            {
                if (_numbers is null)
                {
                    _numbers = Number.Split('.');
                }
                return _numbers;
            }
            set { _numbers = value; }
        }
        string[] _numbers;




        /// <summary>
        /// Номер пункта
        /// </summary>
        public string Number
        {
            get
            {
                return _Number;
            }
            set
            {
                _Number = value;
                _Number = _Number.Trim(new Char[] { ' ', '.' });
            }
        }
        string _Number;

        /// <summary>
        ///  Поля отмеченные как "Проверять"
        /// </summary>
        public List<string> KeyFilds
        {
            get
            {
                if (_KeyFilds == null)
                {
                    _KeyFilds = new List<string>();
                }
                return _KeyFilds;
            }
            set
            {
                _KeyFilds = value;
            }
        }
        List<string> _KeyFilds;

        public bool KeyEqual(Record recordPrint)
        {
            foreach (string keyFild in KeyFilds)
            {
                if (!recordPrint.KeyFilds.Contains(keyFild))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        ///  Сравнение уровней номеров 2записей
        /// </summary>
        /// <param name="recordAdd"> </param>
        /// <returns></returns>
        public bool LevelEqual(Record recordAdd)
        {
            //if (Number == recordAdd.Number) return true;
            if (Level != recordAdd.Level) return false;

            if (Level == 1 || Level != recordAdd.Level) return false;
            for (int i = 0; i < Numbers.Length - 1; i++)
            {                
               if (Numbers[i] != recordAdd.Numbers[i]) return false;             
            }
            return true;
        }
        /// <summary>
        ///  Библииотека заголовок/ значение
        /// </summary>
        public List<FieldAddress> Addresslist
        {
            get
            {
                if (_Addresslist == null)
                {
                    _Addresslist = new List<FieldAddress>();
                }
                return _Addresslist;
            }
            set
            {
                _Addresslist = value;
            }
        }
        public List<FieldAddress> _Addresslist;

        /// <summary>
        ///  Значения для вывода на лист (столбец вывода \ Значение ячейки)
        /// </summary>
        public Dictionary<int, object> Values
        {
            get
            {
                if (_Values == null)
                {
                    _Values = new Dictionary<int, object>();
                }
                return _Values;
            }
            set
            {
                _Values = value;
            }
        }
        public Dictionary<int, object> _Values;

    }
}
