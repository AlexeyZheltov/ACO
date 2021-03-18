using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO.Offers
{
    /// <summary>
    ///  Запсись КП
    /// </summary>
    class Record
    {
        /// <summary>
        /// Уровень
        /// </summary>
        public int Level
        {
            get
            {
                if (_Level == 0)
                {
                    _Level = string.IsNullOrEmpty(Number)? 0: Numbers.Length + 1;
                }
                return _Level;
            }
            private set
            {
                _Level = value;
            }
        }
        int _Level;

        private int myVar;

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


        //public bool KeyEqual(List<string> keyFilds)
        //{
        //    bool keyEqual = true;
        //    foreach (string key in keyFilds)
        //    {
        //        if (!KeyFilds.Contains(key)) return false;

        //    }
        //    return keyEqual;
        //}

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

        public bool Equal(Record recordPrint)
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
        ///  Сравнение
        /// </summary>
        /// <param name="recordPrint"></param>
        /// <returns></returns>
        public bool LevelEqual(Record recordPrint)
        {
            if (Number == recordPrint.Number) return true;
            if (Level != recordPrint.Level) return false;

            for (int i = 0; i < Numbers.Length - 1;)
            {
                if (Level > 1 && Level == recordPrint.Level)
                {
                    if (Numbers[i] != recordPrint.Numbers[i]) return false;
                }
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

        //   public int Index { get; internal set; }

        //public Dictionary<string, object> Values
        //{
        //    get
        //    {
        //        if (_Values == null)
        //        {
        //            _Values = new Dictionary<string, object>();
        //        }
        //        return _Values;
        //    }
        //    set
        //    {
        //        _Values = value;
        //    }
        //}
        //public Dictionary<string, object> _Values;
    }
}
