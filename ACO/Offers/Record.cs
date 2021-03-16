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
                if (_Level==0)
                {                  
                    string[] numbers = Number.Split('.');
                    _Level = numbers.Length + 1;
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
                _Number =  _Number.Trim(new Char[] {' ', '.'});
            }
        }
        string _Number;

        public bool KeyEqual(List<string> keyFilds)
        {
            bool keyEqual = true;
            foreach (string key in keyFilds)
            {
                if (!KeyFilds.Contains(key)) return false;

            }
            return keyEqual;
        }

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
      

        /// <summary>
        ///  Библииотека заголовок/ значение
        /// </summary>
        public Dictionary<string,object> Values 
        {
            get
            {
                if (_Values == null)
                {
                    _Values = new Dictionary<string, object>();
                }
                return _Values;
            }
            set 
            {
                _Values = value;
            }
        }
        public Dictionary<string, object> _Values;
        
    }
}
