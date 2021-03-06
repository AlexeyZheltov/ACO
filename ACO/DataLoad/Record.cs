﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ACO
{
    /// <summary>
    ///  Запись КП
    /// </summary>
    class Record
    {
        /// <summary>
        /// Уровень 
        /// </summary>
        public int Level { get; set; }
      
        /// <summary>
        ///  Массив номеров пункта.
        /// </summary>
        public string[] Numbers
        {
            get
            {
                if (_numbers is null && !string.IsNullOrEmpty(Number))
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

        public bool IsEmpty()
        {
            bool empty = Level>0 ;
            foreach (string field in KeyFilds)
            {
               if (!string.IsNullOrEmpty(field)) return false;
            }
            return empty;
        }

        private static string[] replaceArr = { "  ", "\r", "\n"};

        /// <summary>
        ///  Сравнение проверяемых полей 
        /// </summary>
        /// <param name="recordPrint"></param>
        /// <returns></returns>
        public bool KeyEqual(Record recordPrint)
        {
            foreach (string keyFild in KeyFilds)
            {
                bool exist = false;
                foreach (string recordField in recordPrint.KeyFilds)
                {                    
                   if ( keyFild == recordField ) { exist = true; }
                   string text1 = keyFild.Trim().ToLower();
                    string text2 = recordField.Trim().ToLower();
                 foreach(string delStr in replaceArr)
                    {
                        text1 = text1.Replace(delStr, "");
                        text2 = text2.Replace(delStr, "");
                    }

                    if (text1 == text2)
                    {
                        exist = true;
                    }
                }
                if (!exist) return false;
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
