using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO
{
    /// <summary>
    ///  Строка КП   
    /// </summary>
    class Item
    {
        /// <summary>
        ///  Строка
        /// </summary>
        public int Row { get; internal set; }
        
        /// <summary>
        /// Заголовок
        /// </summary>
        public string Header { get; set; }
        
        /// <summary>
        /// Уровень
        /// </summary>
        public string Level { get; set; }

        /// <summary>
        /// Номер пункта
        /// </summary>
        public string Number { get; set; }

        /// <summary>
        /// Значение 
        /// </summary>
        public string Value { get; set; }
    }
}
