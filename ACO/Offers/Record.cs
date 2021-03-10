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
        public string Level { get; set; }

        /// <summary>
        /// Номер пункта
        /// </summary>
        public string Number { get; set; }

        public Dictionary<string,string> Values { get; set; }
        
    }
}
