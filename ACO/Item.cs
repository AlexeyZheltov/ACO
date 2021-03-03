using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO
{
   class Item
    {
        /// <summary>
        ///  
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// 
        /// </summary>
        public string Number { get; set; }
        
        public string value { get; set; }
        /// <summary>
        /// Еденица измерения
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// Кол-во
        /// </summary>
        public string Amount { get; set; }
        /// <summary>
        /// Комментарий
        /// </summary>
        public string Note { get; set; }
    }
}
