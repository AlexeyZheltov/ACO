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
        ///  Наименование работ
        /// </summary>
        public string NameWork { get; set; }  
        
        /// <summary>
        /// 
        /// </summary>
        public string Number { get; set; }
        
        /// <summary>
        /// Еденица измерения
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// Кол-во
        /// </summary>
        public string Amount { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public string Value { get; set; }
        /// <summary>
        /// Общая стоимость 
        /// </summary>
        public string CostTotal { get; set; }
        /// <summary>
        /// Цена за штуку
        /// </summary>
        public string PricePerPiece { get; set; }
        /// <summary>
        /// Стоимость материалов за штуку
        /// </summary>
        public string CostMaterialsPerPiece { get; set; }
        /// <summary>
        /// Стоимость материалов общая 
        /// </summary>
        public string CostMaterialsTotal { get; set; }

        /// <summary>
        /// Стоимость работ за единицу
        /// </summary>
        public string CostWorksPerPiece { get; set; }
        /// <summary>
        /// Стоимость работ общая
        /// </summary>
        public string CostWorksTotal { get; set; }
        /// <summary>
        /// Комментарий
        /// </summary>
        public string Note { get; set; }
    }
}
