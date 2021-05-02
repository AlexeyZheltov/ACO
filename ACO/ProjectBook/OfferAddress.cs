using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO.ProjectBook
{
    public class OfferAddress
    {
        /// <summary>
        ///  Столбцы КП на листе анализ
        /// </summary>
    //    public int ColCost { get; set; }


      //  public int ColName { get; set; }

        /// <summary>
        ///  Отклонение по стоимости МАТ
        /// </summary>
        public int ColPercentMaterials { get; set; }

        /// <summary>
        /// Отклонение по стоимости РАБ
        /// </summary>
        public int ColPercentWorks { get; set; }
        /// <summary>
        /// Отклонение по стоимости
        /// </summary>
        public int ColPercentTotal { get; set; }

        /// <summary>
        ///  Стоимость базовой оценки
        /// </summary>
        public int ColTotalCost { get; set; }

        /// <summary>
        /// Комментарии к строкам
        /// </summary>
        public int ColComments { get; set; }
        /// <summary>
        /// Offer_start
        /// </summary>
        public int ColStartOffer { get; set; }
        /// <summary>
        ///  Наименование вида работ/ Offer_end
        /// </summary>
        public int ColStartOfferComments { get; set; }
        public string Name { get; set; }
    }
}
