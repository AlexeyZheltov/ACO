using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO.ProjectBook
{
    /// <summary>
    ///  Столбцы КП на листе анализ
    /// </summary>
  public  class OfferColumns
    {
        /// <summary>
        /// Столбец перед началом КП / offer_start
        /// </summary>
        public int ColStartOffer { get; set; }

        /// <summary>
        ///  Уровень КП / offer_start +2
        /// </summary>
        public int ColLevelOffer { get; set; }

        /// <summary>
        ///  Столбец Наименование работ / КП / offer_start + 4
        /// </summary>
        public int ColNameOffer { get; set; }

        /// <summary>
        /// Ед. изм.  / КП / offer_start + 10
        /// </summary>
        public int ColUnitOffer { get; set; }
        /// <summary>
        ///  Кол-во  / КП / offer_start + 11
        /// </summary>
        public int ColCountOffer { get; set; }

        /// <summary>
        /// Цена материалы за ед.  / КП  / offer_start + 12
        /// </summary>
        public int ColCostMaterialsPerUnitOffer { get; set; }

        /// <summary>
        /// Цена материалы всего / КП  / offer_start + 13
        /// </summary>
        public int ColCostMaterialsTotalOffer { get; set; }

        /// <summary>
        ///  Цена работы за ед. / КП / offer_start + 14
        /// </summary>
        public int ColCostWorksPerUnitOffer { get; set; }

        /// <summary>
        ///  Цена работы всего / КП  / offer_start + 15
        /// </summary>
        public int ColCostWorksTotalOffer { get; set; }

        /// <summary>
        ///  Итого за ед. / КП  / offer_start + 16
        /// </summary>
        public int ColTotalCostPerUnitOffer { get; set; }

        /// <summary>
        ///  Итого / КП   / offer_start + 17
        /// </summary>
        public int ColCostTotalOffer { get; set; }

        /// <summary>
        ///  Примечание / КП   / offer_start + 18
        /// </summary>
        public int ColCommentOffer { get; set; }

        /// <summary>
        /// Сравнение наименований вида работ / offer_end
        /// </summary>
        public int ColStartOfferComments { get; set; }

        /// <summary>
        ///  Комментарии к описанию работ / offer_end + 1
        /// </summary>
        public int ColCommentsDescriptionWorks { get; set; }

        /// <summary>
        ///  Отклонение по объемам / offer_end + 2
        /// </summary>
        public int ColDeviationVolume { get; set; }

        /// <summary>
        /// Комментарии к объемам работ / offer_end + 3
        /// </summary>
        public int ColCommentsVolumeWorks { get; set; }

        /// <summary>
        /// Отклонение по стоимости / offer_end + 4
        /// </summary>
        public int ColDeviationCost { get; set; }

        /// <summary>
        /// Комментарии к стоимости работ / offer_end + 5
        /// </summary>
        public int ColCommentsCostWorks { get; set; }


        /// <summary>
        ///  Отклонение по стоимости МАТ  / offer_end + 6
        /// </summary>
        public int ColDeviationMaterials { get; set; }

        /// <summary>
        /// Отклонение по стоимости РАБ / offer_end + 7
        /// </summary>
        public int ColDeviationWorks { get; set; }

        /// <summary>
        /// Комментарии к строкам / offer_end + 8
        /// </summary>
        public int ColComments { get; set; }


    }
}
