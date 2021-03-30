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
        public int ColPercentMaterial { get; set; }
        public int ColPercentWorks { get; set; }
        public int ColPercentTotal { get; set; }
        public int ColTotalCost { get; set; }
        public int ColComments { get; set; }
        public int ColStartOffer { get; set; }
        public int ColStartOfferComments { get; set; }
    }
}
