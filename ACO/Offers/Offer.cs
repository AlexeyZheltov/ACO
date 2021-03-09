using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO
{
    /// <summary>
    ///  Коммерческое предложение
    /// </summary>
    class Offer
    {
        public static readonly string SheetName = "Расчет по проекту";

       public Offer()
        {
            Items =new List<Item>();
        }

        /// <summary>
        ///  Проект
        /// </summary>
        public string ProjectName { get; set; }
        /// <summary>
        /// Заказчик
        /// </summary>
        public string Customer { get; set; }
        /// <summary>
        /// Номер проекта
        /// </summary>
        public string ProjectNumber { get; set; }
        /// <summary>
        ///  Дата
        /// </summary>
        public string Date { get; set; }

        /// <summary>
        /// Список строк 
        /// </summary>
        public List<Item> Items { get; set; }
      
    }
}
