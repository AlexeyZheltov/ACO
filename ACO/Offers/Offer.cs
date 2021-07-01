using ACO.Offers;
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
       public Offer()
        {         
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
        public List<Record> Records 
        {
            get
            {
                if (_Records == null)
                {
                    _Records = new List<Record>();
                }
                return _Records;
            }
            set
            {
                _Records = value;
            } 
        }
        List<Record> _Records;
    }
}
