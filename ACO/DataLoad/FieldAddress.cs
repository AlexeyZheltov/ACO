using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO
{

    /// <summary>
    ///  Адресация соответствующих полей в КП и Анализе
    /// </summary>
    class FieldAddress
    {

        /// <summary>
        // Номер столбца в КП
        /// </summary>
        public int ColumnOffer { get; set; }
        /// <summary>
        /// Номер вставки столбца в Анализе
        /// </summary>
        public int ColumnPaste{ get; set; }

        // Mаппинг столбца в Анализе
        public ColumnMapping MappingAnalysis { get; set; }

        public FieldAddress()
        {
        }
    }
}
