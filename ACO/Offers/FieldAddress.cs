using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO.Offers
{

    /// <summary>
    ///  Адресация соответствующих полей в КП и Анализе
    /// </summary>
    class FieldAddress
    {
        string Header { get; set; }

        // Номер столбца в КП
        public int ColumnOffer { get; set; }        
        public int ColumnPaste{ get; set; }

        // Mаппинг столбца в Анализе
        public ColumnMapping MappingAnalysis { get; set; }

        public FieldAddress()
        {
        }
    }
}
