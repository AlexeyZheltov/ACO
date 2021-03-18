using ACO.ProjectManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO.Offers
{
    class Field
    {
        string Header { get; set; }
        public int ColumnOffer { get; set; }
        //public int ColumnAnalysis { get; set; }
        public ColumnMapping ColumnAnalysis { get; set; }

        public Field()
        {

        }



    }
}
