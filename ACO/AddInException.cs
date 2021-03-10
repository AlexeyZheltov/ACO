using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACO
{
    class AddInException : ApplicationException
    {
        public bool StopProcess { get;private set; }

        public AddInException(string message )
         : base  (message)
        {
           // Message = message;
        }

    }
}
