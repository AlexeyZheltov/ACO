using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ACO
{
    class AddinWindow : IWin32Window
    {
        readonly Microsoft.Office.Interop.Excel.Window window;
        public AddinWindow(ThisAddIn thisAddIn)
        {
            window = thisAddIn.Application.ActiveWindow;// Windows[1];
        }
        public IntPtr Handle => (IntPtr)window.Hwnd;
    }
}
