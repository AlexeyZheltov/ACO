using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ACO.ProjectManager;
using ACO.ExcelHelpers;

namespace ACO
{
    class ProjectWorkbook
    {
        Excel.Workbook _ProjectBook = Globals.ThisAddIn.Application.ActiveWorkbook;
        Project _project;
        public Excel.Worksheet AnalisysSheet 
        {
            get
            {
                if (_AnalisysSheet is null)
                {
                    _AnalisysSheet = ExcelHelper.GetSheet(_ProjectBook, _project.AnalysisSheetName);
                }
                return _AnalisysSheet;
            }
            set
            {
                _AnalisysSheet = value;
            }
        }
        Excel.Worksheet _AnalisysSheet;
       
        private ProjectWorkbook() 
        {
            _project = new ProjectManager.ProjectManager().ActiveProject;
        }
    }
}
