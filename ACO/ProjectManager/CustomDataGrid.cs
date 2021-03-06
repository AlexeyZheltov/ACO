using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace ACO.ProjectManager
{
    class CustomDataGrid : DataGridView
    {
      public  CustomDataGrid() 
        {
            DoubleBuffered = true;
            SetColumns();
            RowHeadersVisible = false;
            SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            MultiSelect = false;
            AllowUserToAddRows = false;
            EditMode = DataGridViewEditMode.EditProgrammatically;
            
            //BackgroundColor = System.Drawing.Color.White;            
        }


        private void SetColumns()
        {
           
        }
    }
}
