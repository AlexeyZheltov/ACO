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
            RowHeadersVisible = false;
            SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            MultiSelect = false;
            AllowUserToAddRows = false;
            RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            AllowUserToResizeRows = false;
           // EditMode = DataGridViewEditMode.EditProgrammatically;
           // RowsDefaultCellStyle.WrapMode = DataGridViewTriState.True;
           //BackgroundColor = System.Drawing.Color.White;
        }


       
    }
}
