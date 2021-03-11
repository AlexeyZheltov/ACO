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
            AllowUserToResizeColumns = true;
            AllowUserToResizeRows = false;
            RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
           // ReadOnly = false;
            EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            RowsDefaultCellStyle.WrapMode = DataGridViewTriState.False;// True;
            BackgroundColor = System.Drawing.Color.White;
        }


       
    }
}
