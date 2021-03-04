using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ACO.ProjectManager
{
    public partial class FormManager : Form
    {
        //private ProjectManagerSave
        public FormManager()
        {
            InitializeComponent();
            LoadProjects();
        }

        private void LoadProjects()
        {
            ProjectManager manager = new ProjectManager();
            ProjectsTable.DataSource = manager.Projects;

            //foreach (Project project in manager.Projects)
            //{
            //}
        }

        private void customDataGrid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void PageColumns_Click(object sender, EventArgs e)
        {

        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void BtnAddProject_Click(object sender, EventArgs e)
        {
            string name = TboxProjectName.Text;
            new ProjectManager().CreateProject(name);
            LoadProjects();
        }

   

        private void PageProject_Click(object sender, EventArgs e)
        {

        }

        private void BtnAccept_Click(object sender, EventArgs e)
        {

        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {

        }


    }
}
