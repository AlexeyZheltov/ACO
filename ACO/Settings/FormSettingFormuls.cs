using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ACO.Settings
{
        public enum FormulaAnalysis
        {
            DeviationBasis =0 ,
            Avarage = 1 ,
            Median = 2
        }
    public partial class FormSettingFormuls : Form
    {
        readonly ACO.Properties.Settings settings = ACO.Properties.Settings.Default;
        
        //public static readonly FormulaAnalysis formula =FormulaAnalysis.
        public FormulaAnalysis Formula
        {
            get
            {
                _Formula = (FormulaAnalysis)settings.AnalysisFormula;
                return _Formula;
            }
            set
            {
                _Formula = value;
                settings.AnalysisFormula =(byte) _Formula;
                settings.Save();
            }
        }
        FormulaAnalysis _Formula;

        public FormSettingFormuls()
        {
            InitializeComponent();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (Rbtn0.Checked) Formula = FormulaAnalysis.DeviationBasis;
            else if (Rbtn1.Checked) Formula = FormulaAnalysis.Avarage;
            else if (Rbtn2.Checked) Formula = FormulaAnalysis.Median;
        }
    }
}
