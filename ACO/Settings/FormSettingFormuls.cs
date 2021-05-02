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
        readonly Properties.Settings settings = Properties.Settings.Default;
        
        //public static readonly FormulaAnalysis formula =FormulaAnalysis.
        public FormulaAnalysis FormulaCost
        {
            get
            {
                _FormulaCost = (FormulaAnalysis)settings.AnalysisFormulaCost;
                return _FormulaCost;
            }
            set
            {
                _FormulaCost = value;
                settings.AnalysisFormulaCost =(byte) _FormulaCost ;               
            }
        }
        FormulaAnalysis _FormulaCost ;

        public FormulaAnalysis FormulaCount
        {
            get
            {
                _FormulaCount = (FormulaAnalysis)settings.AnalysisFormulaCount;
                return _FormulaCount;
            }
            set
            {
                _FormulaCount = value;
                settings.AnalysisFormulaCount = (byte) _FormulaCount;
            }
        }
        FormulaAnalysis _FormulaCount;

        public double TopBound
        {
            get
            {
                _TopBound = settings.TopBoundAnalysis;
                return _TopBound;
            }
            set
            {
                _TopBound = value;
                settings.TopBoundAnalysis = _TopBound;               
            }
        }
        double _TopBound;

        public double BottomBound
        {
            get
            {
                _BottomBound = settings.BottomBoundAnalysis;
                return _BottomBound;
            }
            set
            {
                _BottomBound = value;
                settings.BottomBoundAnalysis = _BottomBound;
            }
        }
        double _BottomBound;

        public FormSettingFormuls()
        {
            InitializeComponent();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (RbtnBaseCost0.Checked) FormulaCost = FormulaAnalysis.DeviationBasis;
            else if (RbtnAvgCost1.Checked) FormulaCost = FormulaAnalysis.Avarage;
            else if (RbtnCostMedian2.Checked) FormulaCost = FormulaAnalysis.Median;

            if (RbtnBaseCount0.Checked) FormulaCount = FormulaAnalysis.DeviationBasis;
            else if (RbtnAvgCount1.Checked) FormulaCount = FormulaAnalysis.Avarage;
          
            TopBound = double.TryParse(TBoxTop.Text, out double top) ? top : 0;
            BottomBound = double.TryParse(TBoxBottom.Text, out double bottom) ? bottom : 0;
            settings.Save();
            Close();
        }

        private void FormSettingFormuls_Load(object sender, EventArgs e)
        {
            RbtnBaseCost0.Checked = FormulaCost == FormulaAnalysis.DeviationBasis;
            RbtnAvgCost1.Checked = FormulaCost == FormulaAnalysis.Avarage;
            RbtnCostMedian2.Checked = FormulaCost == FormulaAnalysis.Median;

            RbtnBaseCount0.Checked = FormulaCount == FormulaAnalysis.DeviationBasis;
            RbtnAvgCount1.Checked = FormulaCount == FormulaAnalysis.Avarage;

            TBoxTop.Text = TopBound.ToString();
            TBoxBottom.Text = BottomBound.ToString();
        }

        //private void Closing(object sender, FormClosingEventArgs e)
        //{
        //    if (e.CloseReason == CloseReason.UserClosing) DialogResult = DialogResult.Cancel;
        //}

        private void BtnCancel_Click(object sender, EventArgs e)
        {
        }

        private void TBoxTop_KeyPress(object sender, KeyPressEventArgs e)
        {
            /// <summary>
            ///  Проверка ввода Double
            /// </summary>
            /// <param name="sender"></param>
            /// <param name="e"></param>
            if (e.KeyChar == '.') e.KeyChar = ',';

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != ',') && (e.KeyChar != '-'))
                {
                    e.Handled = true;
                }
                //only allow one decimal point
                if ((e.KeyChar == ',') && ((sender as TextBox).Text.IndexOf(',') > -1) && ((sender as TextBox).Text.IndexOf('-') > -1))
                {
                    e.Handled = true;
                }
        }
      
    }
}
