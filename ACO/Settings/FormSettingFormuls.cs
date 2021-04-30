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
            }
        }
        FormulaAnalysis _Formula;

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
            if (Rbtn0.Checked) Formula = FormulaAnalysis.DeviationBasis;
            else if (Rbtn1.Checked) Formula = FormulaAnalysis.Avarage;
            else if (Rbtn2.Checked) Formula = FormulaAnalysis.Median;

            TopBound = double.TryParse(TBoxTop.Text, out double top) ? top : 0;
            BottomBound = double.TryParse(TBoxBottom.Text, out double bottom) ? bottom : 0;
            settings.Save();
            Close();
        }

        private void FormSettingFormuls_Load(object sender, EventArgs e)
        {
            Rbtn0.Checked = Formula == FormulaAnalysis.DeviationBasis;
            Rbtn1.Checked = Formula == FormulaAnalysis.Avarage;
            Rbtn2.Checked = Formula == FormulaAnalysis.Median;

            TBoxTop.Text = TopBound.ToString();
            TBoxBottom.Text = BottomBound.ToString();
        }

        private void Closing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) DialogResult = DialogResult.Cancel;
        }

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
          //  private void TBoxTextIsDigit(object sender, KeyPressEventArgs e)
          //  {

                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
                {
                    e.Handled = true;
                }
                //only allow one decimal point
                if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
                {
                    e.Handled = true;
                }
            //}
        }
    }
}
