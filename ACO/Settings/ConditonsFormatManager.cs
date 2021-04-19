using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ACO
{
    public class ConditonsFormatManager
    {
        public List<ConditionFormat> ListConditionFormats
        {
            get
            {
                if (_ListConditionFormats == null)
                {
                    _ListConditionFormats = GetFromXML();

                }
                return _ListConditionFormats;
            }
            set
            {
                _ListConditionFormats = value;
            }
        }

        private List<ConditionFormat> GetFromXML()
        {
            List<ConditionFormat> listConditionFormats;
            string filename = GetPath();
            if (File.Exists(filename))
            {
                listConditionFormats = GetConditionsFromXml(filename);
            }
            else
            {
                listConditionFormats = GetDefault();
            }
            return listConditionFormats;
        }



        private List<ConditionFormat> GetDefault()
        {

            List<ConditionFormat> listConditionFormats = new List<ConditionFormat>
            {
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationCost],
                    Operator = "Содержит",
                    FontBold = false,
                    ForeColor = Color.Red,
                    InteriorColor = Color.White,
                    Text = "#НД"
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationCost],
                    Operator = "Меньше",
                    FontBold = false,
                    ForeColor = Color.Red,
                    InteriorColor = Color.Yellow,
                    Formula1 = -0.15
                },

                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationCost],
                    Operator = "Между",
                    FontBold = false,
                    ForeColor = Color.Brown,
                    InteriorColor = Color.LightYellow,
                    Formula1 = -0.15,
                    Formula2 = -0.05
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationCost],
                    Operator = "Между",
                    FontBold = false,
                    ForeColor = Color.Black,
                    InteriorColor = Color.LightPink,
                    Formula1 = 0.05,
                    Formula2 = 0.15
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationCost],
                    Operator = "Больше",
                    FontBold = false,
                    ForeColor = Color.White,
                    InteriorColor = Color.Red,
                    Formula1 = 0.15
                },

                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationMat],
                    Operator = "Содержит",
                    FontBold = false,
                    ForeColor = Color.Red,
                    InteriorColor = Color.White,
                    Text = "#НД"
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationMat],
                    Operator = "Меньше",
                    FontBold = false,
                    ForeColor = Color.Red,
                    InteriorColor = Color.Yellow,
                    Formula1 = -0.15
                },

                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationMat],
                    Operator = "Между",
                    FontBold = false,
                    ForeColor = Color.Brown,
                    InteriorColor = Color.LightYellow,
                    Formula1 = -0.15,
                    Formula2 = -0.05
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationMat],
                    Operator = "Между",
                    FontBold = false,
                    ForeColor = Color.Black,
                    InteriorColor = Color.LightPink,
                    Formula1 = 0.05,
                    Formula2 = 0.15
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationMat],
                    Operator = "Больше",
                    FontBold = false,
                    ForeColor = Color.White,
                    InteriorColor = Color.Red,
                    Formula1 = 0.15
                },

                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationWorks],
                    Operator = "Содержит",
                    FontBold = false,
                    ForeColor = Color.Red,
                    InteriorColor = Color.White,
                    Text = "#НД"
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationWorks],
                    Operator = "Меньше",
                    FontBold = false,
                    ForeColor = Color.Red,
                    InteriorColor = Color.Yellow,
                    Formula1 = -0.15
                },

                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationWorks],
                    Operator = "Между",
                    FontBold = false,
                    ForeColor = Color.Brown,
                    InteriorColor = Color.LightYellow,
                    Formula1 = -0.15,
                    Formula2 = -0.05
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationWorks],
                    Operator = "Между",
                    FontBold = false,
                    ForeColor = Color.Black,
                    InteriorColor = Color.LightPink,
                    Formula1 = 0.05,
                    Formula2 = 0.15
                },
                new ConditionFormat()
                {
                    ColumnName = ListAnalysis.ColumnCommentsValues[StaticColumnsComments.DeviationWorks],
                    Operator = "Больше",
                    FontBold = false,
                    ForeColor = Color.White,
                    InteriorColor = Color.Red,
                    Formula1 = 0.15
                },

                new ConditionFormat()
                {
                    ColumnName = "Выделение",
                    Operator = "Содержит",
                    FontBold = false,
                    ForeColor = Color.Red,
                    InteriorColor = Color.White,
                    Text = "#НД"
                }
            };
            return listConditionFormats;
        }
        List<ConditionFormat> _ListConditionFormats;

        private string GetPath()
        {
            string folder = GetFolderSettings();
            string filename = Path.Combine(folder, "condition_format.xml");
            return filename;
        }
        public static string GetFolderSettings()
        {
            string path = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Spectrum",
            "ACO"
            );
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }

        private List<ConditionFormat> GetConditionsFromXml(string filename)
        {
            List<ConditionFormat> listConditionFormats = new List<ConditionFormat>();
            IFormatProvider formatter = new NumberFormatInfo { NumberDecimalSeparator = "," };
            XDocument xdoc = XDocument.Load(filename);
            XElement root = xdoc.Root;

            XElement xeConditions = root.Element("Conditions");
            foreach (XElement xeCondition in xeConditions.Elements())
            {
                ConditionFormat conditionFormat = new ConditionFormat
                {
                    ColumnName = xeCondition.Attribute("ColumnName").Value,
                    Operator = xeCondition.Attribute("Operator").Value
                };
                string formula1 = xeCondition.Attribute("Formula1").Value;

                conditionFormat.Text = xeCondition.Attribute("Text").Value;
                conditionFormat.Formula1 = double.Parse(formula1, formatter);
                string formula2 = xeCondition.Attribute("Formula2").Value;
                conditionFormat.Formula2 = double.Parse(formula2, formatter);
                XElement xeFormate = xeCondition.Element("Format");

                conditionFormat.FontBold = bool.Parse(xeFormate.Attribute("FontBold").Value);
                conditionFormat.ForeColor = Color.FromArgb(int.Parse(xeFormate.Attribute("ForeColor").Value));
                conditionFormat.InteriorColor = Color.FromArgb(int.Parse(xeFormate.Attribute("InteriorColor").Value));

                listConditionFormats.Add(conditionFormat);
            }
            return listConditionFormats;
        }

        public void Save()
        {

            XElement root = new XElement("Settings");
            XElement xeConditions = new XElement("Conditions");
            foreach (ConditionFormat conditionFormat in ListConditionFormats)
            {
                XElement xeCondition = new XElement("Condition");
                xeCondition.Add(new XAttribute("ColumnName", conditionFormat.ColumnName));
                xeCondition.Add(new XAttribute("Operator", conditionFormat.Operator));
                xeCondition.Add(new XAttribute("Text", conditionFormat.Text ?? ""));
                xeCondition.Add(new XAttribute("Formula1", conditionFormat.Formula1.ToString()));
                xeCondition.Add(new XAttribute("Formula2", conditionFormat.Formula2.ToString()));

                XElement xeFormate = new XElement("Format");
                string foreColor = conditionFormat.ForeColor.ToArgb().ToString();
                xeFormate.Add(new XAttribute("ForeColor", foreColor));
                string interiorColor = conditionFormat.InteriorColor.ToArgb().ToString();
                xeFormate.Add(new XAttribute("InteriorColor", interiorColor));
                xeFormate.Add(new XAttribute("FontBold", conditionFormat.FontBold.ToString()));

                xeCondition.Add(xeFormate);
                xeConditions.Add(xeCondition);
            }
            root.Add(xeConditions);

            XDocument xdoc = new XDocument(root);
            string fileName = GetPath();
            xdoc.Save(fileName);
        }
    }
}
