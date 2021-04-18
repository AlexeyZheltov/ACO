using System;
using System.Collections.Generic;
using System.Drawing;
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
            List<ConditionFormat> listConditionFormats = new List<ConditionFormat>();
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
            List<ConditionFormat> listConditionFormats = new List<ConditionFormat>();


            listConditionFormats.Add(
                new ConditionFormat()
                {
                    ColumnName = "",
                    Operator = "Между",
                    FontName = "Tahoma",
                    FontSize = 10,
                    FontStyle = FontStyle.Regular,
                    ForeColor = Color.AliceBlue,
                    InteriorColor = Color.Yellow,
                    Formula1 = -0.1,
                    Formula2 = -0.15
                }
            );

            listConditionFormats.Add(
            new ConditionFormat()
            {
                ColumnName = "",
                Operator = "Меньше равно",
                FontName = "Tahoma",
                FontSize = 10,
                FontStyle = FontStyle.Regular,
                ForeColor = Color.White,
                InteriorColor = Color.Red,
                Formula1 = -0.15,
            }
            );

            listConditionFormats.Add(
                 new ConditionFormat()
                 {
                     ColumnName = "",
                     Operator = "Между",
                     FontName = "Tahoma",
                     FontSize = 10,
                     FontStyle = FontStyle.Regular,
                     ForeColor = Color.Red,
                     InteriorColor = Color.Yellow,
                     Formula1 = 0.1,
                     Formula2 = 0.15
                 }
            );

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

            XDocument xdoc = XDocument.Load(filename);
            XElement root = xdoc.Root;

            XElement xeConditions = root.Element("Conditions");
            foreach (XElement xeCondition in xeConditions.Elements() )
            {
                ConditionFormat conditionFormat = new ConditionFormat();
                conditionFormat.ColumnName = xeCondition.Attribute("ColumnName").Value;
                conditionFormat.Operator = xeCondition.Attribute("Operator").Value;
                conditionFormat.Formula1 = double.Parse(xeCondition.Attribute("Formula1").Value);
                conditionFormat.Formula2 = double.Parse(xeCondition.Attribute("Formula2").Value);

                XElement xeFormate = xeCondition.Element("Format");
                conditionFormat.FontName = xeFormate.Attribute("FontName").Value;
                conditionFormat.FontSize = float.Parse(xeFormate.Attribute("FontSize").Value);

                int styleNum = int.Parse(xeFormate.Attribute("FontStyle").Value);
                System.Drawing.FontStyle style = (FontStyle)styleNum ;
                //0;
                //if (styleNum == 0) style =  FontStyle.Regular;
                //if (styleNum == 1) style =  FontStyle.Bold;
                //if (styleNum == 2) style =  FontStyle.Italic;
                //if (styleNum == 8) style =  FontStyle.Strikeout;
                //if (styleNum == 4) style =  FontStyle.Underline;

                conditionFormat.FontStyle = style ;
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
            XElement xeCondition = new XElement("Condition");
            foreach (ConditionFormat conditionFormat in ListConditionFormats)
            {
                xeCondition.Add(new XAttribute("ColumnName", conditionFormat.ColumnName));
                xeCondition.Add(new XAttribute("Operator", conditionFormat.Operator));
                xeCondition.Add(new XAttribute("Formula1", conditionFormat.Formula1));
                xeCondition.Add(new XAttribute("Formula2", conditionFormat.Formula2));

                XElement xeFormate = new XElement("Format");
                string foreColor = conditionFormat.ForeColor.ToArgb().ToString();
                xeFormate.Add(new XAttribute("ForeColor", foreColor));
                string interiorColor = conditionFormat.InteriorColor.ToArgb().ToString();
                xeFormate.Add(new XAttribute("InteriorColor", interiorColor));

                xeFormate.Add(new XAttribute("FontName", conditionFormat.FontName));
                xeFormate.Add(new XAttribute("FontSize", conditionFormat.FontSize.ToString()));
                xeFormate.Add(new XAttribute("FontStyle", conditionFormat.FontStyle.ToString()));

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
