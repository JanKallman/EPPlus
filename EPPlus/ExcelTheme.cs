using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml
{
    public sealed class ExcelTheme : XmlHelper
    {
        const string ThemeColorsPath = @"d:theme/d:themeElements/d:clrScheme";
        private List<string> _colors;
        XmlNamespaceManager _nameSpaceManager;
        XmlDocument _themeXml;
        ExcelWorkbook _wb;

        internal ExcelTheme(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb) :
            base(NameSpaceManager, xml)
        {
            _themeXml = xml;
            _wb = wb;
            _nameSpaceManager = NameSpaceManager;
            SchemaNodeOrder = new string[] { "clrScheme" };
            LoadFromDocument();
        }

        public IEnumerable<string> Colors => _colors;



        private void LoadFromDocument()
        {

            var themeColorNodes = _themeXml.GetElementsByTagName("a:srgbClr");

            _colors = new List<string>();

            if (themeColorNodes.Count == 10)
            {
                _colors.Add("000000");
                _colors.Add("FFFFFF");
            }

            foreach (XmlElement n in themeColorNodes)
            {
                string hex = n.GetAttribute("val");
                _colors.Add("FFFFFF");
            }

            // In GUI, color 0 and 1, and 2 and 3 has switched position, compared to the xml file 
            _colors.Reverse(0, 2);
            _colors.Reverse(2, 2);
        }
    }
}
