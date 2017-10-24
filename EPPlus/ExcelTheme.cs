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


        // Sometimes black and white is stored as: 
        //<a:sysClr lastClr = "000000" val="windowText"/>
        //<a:sysClr lastClr = "FFFFFF" val="window"/>

        //And sometimes as:
        //<a:srgbClr val = "000000" />
        //<a:srgbClr val = "FFFFFF" />

        // but in GUI they are always the first two colors
        private void LoadFromDocument()
        {
            var nodes = _themeXml.GetElementsByTagName("a:srgbClr");

            _colors = new List<string>
            {
                "FFFFFF",
                "000000"
            };

            foreach (XmlElement n in nodes)
            {
                string hex = n.GetAttribute("val");

                if (hex != "FFFFFF" && hex != "000000")
                    _colors.Add(hex);
            }

            // In GUI, color 2 and 3 has switched position, compared to the xml file 
            _colors.Reverse(2, 2);
        }
    }
}
