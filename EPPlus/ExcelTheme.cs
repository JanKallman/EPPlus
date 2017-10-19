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
        XmlDocument _themeXml;
        ExcelWorkbook _wb;
        XmlNamespaceManager _nameSpaceManager;
        
        internal ExcelTheme(XmlNamespaceManager NameSpaceManager, XmlDocument xml, ExcelWorkbook wb) :
            base(NameSpaceManager, xml)
        {
            _themeXml = xml;
            _wb = wb;
            _nameSpaceManager = NameSpaceManager;
            SchemaNodeOrder = new string[] { "clrScheme" };
            LoadFromDocument();
        }

        private void LoadFromDocument()
        {
            var nodes = _themeXml.GetElementsByTagName("a:srgbClr");
         
            // didnt work, not sure why, even with a: in the path, then was bad token or something
            //XmlNode colorNode = _themeXml.SelectSingleNode(ThemeColorsPath, _nameSpaceManager);

            foreach (XmlElement n in nodes)
            {
                string hex = n.GetAttribute("val");

            }
        }
    }
}
