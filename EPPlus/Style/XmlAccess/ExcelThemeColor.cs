using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Style.XmlAccess
{
    public sealed class ExcelThemeColor : StyleXmlHelper
    {
        private string _hexcolor;

        internal ExcelThemeColor(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {
            _hexcolor = "";
        }

        internal override string Id => throw new NotImplementedException();

        internal override XmlNode CreateXmlNode(XmlNode top)
        {
            throw new NotImplementedException();
        }
    }
}
