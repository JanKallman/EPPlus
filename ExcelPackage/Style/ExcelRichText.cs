using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Drawing;

namespace OfficeOpenXml.Style
{
    public class ExcelRichText : XmlHelper
    {
        internal ExcelRichText(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
            SchemaNodeOrder=new string[] {"rPr", "t", "b", "i", "u", "sz", "color", "rFont", "family", "scheme", "charset"};
            PreserveSpace = false;
        }
        internal delegate void CallbackDelegate();
        CallbackDelegate _callback;
        internal void SetCallback(CallbackDelegate callback)
        {
            _callback=callback;
        }
        const string TEXT_PATH="d:t";
        public string Text 
        { 

            get
            {
                return GetXmlNodeString(TEXT_PATH);
            }
            set
            {
                SetXmlNodeString(TEXT_PATH, value);
                if (PreserveSpace)
                {
                    XmlElement elem = TopNode.SelectSingleNode(TEXT_PATH, NameSpaceManager) as XmlElement;
                    elem.SetAttribute("xml:space", "preserve");
                }
                if (_callback != null) _callback();
            }
        }
        bool _preserveSpace=false;
        public bool PreserveSpace
        {
            get
            {
                XmlElement elem = TopNode.SelectSingleNode(TEXT_PATH, NameSpaceManager) as XmlElement;
                if (elem != null)
                {
                    return elem.GetAttribute("xml:space")=="preserve";
                }
                return _preserveSpace;
            }
            set
            {
                XmlElement elem = TopNode.SelectSingleNode(TEXT_PATH, NameSpaceManager) as XmlElement;
                if (elem != null)
                {
                    if (value)
                    {
                        elem.SetAttribute("xml:space", "preserve");
                    }
                    else
                    {
                        elem.RemoveAttribute("xml:space");
                    }
                }
                _preserveSpace = false;
            }
        }
        const string BOLD_PATH = "d:rPr/d:b";
        public bool Bold
        {
            get
            {
                return GetXmlNodeBool(BOLD_PATH, false);
            }
            set
            {
                if (value)
                {
                    CreateNode(BOLD_PATH);
                }
                else
                {
                    DeleteNode(BOLD_PATH);
                }
                if(_callback!=null) _callback();
            }
        }
        const string ITALIC_PATH = "d:rPr/d:i";
        public bool Italic
        {
            get
            {
                return GetXmlNodeBool(ITALIC_PATH, false);
            }
            set
            {
                if (value)
                {
                    CreateNode(ITALIC_PATH);
                }
                else
                {
                    DeleteNode(ITALIC_PATH);
                }
                if (_callback != null) _callback();
            }
        }
        const string UNDERLINE_PATH = "d:rPr/d:u";
        public bool UnderLine
        {
            get
            {
                return GetXmlNodeBool(UNDERLINE_PATH, false);
            }
            set
            {
                if (value)
                {
                    CreateNode(UNDERLINE_PATH);
                }
                else
                {
                    DeleteNode(UNDERLINE_PATH);
                }
                if (_callback != null) _callback();
            }
        }
        const string SIZE_PATH = "d:rPr/d:sz/@val";
        public float Size
        {
            get
            {
                return Convert.ToSingle(GetXmlNodeDecimal(SIZE_PATH));
            }
            set
            {
                SetXmlNodeString(SIZE_PATH, value.ToString(ExcelWorksheet._ci));
                if (_callback != null) _callback();
            }
        }
        const string FONT_PATH = "d:rPr/d:rFont/@val";
        public string FontName
        {
            get
            {
                return GetXmlNodeString(FONT_PATH);
            }
            set
            {
                SetXmlNodeString(FONT_PATH, value);
                if (_callback != null) _callback();
            }
        }
        const string COLOR_PATH = "d:rPr/d:color/@rgb";
        public Color Color
        {
            get
            {
                string col = GetXmlNodeString(COLOR_PATH);
                if (col == "")
                {
                    return Color.Empty;
                }
                else
                {
                    return Color.FromArgb(int.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier));
                }
            }
            set
            {
                SetXmlNodeString(COLOR_PATH, value.ToArgb().ToString("X")/*.Substring(2, 6)*/);
                if (_callback != null) _callback();
            }
        }
    }
}
