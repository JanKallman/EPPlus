using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing;
using System.Drawing;

namespace OfficeOpenXml.Style
{
    public enum eUnderLineType
    {
        Dash,
        DashHeavy,
        DashLong, 
        DashLongHeavy,
        Double,
        DotDash,
        DotDashHeavy,
        DotDotDash,
        DotDotDashHeavy,
        Dotted,
        DottedHeavy,
        Heavy,
        None,
        Single,
        Wavy,
        WavyDbl,
        WavyHeavy,
        Words
    }
    public enum eStrikeType
    {
        Double,
        No,
        Single
    }


    /// <summary>
    /// Used by Rich-text and Paragraphs.
    /// </summary>
    public class ExcelTextFont : XmlHelper
    {
        string _path;
        XmlNode _rootNode;
        public ExcelTextFont(XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder)
            : base(namespaceManager, rootNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _rootNode = rootNode;
            XmlNode node = rootNode.SelectSingleNode(path, namespaceManager);
            if (node != null)
            {
                TopNode = node;
            }
            _path = path;
        }
        string _fontLatinPath = "a:latin/@typeface";
        public string LatinFont
        {
            get
            {
                return GetXmlNode(_fontLatinPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNode(_fontLatinPath, value);
            }
        }

        private void CreateTopNode()
        {
            if (TopNode == _rootNode)
            {
                CreateNode(_path);
                TopNode = _rootNode.SelectSingleNode(_path, NameSpaceManager);
            }
        }
        string _fontCsPath = "a:cs/@typeface";
        public string ComplexFont
        {
            get
            {
                return GetXmlNode(_fontCsPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNode(_fontCsPath, value);
            }
        }
        string _boldPath = "@b";
        public bool Bold
        {
            get
            {
                return GetXmlNodeBool(_boldPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNode(_boldPath, value ? "1" : "0");
            }
        }
        string _underLinePath = "@u";
        public eUnderLineType UnderLine
        {
            get
            {
                return TranslateUnderline(GetXmlNode(_underLinePath));
            }
            set
            {
                CreateTopNode();
                SetXmlNode(_underLinePath, TranslateUnderlineText(value));
            }
        }
        string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";
        public Color UnderLineColor
        {
            get
            {
                string col = GetXmlNode(_underLineColorPath);
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
                CreateTopNode();
                SetXmlNode(_underLineColorPath, value.ToArgb().ToString("X").Substring(2, 6));
            }
        }
        string _italicPath = "@i";
        public bool Italic
        {
            get
            {
                return GetXmlNodeBool(_italicPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNode(_italicPath, value ? "1" : "0");
            }
        }
        string _strikePath = "@strike";
        public eStrikeType Strike
        {
            get
            {
                return TranslateStrike(GetXmlNode(_strikePath));
            }
            set
            {
                CreateTopNode();
                SetXmlNode(_strikePath, TranslateStrikeText(value));
            }
        }
        string _sizePath = "@sz";
        public float Size
        {
            get
            {
                return GetXmlNodeInt(_sizePath) / 100;
            }
            set
            {
                CreateTopNode();
                SetXmlNode(_sizePath, ((int)(value * 100)).ToString());
            }
        }
        string _colorPath = "a:solidFill/a:srgbClr/@val";
        public Color Color
        {
            get
            {
                string col = GetXmlNode(_colorPath);
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
                CreateTopNode();
                SetXmlNode(_colorPath, value.ToArgb().ToString("X").Substring(2, 6));
            }
        }
        #region "Translate methods"
        private eUnderLineType TranslateUnderline(string text)
        {
            switch (text)
            {
                case "sng":
                    return eUnderLineType.Single;
                case "dbl":
                    return eUnderLineType.Double;
                default:
                    return (eUnderLineType)Enum.Parse(typeof(eUnderLineType), text);
            }
        }
        private string TranslateUnderlineText(eUnderLineType value)
        {
            switch (value)
            {
                case eUnderLineType.Single:
                    return "sng";
                case eUnderLineType.Double:
                    return "dbl";
                default:
                    string ret = value.ToString();
                    return ret.Substring(0, 1).ToLower() + ret.Substring(1, ret.Length - 1);
            }
        }
        private eStrikeType TranslateStrike(string text)
        {
            switch (text)
            {
                case "dblStrike":
                    return eStrikeType.Double;
                case "sngStrike":
                    return eStrikeType.Single;
                default:
                    return eStrikeType.No;
            }
        }
        private string TranslateStrikeText(eStrikeType value)
        {
            switch (value)
            {
                case eStrikeType.Single:
                    return "sngStrike";
                case eStrikeType.Double:
                    return "dblStrike";
                default:
                    return "noStrike";
            }
        }
        #endregion
        public void SetFromFont(Font Font)
        {
            LatinFont = Font.Name;
            ComplexFont = Font.Name;
            Size = Font.Size;
            if (Font.Bold) Bold = Font.Bold;
            if (Font.Italic) Italic = Font.Italic;
            if (Font.Underline) UnderLine = eUnderLineType.Single;
            if (Font.Strikeout) Strike = eStrikeType.Single;
        }
    }
}
