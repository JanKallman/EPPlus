/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing;
using System.Drawing;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Linestyle
    /// </summary>
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
    /// <summary>
    /// Type of font strike
    /// </summary>
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
        internal ExcelTextFont(XmlNamespaceManager namespaceManager, XmlNode rootNode, string path, string[] schemaNodeOrder)
            : base(namespaceManager, rootNode)
        {
            SchemaNodeOrder = schemaNodeOrder;
            _rootNode = rootNode;
            if (path != "")
            {
                XmlNode node = rootNode.SelectSingleNode(path, namespaceManager);
                if (node != null)
                {
                    TopNode = node;
                }
            }
            _path = path;
        }
        string _fontLatinPath = "a:latin/@typeface";
        public string LatinFont
        {
            get
            {
                return GetXmlNodeString(_fontLatinPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_fontLatinPath, value);
            }
        }

        protected internal void CreateTopNode()
        {
            if (_path!="" && TopNode == _rootNode)
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
                return GetXmlNodeString(_fontCsPath);
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_fontCsPath, value);
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
                SetXmlNodeString(_boldPath, value ? "1" : "0");
            }
        }
        string _underLinePath = "@u";
        public eUnderLineType UnderLine
        {
            get
            {
                return TranslateUnderline(GetXmlNodeString(_underLinePath));
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_underLinePath, TranslateUnderlineText(value));
            }
        }
        string _underLineColorPath = "a:uFill/a:solidFill/a:srgbClr/@val";
        public Color UnderLineColor
        {
            get
            {
                string col = GetXmlNodeString(_underLineColorPath);
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
                SetXmlNodeString(_underLineColorPath, value.ToArgb().ToString("X").Substring(2, 6));
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
                SetXmlNodeString(_italicPath, value ? "1" : "0");
            }
        }
        string _strikePath = "@strike";
        public eStrikeType Strike
        {
            get
            {
                return TranslateStrike(GetXmlNodeString(_strikePath));
            }
            set
            {
                CreateTopNode();
                SetXmlNodeString(_strikePath, TranslateStrikeText(value));
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
                SetXmlNodeString(_sizePath, ((int)(value * 100)).ToString());
            }
        }
        string _colorPath = "a:solidFill/a:srgbClr/@val";
        public Color Color
        {
            get
            {
                string col = GetXmlNodeString(_colorPath);
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
                SetXmlNodeString(_colorPath, value.ToArgb().ToString("X").Substring(2, 6));
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
                case "":
                    return eUnderLineType.None;
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
        /// <summary>
        /// Set the font style from a font object
        /// </summary>
        /// <param name="Font"></param>
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
