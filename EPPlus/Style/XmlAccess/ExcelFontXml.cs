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
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for fonts
    /// </summary>
    public sealed class ExcelFontXml : StyleXmlHelper
    {
        internal ExcelFontXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {
            _name = "";
            _size = 0;
            _family = int.MinValue;
            _scheme = "";
            _color = _color = new ExcelColorXml(NameSpaceManager);
            _bold = false;
            _italic = false;
            _strike = false;
            _underlineType = ExcelUnderLineType.None ;
            _verticalAlign = "";
        }
        internal ExcelFontXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _name = GetXmlNodeString(namePath);
            _size = (float)GetXmlNodeDecimal(sizePath);
            _family = GetXmlNodeInt(familyPath);
            _scheme = GetXmlNodeString(schemePath);
            _color = new ExcelColorXml(nsm, topNode.SelectSingleNode(_colorPath, nsm));
            _bold = (topNode.SelectSingleNode(boldPath, NameSpaceManager) != null);
            _italic = (topNode.SelectSingleNode(italicPath, NameSpaceManager) != null);
            _strike = (topNode.SelectSingleNode(strikePath, NameSpaceManager) != null);
            _verticalAlign = GetXmlNodeString(verticalAlignPath);
            if (topNode.SelectSingleNode(underLinedPath, NameSpaceManager) != null)
            {
                string ut = GetXmlNodeString(underLinedPath + "/@val");
                if (ut == "")
                {
                    _underlineType = ExcelUnderLineType.Single;
                }
                else
                {
                    _underlineType = (ExcelUnderLineType)Enum.Parse(typeof(ExcelUnderLineType), ut, true);
                }
            }
            else
            {
                _underlineType = ExcelUnderLineType.None;
            }
        }
        internal override string Id
        {
            get
            {
                return Name + "|" + Size + "|" + Family + "|" + Color.Id + "|" + Scheme + "|" + Bold.ToString() + "|" + Italic.ToString() + "|" + Strike.ToString() + "|" + VerticalAlign + "|" + UnderLineType.ToString();
            }
        }
        const string namePath = "d:name/@val";
        string _name;
        /// <summary>
        /// The name of the font
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            set
            {
                Scheme = "";        //Reset schema to avoid corrupt file if unsupported font is selected.
                _name = value;
            }
        }
        const string sizePath = "d:sz/@val";
        float _size;
        /// <summary>
        /// Font size
        /// </summary>
        public float Size
        {
            get
            {
                return _size;
            }
            set
            {
                _size = value;
            }
        }
        const string familyPath = "d:family/@val";
        int _family;
        /// <summary>
        /// Font family
        /// </summary>
        public int Family
        {
            get
            {
                return _family;
            }
            set
            {
                _family=value;
            }
        }
        ExcelColorXml _color = null;
        const string _colorPath = "d:color";
        /// <summary>
        /// Text color
        /// </summary>
        public ExcelColorXml Color
        {
            get
            {
                return _color;
            }
            internal set 
            {
                _color=value;
            }
        }
        const string schemePath = "d:scheme/@val";
        string _scheme="";
        /// <summary>
        /// Font Scheme
        /// </summary>
        public string Scheme
        {
            get
            {
                return _scheme;
            }
            private set
            {
                _scheme=value;
            }
        }
        const string boldPath = "d:b";
        bool _bold;
        /// <summary>
        /// If the font is bold
        /// </summary>
        public bool Bold
        {
            get
            {
                return _bold;
            }
            set
            {
                _bold=value;
            }
        }
        const string italicPath = "d:i";
        bool _italic;
        /// <summary>
        /// If the font is italic
        /// </summary>
        public bool Italic
        {
            get
            {
                return _italic;
            }
            set
            {
                _italic=value;
            }
        }
        const string strikePath = "d:strike";
        bool _strike;
        /// <summary>
        /// If the font is striked out
        /// </summary>
        public bool Strike
        {
            get
            {
                return _strike;
            }
            set
            {
                _strike=value;
            }
        }
        const string underLinedPath = "d:u";
        /// <summary>
        /// If the font is underlined.
        /// When set to true a the text is underlined with a single line
        /// </summary>
        public bool UnderLine
        {
            get
            {
                return UnderLineType!=ExcelUnderLineType.None;
            }
            set
            {
                _underlineType=value ? ExcelUnderLineType.Single : ExcelUnderLineType.None;
            }
        }
        ExcelUnderLineType _underlineType;
        /// <summary>
        /// If the font is underlined
        /// </summary>
        public ExcelUnderLineType UnderLineType
        {
            get
            {
                return _underlineType;
            }
            set
            {
                _underlineType = value;
            }
        }
        const string verticalAlignPath = "d:vertAlign/@val";
        string _verticalAlign;
        /// <summary>
        /// Vertical aligned
        /// </summary>
        public string VerticalAlign
        {
            get
            {
                return _verticalAlign;
            }
            set
            {
                _verticalAlign=value;
            }
        }
        public void SetFromFont(System.Drawing.Font Font)
        {
            Name=Font.Name;
            //Family=fnt.FontFamily.;
            Size=(int)Font.Size;
            Strike=Font.Strikeout;
            Bold = Font.Bold;
            UnderLine=Font.Underline;
            Italic=Font.Italic;
        }
        internal ExcelFontXml Copy()
        {
            ExcelFontXml newFont = new ExcelFontXml(NameSpaceManager);
            newFont.Name = Name;
            newFont.Size = Size;
            newFont.Family = Family;
            newFont.Scheme = Scheme;
            newFont.Bold = Bold;
            newFont.Italic = Italic;
            newFont.UnderLineType = UnderLineType;
            newFont.Strike = Strike;
            newFont.VerticalAlign = VerticalAlign;
            newFont.Color = Color.Copy();
            return newFont;
        }

        internal override XmlNode CreateXmlNode(XmlNode topElement)
        {
            TopNode = topElement;

            if (_bold) CreateNode(boldPath); else DeleteAllNode(boldPath);
            if (_italic) CreateNode(italicPath); else DeleteAllNode(italicPath);
            if (_strike) CreateNode(strikePath); else DeleteAllNode(strikePath);
            
            if (_underlineType == ExcelUnderLineType.None)
            {
                DeleteAllNode(underLinedPath);
            }
            else if(_underlineType==ExcelUnderLineType.Single)
            {
                CreateNode(underLinedPath);
            }
            else
            {
                var v=_underlineType.ToString();
                SetXmlNodeString(underLinedPath + "/@val", v.Substring(0, 1).ToLower() + v.Substring(1));
            }

            if (_verticalAlign!="") SetXmlNodeString(verticalAlignPath, _verticalAlign.ToString());
            SetXmlNodeString(sizePath, _size.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (_color.Exists)
            {
                CreateNode(_colorPath);
                TopNode.AppendChild(_color.CreateXmlNode(TopNode.SelectSingleNode(_colorPath, NameSpaceManager)));
            }
            SetXmlNodeString(namePath, _name);
            if(_family>int.MinValue) SetXmlNodeString(familyPath, _family.ToString());
            if (_scheme != "") SetXmlNodeString(schemePath, _scheme.ToString());

            return TopNode;
        }
    }
}
