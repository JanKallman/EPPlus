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
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// A richtext part
    /// </summary>
    public class ExcelRichText : XmlHelper
    {
        internal ExcelRichText(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
            SchemaNodeOrder=new string[] {"rPr", "t", "b", "i","strike", "u", "vertAlign" , "sz", "color", "rFont", "family", "scheme", "charset"};
            PreserveSpace = false;
        }
        internal delegate void CallbackDelegate();
        CallbackDelegate _callback;
        internal void SetCallback(CallbackDelegate callback)
        {
            _callback=callback;
        }
        const string TEXT_PATH="d:t";
        /// <summary>
        /// The text
        /// </summary>
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
        /// <summary>
        /// Preserves whitespace. Default true
        /// </summary>
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
        /// <summary>
        /// Bold text
        /// </summary>
        public bool Bold
        {
            get
            {
                return ExistNode(BOLD_PATH);
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
        /// <summary>
        /// Italic text
        /// </summary>
        public bool Italic
        {
            get
            {
                //return GetXmlNodeBool(ITALIC_PATH, false);
                return ExistNode(ITALIC_PATH);
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
        const string STRIKE_PATH = "d:rPr/d:strike";
        /// <summary>
        /// Strike-out text
        /// </summary>
        public bool Strike
        {
            get
            {
                return ExistNode(STRIKE_PATH);
            }
            set
            {
                if (value)
                {
                    CreateNode(STRIKE_PATH);
                }
                else
                {
                    DeleteNode(STRIKE_PATH);
                }
                if (_callback != null) _callback();
            }
        }
        const string UNDERLINE_PATH = "d:rPr/d:u";
        /// <summary>
        /// Underlined text
        /// </summary>
        public bool UnderLine
        {
            get
            {
                return ExistNode(UNDERLINE_PATH);
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

        const string VERT_ALIGN_PATH = "d:rPr/d:vertAlign/@val";
        /// <summary>
        /// Vertical Alignment
        /// </summary>
        public ExcelVerticalAlignmentFont VerticalAlign
        {
            get
            {
                string v=GetXmlNodeString(VERT_ALIGN_PATH);
                if(v=="")
                {
                    return ExcelVerticalAlignmentFont.None;
                }
                else
                {
                    try
                    {
                        return (ExcelVerticalAlignmentFont)Enum.Parse(typeof(ExcelVerticalAlignmentFont), v, true);
                    }
                    catch
                    {
                        return ExcelVerticalAlignmentFont.None;
                    }
                }
            }
            set
            {
                SetXmlNodeString(VERT_ALIGN_PATH,((ExcelVerticalAlignmentFont)value) == ExcelVerticalAlignmentFont.None ? "" : value.ToString().ToLower());
            }
        }
        const string SIZE_PATH = "d:rPr/d:sz/@val";
        /// <summary>
        /// Font size
        /// </summary>
        public float Size
        {
            get
            {
                return Convert.ToSingle(GetXmlNodeDecimal(SIZE_PATH));
            }
            set
            {
                SetXmlNodeString(SIZE_PATH, value.ToString(CultureInfo.InvariantCulture));
                if (_callback != null) _callback();
            }
        }
        const string FONT_PATH = "d:rPr/d:rFont/@val";
        /// <summary>
        /// Name of the font
        /// </summary>
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
        /// <summary>
        /// Text color
        /// </summary>
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
