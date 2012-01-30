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
using OfficeOpenXml.Style;
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for border items
    /// </summary>
    public sealed class ExcelBorderItemXml : StyleXmlHelper
    {
        internal ExcelBorderItemXml(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            _borderStyle=ExcelBorderStyle.None;
            _color = new ExcelColorXml(NameSpaceManager);
        }
        internal ExcelBorderItemXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            if (topNode != null)
            {
                _borderStyle = GetBorderStyle(GetXmlNodeString("@style"));
                _color = new ExcelColorXml(nsm, topNode.SelectSingleNode(_colorPath, nsm));
                Exists = true;
            }
            else
            {
                Exists = false;
            }
        }

        private ExcelBorderStyle GetBorderStyle(string style)
        {
            if(style=="") return ExcelBorderStyle.None;
            string sInStyle = style.Substring(0, 1).ToUpper() + style.Substring(1, style.Length - 1);
            try
            {
                return (ExcelBorderStyle)Enum.Parse(typeof(ExcelBorderStyle), sInStyle);
            }
            catch
            {
                return ExcelBorderStyle.None;
            }

        }
        ExcelBorderStyle _borderStyle = ExcelBorderStyle.None;
        /// <summary>
        /// Cell Border style
        /// </summary>
        public ExcelBorderStyle Style
        {
            get
            {
                return _borderStyle;
            }
            set
            {
                _borderStyle = value;
                Exists = true;
            }
        }
        ExcelColorXml _color = null;
        const string _colorPath = "d:color";
        /// <summary>
        /// Border style
        /// </summary>
        public ExcelColorXml Color
        {
            get
            {
                return _color;
            }
            internal set
            {
                _color = value;
            }
        }
        internal override string Id
        {
            get { return Style + Color.Id; }
        }

        internal ExcelBorderItemXml Copy()
        {
            ExcelBorderItemXml borderItem = new ExcelBorderItemXml(NameSpaceManager);
            borderItem.Style = _borderStyle;
            borderItem.Color = _color.Copy();
            return borderItem;
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;

            if (Style != ExcelBorderStyle.None)
            {
                SetXmlNodeString("@style", SetBorderString(Style));
                if (Color.Exists)
                {
                    CreateNode(_colorPath);
                    topNode.AppendChild(Color.CreateXmlNode(TopNode.SelectSingleNode(_colorPath,NameSpaceManager)));
                }
            }
            return TopNode;
        }

        private string SetBorderString(ExcelBorderStyle Style)
        {
            string newName=Enum.GetName(typeof(ExcelBorderStyle), Style);
            return newName.Substring(0, 1).ToLower() + newName.Substring(1, newName.Length - 1);
        }
        public bool Exists { get; private set; }
    }
}
