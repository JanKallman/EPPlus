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
    /// Xml access class for named styles
    /// </summary>
    public sealed class ExcelNamedStyleXml : StyleXmlHelper
    {
        ExcelStyles _styles;
        internal ExcelNamedStyleXml(XmlNamespaceManager nameSpaceManager, ExcelStyles styles)
            : base(nameSpaceManager)
        {
            _styles = styles;
            BuildInId = int.MinValue;
        }
        internal ExcelNamedStyleXml(XmlNamespaceManager NameSpaceManager, XmlNode topNode, ExcelStyles styles) :
            base(NameSpaceManager, topNode)
        {
            StyleXfId = GetXmlNodeInt(idPath);
            Name = GetXmlNodeString(namePath);
            BuildInId = GetXmlNodeInt(buildInIdPath);
            CustomBuildin = GetXmlNodeBool(customBuiltinPath);

            _styles = styles;
            _style = new ExcelStyle(styles, styles.NamedStylePropertyChange, -1, Name, _styleXfId);
        }
        internal override string Id
        {
            get
            {
                return Name;
            }
        }
        int _styleXfId=0;
        const string idPath = "@xfId";
        /// <summary>
        /// Named style index
        /// </summary>
        public int StyleXfId
        {
            get
            {
                return _styleXfId;
            }
            set
            {
                _styleXfId = value;
            }
        }
        int _xfId = int.MinValue;
        /// <summary>
        /// Style index
        /// </summary>
        internal int XfId
        {
            get
            {
                return _xfId;
            }
            set
            {
                _xfId = value;
            }
        }
        const string buildInIdPath = "@builtinId";
        public int BuildInId { get; set; }
        const string customBuiltinPath = "@customBuiltin";
        public bool CustomBuildin { get; set; }
        const string namePath = "@name";
        string _name;
        /// <summary>
        /// Name of the style
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
            internal set
            {
                _name = value;
            }
        }
        ExcelStyle _style = null;
        /// <summary>
        /// The style object
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return _style;
            }
            internal set
            {
                _style = value;
            }
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNodeString(namePath, _name);
            SetXmlNodeString("@xfId", _styles.CellStyleXfs[StyleXfId].newID.ToString());
            if (BuildInId>=0) SetXmlNodeString("@builtinId", BuildInId.ToString());
            if(CustomBuildin) SetXmlNodeBool(customBuiltinPath, true);
            return TopNode;            
        }
    }
}
