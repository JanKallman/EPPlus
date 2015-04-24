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
    /// Xml access class for border top level
    /// </summary>
    public sealed class ExcelBorderXml : StyleXmlHelper
    {
        internal ExcelBorderXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {

        }
        internal ExcelBorderXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _left = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(leftPath, nsm));
            _right = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(rightPath, nsm));
            _top = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(topPath, nsm));
            _bottom = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(bottomPath, nsm));
            _diagonal = new ExcelBorderItemXml(nsm, topNode.SelectSingleNode(diagonalPath, nsm));
            _diagonalUp = GetBoolValue(topNode, diagonalUpPath);
            _diagonalDown = GetBoolValue(topNode, diagonalDownPath);
        }
        internal override string Id
        {
            get
            {
                return Left.Id + Right.Id + Top.Id + Bottom.Id + Diagonal.Id + DiagonalUp.ToString() + DiagonalDown.ToString();
            }
        }
        const string leftPath = "d:left";
        ExcelBorderItemXml _left = null;
        /// <summary>
        /// Left border style properties
        /// </summary>
        public ExcelBorderItemXml Left
        {
            get
            {
                return _left;
            }
            internal set
            {
                _left = value;
            }
        }
        const string rightPath = "d:right";
        ExcelBorderItemXml _right = null;
        /// <summary>
        /// Right border style properties
        /// </summary>
        public ExcelBorderItemXml Right
        {
            get
            {
                return _right;
            }
            internal set
            {
                _right = value;
            }
        }
        const string topPath = "d:top";
        ExcelBorderItemXml _top = null;
        /// <summary>
        /// Top border style properties
        /// </summary>
        public ExcelBorderItemXml Top
        {
            get
            {
                return _top;
            }
            internal set
            {
                _top = value;
            }
        }
        const string bottomPath = "d:bottom";
        ExcelBorderItemXml _bottom = null;
        /// <summary>
        /// Bottom border style properties
        /// </summary>
        public ExcelBorderItemXml Bottom
        {
            get
            {
                return _bottom;
            }
            internal set
            {
                _bottom = value;
            }
        }
        const string diagonalPath = "d:diagonal";
        ExcelBorderItemXml _diagonal = null;
        /// <summary>
        /// Diagonal border style properties
        /// </summary>
        public ExcelBorderItemXml Diagonal
        {
            get
            {
                return _diagonal;
            }
            internal set
            {
                _diagonal = value;
            }
        }
        const string diagonalUpPath = "@diagonalUp";
        bool _diagonalUp = false;
        /// <summary>
        /// Diagonal up border
        /// </summary>
        public bool DiagonalUp
        {
            get
            {
                return _diagonalUp;
            }
            internal set
            {
                _diagonalUp = value;
            }
        }
        const string diagonalDownPath = "@diagonalDown";
        bool _diagonalDown = false;
        /// <summary>
        /// Diagonal down border
        /// </summary>
        public bool DiagonalDown
        {
            get
            {
                return _diagonalDown;
            }
            internal set
            {
                _diagonalDown = value;
            }
        }

        internal ExcelBorderXml Copy()
        {
            ExcelBorderXml newBorder = new ExcelBorderXml(NameSpaceManager);
            newBorder.Bottom = _bottom.Copy();
            newBorder.Diagonal = _diagonal.Copy();
            newBorder.Left = _left.Copy();
            newBorder.Right = _right.Copy();
            newBorder.Top = _top.Copy();
            newBorder.DiagonalUp = _diagonalUp;
            newBorder.DiagonalDown = _diagonalDown;

            return newBorder;

        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            CreateNode(leftPath);
            topNode.AppendChild(_left.CreateXmlNode(TopNode.SelectSingleNode(leftPath, NameSpaceManager)));
            CreateNode(rightPath);
            topNode.AppendChild(_right.CreateXmlNode(TopNode.SelectSingleNode(rightPath, NameSpaceManager)));
            CreateNode(topPath);
            topNode.AppendChild(_top.CreateXmlNode(TopNode.SelectSingleNode(topPath, NameSpaceManager)));
            CreateNode(bottomPath);
            topNode.AppendChild(_bottom.CreateXmlNode(TopNode.SelectSingleNode(bottomPath, NameSpaceManager)));
            CreateNode(diagonalPath);
            topNode.AppendChild(_diagonal.CreateXmlNode(TopNode.SelectSingleNode(diagonalPath, NameSpaceManager)));
            if (_diagonalUp)
            {
                SetXmlNodeString(diagonalUpPath, "1");
            }
            if (_diagonalDown)
            {
                SetXmlNodeString(diagonalDownPath, "1");
            }
            return topNode;
        }
    }
}
