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
using System.Globalization;
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for gradient fillsde
    /// </summary>
    public sealed class ExcelGradientFillXml : ExcelFillXml
    {
        internal ExcelGradientFillXml(XmlNamespaceManager nameSpaceManager)
            : base(nameSpaceManager)
        {
            GradientColor1 = new ExcelColorXml(nameSpaceManager);
            GradientColor2 = new ExcelColorXml(nameSpaceManager);
        }
        internal ExcelGradientFillXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            Degree = GetXmlNodeDouble(_degreePath);
            Type = GetXmlNodeString(_typePath)=="path" ? ExcelFillGradientType.Path : ExcelFillGradientType.Linear;
            GradientColor1 = new ExcelColorXml(nsm, topNode.SelectSingleNode(_gradientColor1Path, nsm));
            GradientColor2 = new ExcelColorXml(nsm, topNode.SelectSingleNode(_gradientColor2Path, nsm));
            
            Top = GetXmlNodeDouble(_topPath);
            Bottom = GetXmlNodeDouble(_bottomPath);
            Left = GetXmlNodeDouble(_leftPath);
            Right = GetXmlNodeDouble(_rightPath);
        }
        const string _typePath = "d:gradientFill/@type";
        /// <summary>
        /// Type of gradient fill. 
        /// </summary>
        public ExcelFillGradientType Type
        {
            get;
            internal set;
        }
        const string _degreePath = "d:gradientFill/@degree";
        /// <summary>
        /// Angle of the linear gradient
        /// </summary>
        public double Degree
        {
            get;
            internal set;
        }
        const string _gradientColor1Path = "d:gradientFill/d:stop[@position=\"0\"]/d:color";
        /// <summary>
        /// Gradient color 1
        /// </summary>
        public ExcelColorXml GradientColor1 
        {
            get;
            private set;
        }
        const string _gradientColor2Path = "d:gradientFill/d:stop[@position=\"1\"]/d:color";
        /// <summary>
        /// Gradient color 2
        /// </summary>
        public ExcelColorXml GradientColor2
        {
            get;
            private set;
        }
        const string _bottomPath = "d:gradientFill/@bottom";
        /// <summary>
        /// Percentage format bottom
        /// </summary>
        public double Bottom
        { 
            get; 
            internal set; 
        }
        const string _topPath = "d:gradientFill/@top";
        /// <summary>
        /// Percentage format top
        /// </summary>
        public double Top
        {
            get;
            internal set;
        }
        const string _leftPath = "d:gradientFill/@left";
        /// <summary>
        /// Percentage format left
        /// </summary>
        public double Left
        {
            get;
            internal set;
        }
        const string _rightPath = "d:gradientFill/@right";
        /// <summary>
        /// Percentage format right
        /// </summary>
        public double Right
        {
            get;
            internal set;
        }
        internal override string Id
        {
            get
            {
                return base.Id + Degree.ToString() + GradientColor1.Id + GradientColor2.Id + Type + Left.ToString() + Right.ToString() + Bottom.ToString() + Top.ToString();
            }
        }

        #region Public Properties
        #endregion
        internal override ExcelFillXml Copy()
        {
            ExcelGradientFillXml newFill = new ExcelGradientFillXml(NameSpaceManager);
            newFill.PatternType = base._fillPatternType;
            newFill.BackgroundColor = _backgroundColor.Copy();
            newFill.PatternColor = _patternColor.Copy();

            newFill.GradientColor1 = GradientColor1.Copy();
            newFill.GradientColor2 = GradientColor2.Copy();
            newFill.Type = Type;
            newFill.Degree = Degree;
            newFill.Top = Top;
            newFill.Bottom = Bottom;
            newFill.Left = Left;
            newFill.Right = Right;
            
            return newFill;
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            CreateNode("d:gradientFill");
            if(Type==ExcelFillGradientType.Path) SetXmlNodeString(_typePath, "path");
            if(!double.IsNaN(Degree)) SetXmlNodeString(_degreePath, Degree.ToString(CultureInfo.InvariantCulture));
            if (GradientColor1!=null)
            {
                /*** Gradient color node 1***/
                var node = TopNode.SelectSingleNode("d:gradientFill", NameSpaceManager);
                var stopNode = node.OwnerDocument.CreateElement("stop", ExcelPackage.schemaMain);
                stopNode.SetAttribute("position", "0");
                node.AppendChild(stopNode);
                var colorNode = node.OwnerDocument.CreateElement("color", ExcelPackage.schemaMain);
                stopNode.AppendChild(colorNode);
                GradientColor1.CreateXmlNode(colorNode);

                /*** Gradient color node 2***/
                stopNode = node.OwnerDocument.CreateElement("stop", ExcelPackage.schemaMain);
                stopNode.SetAttribute("position", "1");
                node.AppendChild(stopNode);
                colorNode = node.OwnerDocument.CreateElement("color", ExcelPackage.schemaMain);
                stopNode.AppendChild(colorNode);

                GradientColor2.CreateXmlNode(colorNode);
            }
            if (!double.IsNaN(Top)) SetXmlNodeString(_topPath, Top.ToString("F5",CultureInfo.InvariantCulture));
            if (!double.IsNaN(Bottom)) SetXmlNodeString(_bottomPath, Bottom.ToString("F5", CultureInfo.InvariantCulture));
            if (!double.IsNaN(Left)) SetXmlNodeString(_leftPath, Left.ToString("F5", CultureInfo.InvariantCulture));
            if (!double.IsNaN(Right)) SetXmlNodeString(_rightPath, Right.ToString("F5", CultureInfo.InvariantCulture));

            return topNode;
        }
    }
}
