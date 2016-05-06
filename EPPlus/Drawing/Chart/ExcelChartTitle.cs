/************** *****************************************************************
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
 *******************************************************************************
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// The title of a chart
    /// </summary>
    public class ExcelChartTitle : XmlHelper
    {
        internal ExcelChartTitle(XmlNamespaceManager nameSpaceManager, XmlNode node) :
            base(nameSpaceManager, node)
        {
            XmlNode topNode = node.SelectSingleNode("c:title", NameSpaceManager);
            if (topNode == null)
            {
                topNode = node.OwnerDocument.CreateElement("c", "title", ExcelPackage.schemaChart);
                node.InsertBefore(topNode, node.ChildNodes[0]);
                topNode.InnerXml = "<c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:pPr><a:defRPr sz=\"1800\" b=\"0\" /></a:pPr><a:r><a:t /></a:r></a:p></c:rich></c:tx><c:layout /><c:overlay val=\"0\" />";
            }
            TopNode = topNode;
            SchemaNodeOrder = new string[] { "tx","bodyPr", "lstStyle", "layout", "overlay" };
        }
        const string titlePath = "c:tx/c:rich/a:p/a:r/a:t";
        /// <summary>
        /// The text
        /// </summary>
        public string Text
        {
            get
            {
                //return GetXmlNode(titlePath);
                return RichText.Text;
            }
            set
            {
                //SetXmlNode(titlePath, value);
                RichText.Text = value;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// A reference to the border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "c:spPr/a:ln");
                }
                return _border;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "c:spPr");
                }
                return _fill;
            }
        }
        //ExcelTextFont _font = null;
        /// <summary>
        /// A reference to the font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                //if (_font == null)
                //{
                //    _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:tx/c:rich/a:p/a:r/a:rPr", new string[] { "rPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" });
                //}
                //return _font;
                if (_richText==null || _richText.Count == 0)
                {
                    RichText.Add("");
                }
                return _richText[0];
            }
        }
        string[] paragraphNodeOrder = new string[] { "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" };
        ExcelParagraphCollection _richText = null;
        /// <summary>
        /// Richtext
        /// </summary>
        public ExcelParagraphCollection RichText
        {
            get
            {
                if (_richText == null)
                {
                    _richText = new ExcelParagraphCollection(NameSpaceManager, TopNode, "c:tx/c:rich/a:p", paragraphNodeOrder);
                }
                return _richText;
            }
        }
        /// <summary>
        /// Show without overlaping the chart.
        /// </summary>
        public bool Overlay
        {
            get
            {
                return GetXmlNodeBool("c:overlay/@val");
            }
            set
            {
                SetXmlNodeBool("c:overlay/@val", value);
            }
        }
        /// <summary>
        /// Specifies the centering of the text box. 
        /// The way it works fundamentally is to determine the smallest possible "bounds box" for the text and then to center that "bounds box" accordingly. 
        /// This is different than paragraph alignment, which aligns the text within the "bounds box" for the text. 
        /// This flag is compatible with all of the different kinds of anchoring. 
        /// If this attribute is omitted, then a value of 0 or false is implied.
        /// </summary>
        public bool AnchorCtr
        {
            get
            {
                return GetXmlNodeBool("c:tx/c:rich/a:bodyPr/@anchorCtr", false);
            }
            set
            {
                SetXmlNodeBool("c:tx/c:rich/a:bodyPr/@anchorCtr", value, false);
            }
        }
        public eTextAnchoringType Anchor
        {
            get
            {
                return ExcelDrawing.GetTextAchoringEnum(GetXmlNodeString("c:tx/c:rich/a:bodyPr/@anchor"));
            }
            set
            {
                SetXmlNodeString("c:tx/c:rich/a:bodyPr/@anchorCtr", ExcelDrawing.GetTextAchoringText(value));
            }
        }
        const string TextVerticalPath = "xdr:sp/xdr:txBody/a:bodyPr/@vert";
        /// <summary>
        /// Vertical text
        /// </summary>
        public eTextVerticalType TextVertical
        {
            get
            {
                return ExcelDrawing.GetTextVerticalEnum(GetXmlNodeString("c:tx/c:rich/a:bodyPr/@vert"));
            }
            set
            {
                SetXmlNodeString("c:tx/c:rich/a:bodyPr/@vert", ExcelDrawing.GetTextVerticalText(value));
            }
        }
        /// <summary>
        /// Rotation in degrees (0-360)
        /// </summary>
        public double Rotation
        {
            get
            {
                var i=GetXmlNodeInt("c:tx/c:rich/a:bodyPr/@rot");
                if (i < 0)
                {
                    return 360 - (i / 60000);
                }
                else
                {
                    return (i / 60000);
                }
            }
            set
            {
                int v;
                if(value <0 || value > 360)
                {
                    throw(new ArgumentOutOfRangeException("Rotation must be between 0 and 360"));
                }

                if (value > 180)
                {
                    v = (int)((value - 360) * 60000);
                }
                else
                {
                    v = (int)(value * 60000);
                }
                SetXmlNodeString("c:tx/c:rich/a:bodyPr/@rot", v.ToString());
            }
        }
    }
}
