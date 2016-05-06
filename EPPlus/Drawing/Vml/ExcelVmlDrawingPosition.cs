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
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// The position of a VML drawing. Used for comments
    /// </summary>
    public class ExcelVmlDrawingPosition : XmlHelper
    {
        int _startPos;
        internal ExcelVmlDrawingPosition(XmlNamespaceManager ns, XmlNode topNode, int startPos) : 
            base(ns, topNode)
        {
            _startPos = startPos;
        }
        /// <summary>
        /// Row. Zero based
        /// </summary>
        public int Row
        {
            get
            {
                return GetNumber(2);
            }
            set
            {
                SetNumber(2, value);
            } 
        }
        /// <summary>
        /// Row offset in pixels. Zero based
        /// </summary>
        public int RowOffset
        {
            get
            {
                return GetNumber(3);
            }
            set
            {
                SetNumber(3, value);
            }
        }
        /// <summary>
        /// Column. Zero based
        /// </summary>
        public int Column
        {
            get
            {
                return GetNumber(0);
            }
            set
            {
                SetNumber(0, value);
            }
        }
        /// <summary>
        /// Column offset. Zero based
        /// </summary>
        public int ColumnOffset
        {
            get
            {
                return GetNumber(1);
            }
            set
            {
                SetNumber(1, value);
            }
        }
        private void SetNumber(int pos, int value)
        {
            string anchor = GetXmlNodeString("x:Anchor");
            string[] numbers = anchor.Split(',');
            if (numbers.Length == 8)
            {
                numbers[_startPos + pos] = value.ToString();
            }
            else
            {
                throw (new Exception("Anchor element is invalid in vmlDrawing"));
            }
            SetXmlNodeString("x:Anchor", string.Join(",",numbers));
        }

        private int GetNumber(int pos)
        {
            string anchor = GetXmlNodeString("x:Anchor");
            string[] numbers = anchor.Split(',');
            if (numbers.Length == 8)
            {
                int ret;
                if (int.TryParse(numbers[_startPos + pos], out ret))
                {
                    return ret;
                }
            }
            throw(new Exception("Anchor element is invalid in vmlDrawing"));
        }
    }
}
