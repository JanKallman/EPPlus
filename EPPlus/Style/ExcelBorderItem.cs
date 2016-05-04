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
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Cell border style
    /// </summary>
    public sealed class ExcelBorderItem : StyleBase
    {
        eStyleClass _cls;
        StyleBase _parent;
        internal ExcelBorderItem (ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent) : 
            base(styles, ChangedEvent, worksheetID, address)
	    {
            _cls=cls;
            _parent = parent;
	    }
        /// <summary>
        /// The line style of the border
        /// </summary>
        public ExcelBorderStyle Style
        {
            get
            {
                return GetSource().Style;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(_cls, eStyleProperty.Style, value, _positionID, _address));
            }
        }
        ExcelColor _color=null;
        /// <summary>
        /// The color of the border
        /// </summary>
        public ExcelColor Color
        {
            get
            {
                if (_color == null)
                {
                    _color = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, _cls, _parent);
                }
                return _color;
            }
        }

        internal override string Id
        {
            get { return Style + Color.Id; }
        }
        internal override void SetIndex(int index)
        {
            _parent.Index = index;
        }
        private ExcelBorderItemXml GetSource()
        {
            int ix = _parent.Index < 0 ? 0 : _parent.Index;

            switch(_cls)
            {
                case eStyleClass.BorderTop:
                    return _styles.Borders[ix].Top;
                case eStyleClass.BorderBottom:
                    return _styles.Borders[ix].Bottom;
                case eStyleClass.BorderLeft:
                    return _styles.Borders[ix].Left;
                case eStyleClass.BorderRight:
                    return _styles.Borders[ix].Right;
                case eStyleClass.BorderDiagonal:
                    return _styles.Borders[ix].Diagonal;
                default:
                    throw new Exception("Invalid class for Borderitem");
            }

        }
    }
}
