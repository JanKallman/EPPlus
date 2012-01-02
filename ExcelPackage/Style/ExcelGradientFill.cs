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
using System.Globalization;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// The background fill of a cell
    /// </summary>
    public class ExcelGradientFill : StyleBase
    {
        internal ExcelGradientFill(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)

        {
            Index = index;
        }
        /// <summary>
        /// Angle of the linear gradient
        /// </summary>
        public double Degree
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Degree;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientDegree, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Linear or Path gradient
        /// </summary>
        public ExcelFillGradientType Type
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Type;
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientType, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Specifies in percentage format(from the top to the bottom) the position of the top edge of the inner rectangle (color 1). For top, 0 means the top edge of the inner rectangle is on the top edge of the cell, and 1 means it is on the bottom edge of the cell. (applies to From Corner and From Center gradients).
        /// </summary>
        public double Top
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Top;
            }
            set
            {
                if (value < 0 | value > 1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientTop, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Specifies in percentage format (from the top to the bottom) the position of the bottom edge of the inner rectangle (color 1). For bottom, 0 means the bottom edge of the inner rectangle is on the top edge of the cell, and 1 means it is on the bottom edge of the cell.
        /// </summary>
        public double Bottom
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Bottom;
            }
            set
            {
                if (value < 0 | value > 1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientBottom, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Specifies in percentage format (from the left to the right) the position of the left edge of the inner rectangle (color 1). For left, 0 means the left edge of the inner rectangle is on the left edge of the cell, and 1 means it is on the right edge of the cell. (applies to From Corner and From Center gradients).
        /// </summary>
        public double Left
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Left;
            }
            set
            {
                if (value < 0 | value > 1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientLeft, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Specifies in percentage format (from the left to the right) the position of the right edge of the inner rectangle (color 1). For right, 0 means the right edge of the inner rectangle is on the left edge of the cell, and 1 means it is on the right edge of the cell. (applies to From Corner and From Center gradients).
        /// </summary>
        public double Right
        {
            get
            {
                return ((ExcelGradientFillXml)_styles.Fills[Index]).Right;
            }
            set
            {
                if (value < 0 | value > 1)
                {
                    throw (new ArgumentOutOfRangeException("Value must be between 0 and 1"));
                }
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.GradientFill, eStyleProperty.GradientRight, value, _positionID, _address));
            }
        }
        ExcelColor _gradientColor1 = null;
        /// <summary>
        /// Gradient Color 1
        /// </summary>
        public ExcelColor Color1
        {
            get
            {
                if (_gradientColor1 == null)
                {
                    _gradientColor1 = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillGradientColor1, this);
                }
                return _gradientColor1;

            }
        }
        ExcelColor _gradientColor2 = null;
        /// <summary>
        /// Gradient Color 2
        /// </summary>
        public ExcelColor Color2
        {
            get
            {
                if (_gradientColor2 == null)
                {
                    _gradientColor2 = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillGradientColor2, this);
                }
                return _gradientColor2;

            }
        }
        internal override string Id
        {
            get { return Degree.ToString() + Type + Color1.Id + Color2.Id + Top.ToString() + Bottom.ToString() + Left.ToString() + Right.ToString(); }
        }
    }
}
