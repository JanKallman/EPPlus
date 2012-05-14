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
    public class ExcelFill : StyleBase
    {
        internal ExcelFill(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string address, int index) :
            base(styles, ChangedEvent, PositionID, address)

        {
            Index = index;
        }
        /// <summary>
        /// The pattern for solid fills.
        /// </summary>
        public ExcelFillStyle PatternType
        {
            get
            {
                if (Index == int.MinValue)
                {
                    return ExcelFillStyle.None;
                }
                else
                {
                    return _styles.Fills[Index].PatternType;
                }
            }
            set
            {
                if (_gradient != null) _gradient = null;
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Fill, eStyleProperty.PatternType, value, _positionID, _address));
            }
        }
        ExcelColor _patternColor = null;
        /// <summary>
        /// The color of the pattern
        /// </summary>
        public ExcelColor PatternColor
        {
            get
            {
                if (_patternColor == null)
                {
                    _patternColor = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillPatternColor, this);
                    if (_gradient != null) _gradient = null;
                }
                return _patternColor;
            }
        }
        ExcelColor _backgroundColor = null;
        /// <summary>
        /// The background color
        /// </summary>
        public ExcelColor BackgroundColor
        {
            get
            {
                if (_backgroundColor == null)
                {
                    _backgroundColor = new ExcelColor(_styles, _ChangedEvent, _positionID, _address, eStyleClass.FillBackgroundColor, this);
                    if (_gradient != null) _gradient = null;
                }
                return _backgroundColor;
                
            }
        }
        ExcelGradientFill _gradient=null;
        /// <summary>
        /// Access to properties for gradient fill.
        /// </summary>
        public ExcelGradientFill Gradient 
        {
            get
            {
                if (_gradient == null)
                {                    
                    _gradient = new ExcelGradientFill(_styles, _ChangedEvent, _positionID, _address, Index);
                    _backgroundColor = null;
                    _patternColor = null;
                }
                return _gradient;
            }
        }
        internal override string Id
        {
            get
            {
                if (_gradient == null)
                {
                    return PatternType + PatternColor.Id + BackgroundColor.Id;
                }
                else
                {
                    return _gradient.Id;
                }
            }
        }
    }
}
