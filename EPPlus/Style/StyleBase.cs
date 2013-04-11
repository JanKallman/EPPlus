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

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Border line style
    /// </summary>
    public enum ExcelBorderStyle
    {
        None,
        Hair,
        Dotted,
        DashDot,
        Thin,
        DashDotDot,
        Dashed,
        MediumDashDotDot,
        MediumDashed,
        MediumDashDot,
        Thick,
        Medium,
        Double
    };
    /// <summary>
    /// Horizontal text alignment
    /// </summary>
    public enum ExcelHorizontalAlignment
    {
        General,
        Left,
        Center,
        CenterContinuous,
        Right,
        Fill,
        Distributed,
        Justify
    }
    /// <summary>
    /// Vertical text alignment
    /// </summary>
    public enum ExcelVerticalAlignment
    {
        Top,
        Center,
        Bottom,
        Distributed,
        Justify
    }
    /// <summary>
    /// Font-Vertical Align
    /// </summary>
    public enum ExcelVerticalAlignmentFont
    {
        None,
        Subscript,
        Superscript
    }
    /// <summary>
    /// Font-Underlinestyle for 
    /// </summary>
    public enum ExcelUnderLineType
    {
        None,
        Single,
        Double,
        SingleAccounting,
        DoubleAccounting
    }
    /// <summary>
    /// Fill pattern
    /// </summary>
    public enum ExcelFillStyle
    {
        None,
        Solid,
        DarkGray,
        MediumGray,
        LightGray,
        Gray125,
        Gray0625,
        DarkVertical,
        DarkHorizontal,
        DarkDown,
        DarkUp,
        DarkGrid,
        DarkTrellis,
        LightVertical,
        LightHorizontal,
        LightDown,
        LightUp,
        LightGrid,
        LightTrellis
    }
    /// <summary>
    /// Type of gradient fill
    /// </summary>
    public enum ExcelFillGradientType
    {
        /// <summary>
        /// No gradient fill. 
        /// </summary>
        None,
        /// <summary>
        /// This gradient fill is of linear gradient type. Linear gradient type means that the transition from one color to the next is along a line (e.g., horizontal, vertical,diagonal, etc.)
        /// </summary>
        Linear,
        /// <summary>
        /// This gradient fill is of path gradient type. Path gradient type means the that the boundary of transition from one color to the next is a rectangle, defined by top,bottom, left, and right attributes on the gradientFill element.
        /// </summary>
        Path
    }
    /// <summary>
    /// The reading order
    /// </summary>
    public enum ExcelReadingOrder
    {
        /// <summary>
        /// Reading order is determined by scanning the text for the first non-whitespace character: if it is a strong right-to-left character, the reading order is right-to-left; otherwise, the reading order left-to-right.
        /// </summary>
        ContextDependent=0,
        /// <summary>
        /// Left to Right
        /// </summary>
        LeftToRight=1,
        /// <summary>
        /// Right to Left
        /// </summary>
        RightToLeft=2
    }
    public abstract class StyleBase
    {
        protected ExcelStyles _styles;
        internal OfficeOpenXml.XmlHelper.ChangedEventHandler _ChangedEvent;
        protected int _positionID;
        protected string _address;
        internal StyleBase(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address)
        {
            _styles = styles;
            _ChangedEvent = ChangedEvent;
            _address = Address;
            _positionID = PositionID;
        }
        internal int Index { get; set;}
        internal abstract string Id {get;}

        internal virtual void SetIndex(int index)
        {
            Index = index;
        }
    }
}
