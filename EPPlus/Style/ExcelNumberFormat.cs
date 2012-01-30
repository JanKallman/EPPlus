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
using System.Text.RegularExpressions;
using System.Globalization;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// The numberformat of the cell
    /// </summary>
    public sealed class ExcelNumberFormat : StyleBase
    {
        internal ExcelNumberFormat(ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int PositionID, string Address, int index) :
            base(styles, ChangedEvent, PositionID, Address)
        {
            Index = index;
        }
        /// <summary>
        /// The numeric index fror the format
        /// </summary>
        public int NumFmtID 
        {
            get
            {
                return Index;
            }
            //set
            //{
            //    _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Numberformat, "NumFmtID", value, _workSheetID, _address));
            //}
        }
        /// <summary>
        /// The numberformat 
        /// </summary>
        public string Format
        {
            get
            {
                for(int i=0;i<_styles.NumberFormats.Count;i++)
                {
                    if(Index==_styles.NumberFormats[i].NumFmtId)
                    {
                        return _styles.NumberFormats[i].Format;
                    }
                }
                return "general";
            }
            set
            {
                _ChangedEvent(this, new StyleChangeEventArgs(eStyleClass.Numberformat, eStyleProperty.Format, value, _positionID, _address));
            }
        }

        internal override string Id
        {
            get 
            {
                return Format;
            }
        }
        /// <summary>
        /// If the numeric format is a build-in from.
        /// </summary>
        public bool BuildIn { get; private set; }
    }
}
