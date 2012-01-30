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
 * Jan Källman		Added this class		        2010-01-24
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// HyperlinkClass
    /// </summary>
    public class ExcelHyperLink : Uri
    {
        /// <summary>
        /// A new hyperlink with the specified URI
        /// </summary>
        /// <param name="uriString">The URI</param>
        public ExcelHyperLink(string uriString) :
            base(uriString)
        {

        }
        /// <summary>
        /// A new hyperlink with the specified URI. This syntax is obsolete
        /// </summary>
        /// <param name="uriString">The URI</param>
        /// <param name="dontEscape"></param>
        public ExcelHyperLink(string uriString, bool dontEscape) :
            base(uriString, dontEscape)
        {

        }
        /// <summary>
        /// A new hyperlink with the specified URI and kind
        /// </summary>
        /// <param name="uriString">The URI</param>
        /// <param name="uriKind">Kind (absolute/relative or indeterminate)</param>
        public ExcelHyperLink(string uriString, UriKind uriKind) :
            base(uriString, uriKind)
        {

        }
        /// <summary>
        /// Sheet internal reference
        /// </summary>
        /// <param name="referenceAddress">Address</param>
        /// <param name="display">Displayed text</param>
        public ExcelHyperLink(string referenceAddress, string display) :
            base("xl://internal")   //URI is not used on internal links so put a dummy uri here.
        {
            _referenceAddress = referenceAddress;
            _display = display;
        }

        string _referenceAddress = null;
        /// <summary>
        /// The Excel address for internal links.
        /// </summary>
        public string ReferenceAddress
        {
            get
            {
                return _referenceAddress;
            }
            set
            {
                _referenceAddress = value;
            }
        }
        string _display = "";
        /// <summary>
        /// Displayed text
        /// </summary>
        public string Display
        {
            get
            {
                return _display;
            }
            set
            {
                _display = value;
            }
        }
        /// <summary>
        /// Tooltip
        /// </summary>
        public string ToolTip
        {
            get;
            set;
        }
        int _colSpann = 0;
        /// <summary>
        /// If the hyperlink spans multiple columns
        /// </summary>
        public int ColSpann
        {
            get
            {
                return _colSpann;
            }
            set
            {
                _colSpann = value;
            }
        }
        int _rowSpann = 0;
        /// <summary>
        /// If the hyperlink spans multiple rows
        /// </summary>
        public int RowSpann
        {
            get
            {
                return _rowSpann;
            }
            set
            {
                _rowSpann = value;
            }
        }
    }
}
