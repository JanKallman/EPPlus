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
 * Jan Källman		Added		30-AUG-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// Build-in table row functions
    /// </summary>
    public enum RowFunctions
    {
        Average,        
        Count,
        CountNums,
        Custom,
        Max,
        Min,
        None,
        StdDev,
        Sum,
        Var
    }

    /// <summary>
    /// A table column
    /// </summary>
    public class ExcelTableColumn : XmlHelper
    {
        ExcelTable _tbl;
        internal ExcelTableColumn(XmlNamespaceManager ns, XmlNode topNode, ExcelTable tbl, int pos) :
            base(ns, topNode)
        {
            _tbl = tbl;
            Position = pos;
        }
        /// <summary>
        /// The column id
        /// </summary>
        public int Id 
        {
            get
            {
                return GetXmlNodeInt("@id");
            }
            set
            {
                SetXmlNodeString("@id", value.ToString());
            }
        }
        /// <summary>
        /// The position of the column
        /// </summary>
        public int Position
        {
            get;
            private set;
        }
        /// <summary>
        /// The name of the column
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// A string text in the total row
        /// </summary>
        public string TotalsRowLabel
        {
            get
            {
                return GetXmlNodeString("@totalsRowLabel");
            }
            set
            {
                SetXmlNodeString("@totalsRowLabel", value);
            }
        }
        /// <summary>
        /// Build-in total row functions.
        /// To set a custom Total row formula use the TotalsRowFormula property
        /// <seealso cref="TotalsRowFormula"/>
        /// </summary>
        public RowFunctions TotalsRowFunction
        {
            get
            {
                if (GetXmlNodeString("@totalsRowFunction") == "")
                {
                    return RowFunctions.None;
                }
                else
                {
                    return (RowFunctions)Enum.Parse(typeof(RowFunctions), GetXmlNodeString("@totalsRowFunction"), true);
                }
            }
            set
            {
                if (value == RowFunctions.Custom)
                {
                    throw(new Exception("Use the TotalsRowFormula-property to set a custom table formula"));
                }
                string s = value.ToString();
                s = s.Substring(0, 1).ToLower() + s.Substring(1, s.Length - 1);
                SetXmlNodeString("@totalsRowFunction", s);
            }
        }
        const string TOTALSROWFORMULA_PATH = "d:totalsRowFormula";
        /// <summary>
        /// Sets a custom Totals row Formula.
        /// Be carefull with this property since it is not validated. 
        /// <example>
        /// tbl.Columns[9].TotalsRowFormula = string.Format("SUM([{0}])",tbl.Columns[9].Name);
        /// </example>
        /// </summary>
        public string TotalsRowFormula
        {
            get
            {
                return GetXmlNodeString(TOTALSROWFORMULA_PATH);
            }
            set
            {
                if (value.StartsWith("=")) value = value.Substring(1, value.Length - 1);
                SetXmlNodeString("@totalsRowFunction", "custom");                
                SetXmlNodeString(TOTALSROWFORMULA_PATH, value);
            }
        }
        const string DATACELLSTYLE_PATH = "@dataCellStyle";
        /// <summary>
        /// The named style for datacells in the column
        /// </summary>
        public string DataCellStyleName
        {
            get
            {
                return GetXmlNodeString(DATACELLSTYLE_PATH);
            }
            set
            {
                if(_tbl.WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value)<0)
                {
                    throw(new Exception(string.Format("Named style {0} does not exist.",value)));
                }
                SetXmlNodeString(TopNode, DATACELLSTYLE_PATH, value,true);
               
                int fromRow=_tbl.Address._fromRow + (_tbl.ShowHeader?1:0),
                    toRow=_tbl.Address._toRow - (_tbl.ShowTotal?1:0),
                    col=_tbl.Address._fromCol+Position;

                if (fromRow < toRow)
                {
                    _tbl.WorkSheet.Cells[fromRow, col, toRow, col].StyleName = value;
                }
            }
        }
  		const string CALCULATEDCOLUMNFORMULA_PATH = "d:calculatedColumnFormula";
 		/// <summary>
 		/// Sets a calculated column Formula.
 		/// Be carefull with this property since it is not validated. 
 		/// <example>
 		/// tbl.Columns[9].CalculatedColumnFormula = string.Format("SUM(MyDataTable[[#This Row],[{0}]])",tbl.Columns[9].Name);
 		/// </example>
 		/// </summary>
 		public string CalculatedColumnFormula
 		{
 			get
 			{
 				return GetXmlNodeString(CALCULATEDCOLUMNFORMULA_PATH);
 			}
 			set
 			{
 				if (value.StartsWith("=")) value = value.Substring(1, value.Length - 1);
 				SetXmlNodeString(CALCULATEDCOLUMNFORMULA_PATH, value);
 			}
 		}

    }
}
