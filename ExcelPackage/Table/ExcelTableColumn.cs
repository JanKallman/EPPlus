/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * See http://epplus.codeplex.com/ for details
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		13-SEP-2010
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
        public int Position
        {
            get;
            private set;
        }
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
    }
}
