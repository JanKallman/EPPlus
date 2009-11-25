/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
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
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/

/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * ExcelPackage provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/ExcelPackage for details.
 * 
 * Copyright 2007 © Dr John Tunnicliffe 
 * mailto:dr.john.tunnicliffe@btinternet.com
 * All rights reserved.
 * 
 * ExcelPackage is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
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
 */

/*
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * John Tunnicliffe		Initial Release		01-Jan-2007
 * Jan Källman          Dont access the XML Direct anymore.     05-Oct-2009
 * ******************************************************************************
 */
using System;
using System.Xml;
using OfficeOpenXml.Style;
namespace OfficeOpenXml
{
	/// <summary>
	/// Represents an individual row in the spreadsheet.
	/// </summary>
	public class ExcelRow
	{
		private ExcelWorksheet _xlWorksheet;
		private XmlElement _rowElement = null;

        public ulong RowID 
        {
            get
            {
                return GetRowID(_xlWorksheet.SheetID, Row);
            }
        }
		#region ExcelRow Constructor
		/// <summary>
		/// Creates a new instance of the ExcelRow class. 
		/// For internal use only!
		/// </summary>
		/// <param name="Worksheet">The parent worksheet</param>
		/// <param name="row">The row number</param>
		protected internal ExcelRow(ExcelWorksheet Worksheet, int row)
		{
			_xlWorksheet = Worksheet;
            Row = row;
            Height = _xlWorksheet.defaultRowHeight;            
		}
		#endregion

		/// <summary>
		/// Provides access to the node representing the row.
		/// For internal use only!
		/// </summary>
		protected internal XmlNode Node { get { return (_rowElement); } }

		#region ExcelRow Hidden
		/// <summary>
		/// Allows the row to be hidden in the worksheet
		/// </summary>
		public bool Hidden
        {
            get;
            set;
        }        
		#endregion

		#region ExcelRow Height
		/// <summary>
		/// Sets the height of the row
		/// </summary>
		public double Height
        {
            get;
            set;
        }
		#endregion

        string _styleName = "";
        /// <summary>
        /// Sets the style for the entire column using a style name.
        /// </summary>
        public string StyleName
        {
            get
            {
                return _styleName;
            }
            set
            {
                _styleId = _xlWorksheet.Workbook.Styles.GetStyleIdFromName(value);
                _styleName = value;
            }
        }
        
        int _styleId = 0;
		/// <summary>
		/// Sets the style for the entire row using the style ID.  
		/// </summary>
        public int StyleID
		{
			get
			{
				return _styleId; 
			}
			set	
			{
                _styleId = value;
			}
		}

        /// <summary>
        /// Rownumber
        /// </summary>
        public int Row
        {
            get;
            set;
        }
        /// <summary>
        /// If outline level is set this tells that the row is collapsed
        /// </summary>
        public bool Collapsed
        {
            get;
            set;
        }
        /// <summary>
        /// Outline level.
        /// </summary>
        public int OutlineLevel
        {
            get;
            set;
        }        
        /// <summary>
        /// Show phonetic Information
        /// </summary>
        public bool Phonetic 
        {
            get;
            set;
        }
        /// <summary>
        /// Reference to style
        /// </summary>
        ExcelStyle _style = null;
        /// <summary>
        /// The Style applied to the whole row. Only effekt cells with no individual style set. 
        /// Use ExcelRange object if you want to set specific styles.
        /// </summary>
        public ExcelStyle Style
        {
            get
            {
                return _xlWorksheet.Workbook.Styles.GetStyleObject(StyleID,_xlWorksheet.PositionID ,Row.ToString()+":"+Row.ToString());                
            }
        }

        internal static ulong GetRowID(int sheetID, int row)
        {
            return ((ulong)sheetID) + (((ulong)row) << 29);

        }
    }
}
