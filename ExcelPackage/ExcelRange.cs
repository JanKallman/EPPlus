/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
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
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Style;
using System.Data;
namespace OfficeOpenXml
{
    /// <summary>
    /// A range of cells. 
    /// </summary>
    public class ExcelRange : ExcelRangeBase
    {
        #region "Constructors"
        internal ExcelRange(ExcelWorksheet sheet) : 
            base(sheet)
        {

        }
        internal ExcelRange(ExcelWorksheet sheet, string address)
            : base(sheet, address)
        {

        }
        #endregion
        #region "Indexers"
        /// <summary>
        /// Access the range using an address
        /// </summary>
        /// <param name="Address">The address</param>
        /// <returns>A range object</returns>
        public ExcelRange this[string Address]
        {
            get
            {
                base.Address = Address;
                return this;
            }
        }
        /// <summary>
        /// Access a single cell
        /// </summary>
        /// <param name="Row">The row</param>
        /// <param name="Col">The column</param>
        /// <returns>A range object</returns>
        public ExcelRange this[int Row, int Col]
        {
            get
            {
                _fromCol = Col;
                _fromRow = Row;
                _toCol = Col;
                _toRow = Row;
                base.Address = GetAddress(_fromRow, _fromCol);
                return this;
            }
        }
        /// <summary>
        /// Access a range of cells
        /// </summary>
        /// <param name="FromRow">Start row</param>
        /// <param name="FromCol">Start column</param>
        /// <param name="ToRow">End Row</param>
        /// <param name="ToCol">End Column</param>
        /// <returns></returns>
        public ExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol]
        {
            get
            {
                _fromCol = FromCol;
                _fromRow = FromRow;
                _toCol = ToCol;
                _toRow = ToRow;
                base.Address = GetAddress(_fromRow, _fromCol, _toRow, _toCol);
                return this;
            }
        }
        #endregion
    }
}
