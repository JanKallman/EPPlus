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
    /// A collection of table columns
    /// </summary>
    public class ExcelTableColumnCollection : IEnumerable<ExcelTableColumn>
    {
        List<ExcelTableColumn> _cols = new List<ExcelTableColumn>();
        Dictionary<string, int> _colNames = new Dictionary<string, int>();
        public ExcelTableColumnCollection(ExcelTable table)
        {
            Table = table;
            foreach(XmlNode node in table.TableXml.SelectNodes("//d:table/d:tableColumns/d:tableColumn",table.NameSpaceManager))
            {
                _cols.Add(new ExcelTableColumn(table.NameSpaceManager, node, table, _cols.Count));
                _colNames.Add(_cols[_cols.Count - 1].Name, _cols.Count - 1);
            }
        }
        public ExcelTable Table
        {
            get;
            private set;
        }
        public int Count
        {
            get
            {
                return _cols.Count;
            }
        }
        /// <summary>
        /// The column Index. Base 0.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelTableColumn this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _cols.Count)
                {
                    throw (new ArgumentOutOfRangeException("Column index out of range"));
                }
                return _cols[Index] as ExcelTableColumn;
            }
        }
        public ExcelTableColumn this[string Name]
        {
            get
            {
                if (_colNames.ContainsKey(Name))
                {
                    return _cols[_colNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }

        IEnumerator<ExcelTableColumn> IEnumerable<ExcelTableColumn>.GetEnumerator()
        {
            return _cols.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _cols.GetEnumerator();
        }
    }
}
