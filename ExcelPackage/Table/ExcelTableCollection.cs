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
using System.IO.Packaging;
using System.Xml;

namespace OfficeOpenXml.Table
{
    public class ExcelTableCollection : IEnumerable<ExcelTable>
    {
        List<ExcelTable> _tables = new List<ExcelTable>();
        internal Dictionary<string, int> _tableNames = new Dictionary<string, int>();
        ExcelWorksheet _ws;        
        internal ExcelTableCollection(ExcelWorksheet ws)
        {
            Package pck = ws.xlPackage.Package;
            _ws = ws;
            foreach(XmlElement node in ws.WorksheetXml.SelectNodes("//d:tableParts/d:tablePart", ws.NameSpaceManager))
            {
                var rel = ws.Part.GetRelationship(node.GetAttribute("id",ExcelPackage.schemaRelationships));
                var tbl = new ExcelTable(rel, ws);
            }
        }
        private ExcelTable Add(ExcelTable tbl)
        {
            _tables.Add(tbl);
            _tableNames.Add(tbl.Name, _tables.Count - 1);
            if (tbl.Id >= _ws.Workbook._nextTableID)
            {
                _ws.Workbook._nextTableID = tbl.Id + 1;
            }
            return tbl;
        }

        /// <summary>
        /// Create a table on the supplied range
        /// </summary>
        /// <param name="Range">The range address including header and total row</param>
        /// <param name="Name">The name of the table. Must be unique </param>
        /// <returns>The table object</returns>
        public ExcelTable Add(ExcelAddressBase Range, string Name)
        {
            if (string.IsNullOrEmpty(Name))
            {
                Name = "Table1";
                int i=2;
                while (_ws.Workbook.ExistsTableName(Name))
                {
                    Name=string.Format("Table{0}", i++);
                }
            }
            else if (_ws.Workbook.ExistsTableName(Name))
            {
                throw (new ArgumentException("Tablename is not unique"));
            }
            foreach (var t in _tables)
            {
                if (t.Address.Collide(Range) != ExcelAddressBase.eAddressCollition.No)
                {
                    throw (new ArgumentException(string.Format("Table range collides with table {0}", t.Name)));
                }
            }
            return Add(new ExcelTable(_ws, Range, Name, _ws.Workbook._nextTableID));
            //var tbl = new ExcelTable(_ws, Range, Name, _ws.Workbook._nextTableID++);
            //_tables.Add(tbl);
            //_tableNames.Add(Name, _tables.Count - 1);
            //return tbl;
        }
        public int Count
        {
            get
            {
                return _tables.Count;
            }
        }
        /// <summary>
        /// The table Index. Base 0.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelTable this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _tables.Count)
                {
                    throw (new ArgumentOutOfRangeException("Table index out of range"));
                }
                return _tables[Index];
            }
        }
        public ExcelTable this[string Name]
        {
            get
            {
                if (_tableNames.ContainsKey(Name))
                {
                    return _tables[_tableNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }
        public IEnumerator<ExcelTable> GetEnumerator()
        {
            return _tables.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _tables.GetEnumerator();
        }
    }
}
