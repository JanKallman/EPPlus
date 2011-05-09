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
 * Jan Källman		Added		21-MAR-2011
 *******************************************************************************/
using System;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO.Packaging;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A collection of pivottable objects
    /// </summary>
    public class ExcelPivotTableCollection : IEnumerable<ExcelPivotTable>
    {
        List<ExcelPivotTable> _pivotTables = new List<ExcelPivotTable>();
        internal Dictionary<string, int> _pivotTableNames = new Dictionary<string, int>();
        ExcelWorksheet _ws;        
        internal ExcelPivotTableCollection(ExcelWorksheet ws)
        {
            Package pck = ws.xlPackage.Package;
            _ws = ws;            
            foreach(var rel in ws.Part.GetRelationships())
            {
                if (rel.RelationshipType == ExcelPackage.schemaRelationships + "/pivotTable")
                {
                    var tbl = new ExcelPivotTable(rel, ws);
                    _pivotTableNames.Add(tbl.Name, _pivotTables.Count);
                    _pivotTables.Add(tbl);
                }
            }
        }
        private ExcelPivotTable Add(ExcelPivotTable tbl)
        {
            _pivotTables.Add(tbl);
            _pivotTableNames.Add(tbl.Name, _pivotTables.Count - 1);
            if (tbl.Id >= _ws.Workbook._nextTableID)
            {
                _ws.Workbook._nextTableID = tbl.Id + 1;
            }
            return tbl;
        }

        /// <summary>
        /// Create a pivottable on the supplied range
        /// </summary>
        /// <param name="Range">The range address including header and total row</param>
        /// <param name="Source">The Source data range address</param>
        /// <param name="Name">The name of the table. Must be unique </param>
        /// <returns>The table object</returns>
        public ExcelPivotTable Add(ExcelAddressBase Range, ExcelRangeBase Source, string Name)
        {
            if (string.IsNullOrEmpty(Name))
            {
                Name = GetNewTableName();
            }
            if (Range.WorkSheet != _ws.Name)
            {
                throw(new Exception("The Range must be in the current worksheet"));
            }
            else if (_ws.Workbook.ExistsTableName(Name))
            {
                throw (new ArgumentException("Tablename is not unique"));
            }
            foreach (var t in _pivotTables)
            {
                if (t.Address.Collide(Range) != ExcelAddressBase.eAddressCollition.No)
                {
                    throw (new ArgumentException(string.Format("Table range collides with table {0}", t.Name)));
                }
            }
            return Add(new ExcelPivotTable(_ws, Range, Source, Name, _ws.Workbook._nextPivotTableID++));
        }

        internal string GetNewTableName()
        {
            string name = "Pivottable1";
            int i = 2;
            while (_ws.Workbook.ExistsPivotTableName(name))
            {
                name = string.Format("Pivottable{0}", i++);
            }
            return name;
        }
        public int Count
        {
            get
            {
                return _pivotTables.Count;
            }
        }
        /// <summary>
        /// The table Index. Base 0.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public ExcelPivotTable this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _pivotTables.Count)
                {
                    throw (new ArgumentOutOfRangeException("PivotTable index out of range"));
                }
                return _pivotTables[Index];
            }
        }
        public ExcelPivotTable this[string Name]
        {
            get
            {
                if (_pivotTableNames.ContainsKey(Name))
                {
                    return _pivotTables[_pivotTableNames[Name]];
                }
                else
                {
                    return null;
                }
            }
        }
        public IEnumerator<ExcelPivotTable> GetEnumerator()
        {
            return _pivotTables.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _pivotTables.GetEnumerator();
        }
    }
}
