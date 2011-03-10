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
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>
    {
        protected ExcelPivotTable _table;
        protected List<T> _list = new List<T>();
        internal ExcelPivotTableFieldCollectionBase(ExcelPivotTable table)
        {
            _table = table;
        }
        public IEnumerator<T> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }
        public int Count
        {
            get
            {
                return _list.Count;
            }
        }
        internal void AddInternal(T field)
        {
            _list.Add(field);
        }
        internal void Clear()
        {
            _list.Clear();
        }
        public T this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _list.Count)
                {
                    throw (new ArgumentOutOfRangeException("PivotTable field index out of range"));
                }
                return _list[Index];
            }
        }
    }
    public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
    {
        internal string _topNode;
        public ExcelPivotTableFieldCollection(ExcelPivotTable table, string topNode) :
            base(table)
	    {
            _topNode=topNode;
	    }

        public void Add(ExcelPivotTableField Field)
        {
            SetFlag(Field, true);
            _list.Add(Field);
        }
        internal void Insert(ExcelPivotTableField Field, int Index)
        {
            SetFlag(Field, true);
            _list.Insert(Index, Field);
        }
        private void SetFlag(ExcelPivotTableField field, bool value)
        {
            switch (_topNode)
            {
                case "rowFields":
                    if (field.IsColumnField || field.IsPageField)
                    {
                        throw(new Exception("This field is a column or page field. Can's add it to the RowFields collection"));
                    }
                    field.IsRowField = value;
                    field.Axis = ePivotFieldAxis.Row;
                    break;
                case "colFields":
                    if (field.IsRowField || field.IsPageField)
                    {
                        throw (new Exception("This field is a row or page field. Can's add it to the ColumnFields collection"));
                    }
                    field.IsColumnField = value;
                    field.Axis = ePivotFieldAxis.Column;
                    break;
                case "pageFields":
                    if (field.IsColumnField || field.IsRowField)
                    {
                        throw (new Exception("Field is a column or row field. Can's add it to the PageFields collection"));
                    }
                    if (_table.Address._fromRow < 3)
                    {
                        throw(new Exception(string.Format("A pivot table with page fields must be located above row 3. Currenct location is {0}", _table.Address.Address)));
                    }
                    field.IsPageField = value;
                    field.Axis = ePivotFieldAxis.Page;
                    break;
                case "dataFields":
                    
                    break;
            }
        }
        public void Remove(ExcelPivotTableField Field)
        {
            if(!_list.Contains(Field))
            {
                throw new ArgumentException("Field not in collection");
            }
            SetFlag(Field, false);            
            _list.Remove(Field);            
        }
        public void RemoveAt(int Index)
        {
            if (Index > -1 && Index < _list.Count)
            {
                throw(new IndexOutOfRangeException());
            }
            SetFlag(_list[Index], false);
            _list.RemoveAt(Index);      
        }
    }
    public class ExcelPivotTableDataFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableDataField>
    {
        public ExcelPivotTableDataFieldCollection(ExcelPivotTable table) :
            base(table)
        {

        }
        public ExcelPivotTableDataField Add(ExcelPivotTableField field)
        {
            var dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
            if (dataFieldsNode == null)
            {
                _table.CreateNode("d:dataFields");
                dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
            }

            XmlElement node = field.AppendField(dataFieldsNode, field.Index, "dataField", "fld");
            field.SetXmlNodeBool("@dataField", true,false);

            var dataField = new ExcelPivotTableDataField(field.NameSpaceManager, dataFieldsNode, field);
            _list.Add(dataField);
            return dataField;
        }
        public void Remove(ExcelPivotTableDataField dataField)
        {
            XmlElement node = dataField.Field.TopNode.SelectSingleNode(string.Format("../../d:dataFields/d:dataField[@fld={0}]", dataField.Index), dataField.NameSpaceManager) as XmlElement;
            if (node != null)
            {
                node.ParentNode.RemoveChild(node);
            }
            _list.Remove(dataField);
        }
    }
}
