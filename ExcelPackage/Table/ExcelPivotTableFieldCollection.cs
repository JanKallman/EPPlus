using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.Table
{
    public class ExcelPivotTableFieldCollectionBase : IEnumerable<ExcelPivotTableField>
    {
        protected ExcelPivotTable _table;
        protected List<ExcelPivotTableField> _list = new List<ExcelPivotTableField>();
        public ExcelPivotTableFieldCollectionBase(ExcelPivotTable table)
        {
            _table = table;
        }
        public IEnumerator<ExcelPivotTableField> GetEnumerator()
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
        internal void AddInternal(ExcelPivotTableField field)
        {
            _list.Add(field);
        }

        internal void Clear()
        {
            _list.Clear();
        }
        public ExcelPivotTableField this[int Index]
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
    public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase
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

        private void SetFlag(ExcelPivotTableField field, bool value)
        {
            switch (_topNode)
            {
                case "rowFields":
                    if (field.IsColumnField)
                    {
                        throw(new Exception("This field is a column field. Can's add it to the RowFields collection"));
                    }
                    field.IsRowField = value;
                    field.Axis = ePivotFieldAxis.Row;
                    break;
                case "colFields":
                    if (field.IsRowField)
                    {
                        throw(new Exception("This field is a row field. Can's add it to the ColumnFields collection"));
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
                    field.IsDataField = value;
                    break;
            }
        }
        public void Remove(ExcelPivotTableField Field)
        {
            if(!_list.Contains(Field))
            {
                throw new ArgumentException("Field not in collection");
            }
            Field._dataFieldSettings = null;
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
}
