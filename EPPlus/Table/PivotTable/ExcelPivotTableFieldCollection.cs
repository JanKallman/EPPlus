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
 * Jan Källman		Added		21-MAR-2011
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Base collection class for pivottable fields
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>
    {
        protected ExcelPivotTable _table;
        internal List<T> _list = new List<T>();
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
                    throw (new ArgumentOutOfRangeException("Index out of range"));
                }
                return _list[Index];
            }
        }
    }
    public class ExcelPivotTableFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
    {
        internal ExcelPivotTableFieldCollection(ExcelPivotTable table, string topNode) :
            base(table)
        {

        }
        /// <summary>
        /// Indexer by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public ExcelPivotTableField this[string name]
        {
            get
            {
                foreach (var field in _list)
                {
                    if (field.Name.Equals(name,StringComparison.InvariantCultureIgnoreCase))
                    {
                        return field;
                    }
                }
                return null;
            }
        }
        /// <summary>
        /// Returns the date group field.
        /// </summary>
        /// <param name="GroupBy">The type of grouping</param>
        /// <returns>The matching field. If none is found null is returned</returns>
        public ExcelPivotTableField GetDateGroupField(eDateGroupBy GroupBy)
        {
            foreach (var fld in _list)
            {
                if (fld.Grouping is ExcelPivotTableFieldDateGroup && (((ExcelPivotTableFieldDateGroup)fld.Grouping).GroupBy) == GroupBy)
                {
                    return fld;
                }
            }
            return null;
        }
        /// <summary>
        /// Returns the numeric group field.
        /// </summary>
        /// <returns>The matching field. If none is found null is returned</returns>
        public ExcelPivotTableField GetNumericGroupField()
        {
            foreach (var fld in _list)
            {
                if (fld.Grouping is ExcelPivotTableFieldNumericGroup)
                {
                    return fld;
                }
            }
            return null;
        }
    }
    /// <summary>
    /// Collection class for Row and column fields in a Pivottable 
    /// </summary>
    public class ExcelPivotTableRowColumnFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableField>
    {
        internal string _topNode;
        internal ExcelPivotTableRowColumnFieldCollection(ExcelPivotTable table, string topNode) :
            base(table)
	    {
            _topNode=topNode;
	    }

        /// <summary>
        /// Add a new row/column field
        /// </summary>
        /// <param name="Field">The field</param>
        /// <returns>The new field</returns>
        public ExcelPivotTableField Add(ExcelPivotTableField Field)
        {
            SetFlag(Field, true);
            _list.Add(Field);
            return Field;
        }
        /// <summary>
        /// Insert a new row/column field
        /// </summary>
        /// <param name="Field">The field</param>
        /// <param name="Index">The position to insert the field</param>
        /// <returns>The new field</returns>
        internal ExcelPivotTableField Insert(ExcelPivotTableField Field, int Index)
        {
            SetFlag(Field, true);
            _list.Insert(Index, Field);
            return Field;
        }
        private void SetFlag(ExcelPivotTableField field, bool value)
        {
            switch (_topNode)
            {
                case "rowFields":
                    if (field.IsColumnField || field.IsPageField)
                    {
                        throw(new Exception("This field is a column or page field. Can't add it to the RowFields collection"));
                    }
                    field.IsRowField = value;
                    field.Axis = ePivotFieldAxis.Row;
                    break;
                case "colFields":
                    if (field.IsRowField || field.IsPageField)
                    {
                        throw (new Exception("This field is a row or page field. Can't add it to the ColumnFields collection"));
                    }
                    field.IsColumnField = value;
                    field.Axis = ePivotFieldAxis.Column;
                    break;
                case "pageFields":
                    if (field.IsColumnField || field.IsRowField)
                    {
                        throw (new Exception("Field is a column or row field. Can't add it to the PageFields collection"));
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
        /// <summary>
        /// Remove a field
        /// </summary>
        /// <param name="Field"></param>
        public void Remove(ExcelPivotTableField Field)
        {
            if(!_list.Contains(Field))
            {
                throw new ArgumentException("Field not in collection");
            }
            SetFlag(Field, false);            
            _list.Remove(Field);            
        }
        /// <summary>
        /// Remove a field at a specific position
        /// </summary>
        /// <param name="Index"></param>
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
    /// <summary>
    /// Collection class for data fields in a Pivottable 
    /// </summary>
    public class ExcelPivotTableDataFieldCollection : ExcelPivotTableFieldCollectionBase<ExcelPivotTableDataField>
    {
        internal ExcelPivotTableDataFieldCollection(ExcelPivotTable table) :
            base(table)
        {

        }
        /// <summary>
        /// Add a new datafield
        /// </summary>
        /// <param name="field">The field</param>
        /// <returns>The new datafield</returns>
        public ExcelPivotTableDataField Add(ExcelPivotTableField field)
        {
            var dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
            if (dataFieldsNode == null)
            {
                _table.CreateNode("d:dataFields");
                dataFieldsNode = field.TopNode.SelectSingleNode("../../d:dataFields", field.NameSpaceManager);
            }

            XmlElement node = _table.PivotTableXml.CreateElement("dataField", ExcelPackage.schemaMain);
            node.SetAttribute("fld", field.Index.ToString());
            dataFieldsNode.AppendChild(node);

            //XmlElement node = field.AppendField(dataFieldsNode, field.Index, "dataField", "fld");
            field.SetXmlNodeBool("@dataField", true,false);

            var dataField = new ExcelPivotTableDataField(field.NameSpaceManager, node, field);
            ValidateDupName(dataField);

            _list.Add(dataField);
            return dataField;
        }
        private void ValidateDupName(ExcelPivotTableDataField dataField)
        {
            if(ExistsDfName(dataField.Field.Name, null))
            {
                var index = 2;
                string name;
                do
                {
                    name = dataField.Field.Name + "_" + index++.ToString();
                }
                while (ExistsDfName(name,null));
                dataField.Name = name;
            }
        }

        internal bool ExistsDfName(string name, ExcelPivotTableDataField datafield)
        {
            foreach (var df in _list)
            {
                if (((!string.IsNullOrEmpty(df.Name) && df.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase) ||
                     (string.IsNullOrEmpty(df.Name) && df.Field.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase)))) && datafield != df)
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// Remove a datafield
        /// </summary>
        /// <param name="dataField"></param>
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