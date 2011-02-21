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
using System.Globalization;

namespace OfficeOpenXml.Table.PivotTable
{
    
    /// <summary>
    /// defines the axis for a PivotTable
    /// </summary>
    public enum ePivotFieldAxis
    {
        None=-1,
        /// <summary>
        /// Column axis
        /// </summary>
        Column,
        /// <summary>
        /// Page axis (Include Count Filter) 
        /// </summary>
        Page,
        /// <summary>
        /// Row axis
        /// </summary>
        Row,
        /// <summary>
        /// Values axis
        /// </summary>
        Values 
    }
    /// <summary>
    /// Build-in table row functions
    /// </summary>
    public enum DataFieldFunctions
    {
        Average,
        Count,
        CountNums,
        Max,
        Min,
        Product,
        None,
        StdDev,
        StdDevP,
        Sum,
        Var,
        VarP
    }
    public class ExcelPivotTableField : XmlHelper
    {
        internal ExcelPivotTable _table;
        public ExcelPivotTableField(XmlNamespaceManager ns, XmlNode topNode,ExcelPivotTable table, int index) :
            base(ns, topNode)
        {
            Index = index;
            _table = table;
        }
        public int Index
        {
            get;
            set;
        }
        public bool Compact
        { 
            get
            {
                return GetXmlNodeBool("@compact");
            }
            set
            {
                SetXmlNodeBool("@compact",value);
            }
        }
        public bool Outline 
        { 
            get
            {
                return GetXmlNodeBool("@outline");
            }
            set
            {
                SetXmlNodeBool("@outline",value);
            }
        }
        public bool SubtotalTop 
        { 
            get
            {
                return GetXmlNodeBool("@subtotalTop");
            }
            set
            {
                SetXmlNodeBool("@subtotalTop",value);
            }
        }
        public bool ShowAll 
        { 
            get
            {
                return GetXmlNodeBool("@showAll");
            }
            set
            {
                SetXmlNodeBool("@showAll",value);
            }
        }
        public bool IncludeNewItemsInFilter
        { 
            get
            {
                return GetXmlNodeBool("@includeNewItemsInFilter");
            }
            set
            {
                SetXmlNodeBool("@includeNewItemsInFilter",value);
            }
        }
        public ePivotFieldAxis Axis
        {
            get
            {
                switch(GetXmlNodeString("@axis"))
                {
                    case "axisRow":
                        return ePivotFieldAxis.Row;
                    case "axisCol":
                        return ePivotFieldAxis.Column;
                    case "axisPage":
                        return ePivotFieldAxis.Page;
                    case "axisValues":
                        return ePivotFieldAxis.Values;
                    default:
                        return ePivotFieldAxis.None;
                }
            }
            internal set
            {
                switch (value)
                {
                    case ePivotFieldAxis.Row:
                        SetXmlNodeString("@axis","axisRow");
                        break;
                    case ePivotFieldAxis.Column:
                        SetXmlNodeString("@axis","axisCol");
                        break;
                    case ePivotFieldAxis.Values:
                        SetXmlNodeString("@axis", "axisValues");
                        break;
                    case ePivotFieldAxis.Page:
                        SetXmlNodeString("@axis", "axisPage");
                        break;
                    default:
                        DeleteNode("@axis");
                        break;
                }
            }
        }        
        public bool IsRowField
        {
            get
            {
                return (TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", Index), NameSpaceManager) != null);
            }
            internal set
            {
                if (value)
                {
                    var rowsNode = TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);
                    if (rowsNode == null)
                    {
                        _table.CreateNode("d:rowFields");
                    }
                    rowsNode = TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);

                    AppendField(rowsNode, Index, "field", "x");
                    TopNode.InnerXml="<items count=\"1\"><item t=\"default\" /></items>";
                }
                else
                {
                    XmlElement node = TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", Index), NameSpaceManager) as XmlElement;
                     if (node != null)
                     {
                         node.ParentNode.RemoveChild(node);
                     }
                }
            }
        }
        public bool IsColumnField
        {
            get
            {
                return (TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", Index), NameSpaceManager) != null);
            }
            internal set
            {
                if (value)
                {
                    var columnsNode = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);
                    if (columnsNode == null)
                    {
                        _table.CreateNode("d:colFields");
                    }
                    columnsNode = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);

                    AppendField(columnsNode, Index, "field", "x");
                    TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
                }
                else
                {
                    XmlElement node = TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", Index), NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
        }
        public bool IsDataField
        {
            get
            {
                return GetXmlNodeBool("@dataField", false);
            }
            internal set
            {
                if (value)
                {
                    if (_dataFieldSettings != null) return;
                    var dataFieldsNode = TopNode.SelectSingleNode("../../d:dataFields", NameSpaceManager);
                    if (dataFieldsNode == null)
                    {
                        _table.CreateNode("d:dataFields");
                        dataFieldsNode = TopNode.SelectSingleNode("../../d:dataFields", NameSpaceManager);
                    }

                    XmlElement node = AppendField(dataFieldsNode, Index, "dataField", "fld");
                    _dataFieldSettings = new ExcelPivotTableDataFieldSettings(NameSpaceManager, node, this, Index);
                    _pageFieldSettings = null;
                }
                else
                {
                    _dataFieldSettings = null;
                    XmlElement node = TopNode.SelectSingleNode(string.Format("../../d:dataFields/d:dataField[@fld={0}]", Index), NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
                SetXmlNodeBool("@dataField",value,false);
            }
        }
        public bool IsPageField
        {
            get
            {
                return (Axis==ePivotFieldAxis.Page);
            }
            internal set
            {
                if (value)
                {
                    if (_dataFieldSettings != null) return;
                    var dataFieldsNode = TopNode.SelectSingleNode("../../d:pageFields", NameSpaceManager);
                    if (dataFieldsNode == null)
                    {
                        _table.CreateNode("d:pageFields");
                        dataFieldsNode = TopNode.SelectSingleNode("../../d:pageFields", NameSpaceManager);
                    }

                    TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";

                    XmlElement node = AppendField(dataFieldsNode, Index, "pageField", "fld");
                    _pageFieldSettings = new ExcelPivotTablePageFieldSettings(NameSpaceManager, node, this, Index);
                    _dataFieldSettings = null;
                }
                else
                {
                    _pageFieldSettings = null;
                    XmlElement node = TopNode.SelectSingleNode(string.Format("../../d:pageFields/d:pageField[@fld={0}]", Index), NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
        }
        //public ExcelPivotGrouping DateGrouping
        //{

        //}
        internal ExcelPivotTableDataFieldSettings _dataFieldSettings = null;
        public ExcelPivotTableDataFieldSettings DataFieldSettings
        {
            get
            {
                return _dataFieldSettings;
            }
        }
        internal ExcelPivotTablePageFieldSettings _pageFieldSettings = null;
        public ExcelPivotTablePageFieldSettings PageFieldSettings
        {
            get
            {
                return _pageFieldSettings;
            }
        }
        internal ExcelPivotTableFieldGroupCollection _groups = null;
        public ExcelPivotTableFieldGroupCollection Grouping
        {
            get
            {
                if (_groups == null)
                {
                    _groups = new ExcelPivotTableFieldGroupCollection(this);
                }
                return _groups;
            }
        }
        ExcelPivotTableFieldGroup _grouping=null;
        public ExcelPivotTableFieldGroup Group
        {
            get
            {
                return _grouping;
            }
        }
        #region Private & internal Methods
        private XmlElement AppendField(XmlNode rowsNode, int index, string fieldNodeText, string indexAttrText)
        {
            XmlElement prevField = null, newElement;
            foreach (XmlElement field in rowsNode.ChildNodes)
            {
                string x = field.GetAttribute(indexAttrText);
                int fieldIndex;
                if(int.TryParse(x, out fieldIndex))
                {
                    if (fieldIndex == index)    //Row already exist
                    {
                        return field;
                    }
                    else if (fieldIndex > index)
                    {
                        newElement = rowsNode.OwnerDocument.CreateElement(fieldNodeText, ExcelPackage.schemaMain);
                        newElement.SetAttribute(indexAttrText, index.ToString());
                        rowsNode.InsertAfter(newElement, field);
                    }
                }
                prevField=field;
            }
            newElement = rowsNode.OwnerDocument.CreateElement(fieldNodeText, ExcelPackage.schemaMain);
            newElement.SetAttribute(indexAttrText, index.ToString());
            rowsNode.InsertAfter(newElement, prevField);

            return newElement;
        }
        internal XmlHelperInstance _cacheFieldHelper = null;
        internal void SetCacheFieldNode(XmlNode cacheField)
        {
            _cacheFieldHelper = new XmlHelperInstance(NameSpaceManager, cacheField);
        }
        #endregion
        #region Grouping
        public ExcelPivotTableFieldGroup SetDateGroup(eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate)
        {
            //foreach (var grp in _list)
            //{
            //    if (grp.GroupBy == GroupBy)
            //    {
            //        throw (new ArgumentException("Grouping already exist in collection"));
            //    }
            //}
            ExcelPivotTableFieldGroup group;
            //if (_list.Count == 0)
            //{
                group = new ExcelPivotTableFieldGroup(NameSpaceManager, _cacheFieldHelper.TopNode, GroupBy);
                _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsDate", true);
                _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);

                group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"{1}\" /><groupItems /></fieldGroup>", Index.ToString(), GroupBy.ToString().ToLower());
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", StartDate.ToString("s", CultureInfo.InvariantCulture));
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", EndDate.ToString("s", CultureInfo.InvariantCulture));
                int items = AddGroupItems(group, GroupBy, StartDate, EndDate);
                AddFieldItems(items);
            //}
            //else
            //{
            //    var cacheXml = _table.CacheDefinition.CacheDefinitionXml;
            //    var fields = cacheXml.DocumentElement.SelectSingleNode("d:cacheFields\\d:cacheFields");
            //    var node = cacheXml.CreateElement("cacheField", ExcelPackage.schemaMain);

            //    node.SetAttribute("Name", GroupBy.ToString());
            //    node.SetAttribute("numFmtId", "0");
            //    node.SetAttribute("databaseField", Index.ToString());
            //    fields.AppendChild(node);
            //    group = new ExcelPivotTableFieldGroup(NameSpaceManager, node, GroupBy);
            //    group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"quarters\" /></fieldGroup>", Index.ToString());
            //}
            //_list.Add(group);
            return group;
        }

        private void AddFieldItems(int items)
        {
            XmlElement prevNode = null;
            XmlElement itemsNode = TopNode.SelectSingleNode("//d:items", NameSpaceManager) as XmlElement;
            for (int x = 0; x < items; x++)
            {
                var itemNode = itemsNode.OwnerDocument.CreateElement("item", ExcelPackage.schemaMain);
                itemNode.SetAttribute("x", x.ToString());
                if (prevNode == null)
                {
                    itemsNode.PrependChild(itemNode);
                }
                else
                {
                    itemsNode.InsertAfter(itemNode, prevNode);
                }
                prevNode = itemNode;
            }
            itemsNode.SetAttribute("count", (items + 1).ToString());
        }

        private int AddGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate)
        {
            XmlElement groupItems = group.TopNode.SelectSingleNode("//d:groupItems", group.NameSpaceManager) as XmlElement;
            int items = 2;
            //First date
            AddGroupItem(groupItems, "<" + StartDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));

            switch (GroupBy)
            {
                case eDateGroupBy.Seconds:
                case eDateGroupBy.Minutes:
                    AddTimeSerie(60, groupItems);
                    items += 60;
                    break;
                case eDateGroupBy.Hours:
                    AddTimeSerie(24, groupItems);
                    items += 60;
                    break;
                case eDateGroupBy.Days:
                    DateTime dt = new DateTime(2008, 1, 1); //pick a year with 366 days
                    while (dt.Year == 2008)
                    {
                        AddGroupItem(groupItems, dt.ToString("dd-MMM"));
                        dt = dt.AddDays(1);
                    }
                    items += 366;
                    break;
                case eDateGroupBy.Months:
                    AddGroupItem(groupItems, "jan");
                    AddGroupItem(groupItems, "feb");
                    AddGroupItem(groupItems, "mar");
                    AddGroupItem(groupItems, "apr");
                    AddGroupItem(groupItems, "may");
                    AddGroupItem(groupItems, "jun");
                    AddGroupItem(groupItems, "jul");
                    AddGroupItem(groupItems, "aug");
                    AddGroupItem(groupItems, "sep");
                    AddGroupItem(groupItems, "oct");
                    AddGroupItem(groupItems, "nov");
                    AddGroupItem(groupItems, "dec");
                    items += 12;
                    break;
                case eDateGroupBy.Quarters:
                    AddGroupItem(groupItems, "Qtr1");
                    AddGroupItem(groupItems, "Qtr2");
                    AddGroupItem(groupItems, "Qtr3");
                    AddGroupItem(groupItems, "Qtr4");
                    items += 4;
                    break;
                case eDateGroupBy.Years:
                    for (int year = StartDate.Year; year <= EndDate.Year; year++)
                    {
                        AddGroupItem(groupItems, year.ToString());
                    }
                    break;
                default:
                    throw (new Exception("unsupported grouping"));
            }

            //Lastdate
            AddGroupItem(groupItems, ">" + EndDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));
            return items;
        }

        private void AddTimeSerie(int count, XmlElement groupItems)
        {
            for (int i = 0; i < count; i++)
            {
                AddGroupItem(groupItems, string.Format(":{0:00}", count));
            }
        }

        private void AddGroupItem(XmlElement groupItems, string value)
        {
            var s = groupItems.OwnerDocument.CreateElement("s", ExcelPackage.schemaMain);
            s.SetAttribute("v", value);
            groupItems.AppendChild(s);
        }
        #endregion
    }
}
