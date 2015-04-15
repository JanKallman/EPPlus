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
using System.Globalization;

namespace OfficeOpenXml.Table.PivotTable
{
    
    /// <summary>
    /// Defines the axis for a PivotTable
    /// </summary>
    public enum ePivotFieldAxis
    {
        /// <summary>
        /// None
        /// </summary>
        None=-1,
        /// <summary>
        /// Column axis
        /// </summary>
        Column,
        /// <summary>
        /// Page axis (Include Count Filter) 
        /// 
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
    /// <summary>
    /// Defines the data formats for a field in the PivotTable
    /// </summary>
    public enum eShowDataAs
    {
        /// <summary>
        /// Indicates the field is shown as the "difference from" a value.
        /// </summary>
        Difference,
        /// <summary>
        /// Indicates the field is shown as the "index.
        /// </summary>
        Index, 
        /// <summary>
        /// Indicates that the field is shown as its normal datatype.
        /// </summary>
        Normal, 
        /// <summary>
        /// /Indicates the field is show as the "percentage of" a value
        /// </summary>
        Percent, 
        /// <summary>
        /// Indicates the field is shown as the "percentage difference from" a value.
        /// </summary>
        PercentDiff, 
        /// <summary>
        /// Indicates the field is shown as the percentage of column.
        /// </summary>
        PercentOfCol,
        /// <summary>
        /// Indicates the field is shown as the percentage of row
        /// </summary>
        PercentOfRow, 
        /// <summary>
        /// Indicates the field is shown as percentage of total.
        /// </summary>
        PercentOfTotal, 
        /// <summary>
        /// Indicates the field is shown as running total in the table.
        /// </summary>
        RunTotal,        
    }
      /// <summary>
     /// Built-in subtotal functions
     /// </summary>
    [Flags] 
    public enum eSubTotalFunctions 
     {
         None=1,
         Count=2, 
         CountA=4, 
         Avg=8, 
         Default=16, 
         Min=32, 
         Max=64, 
         Product=128, 
         StdDev=256, 
         StdDevP=512, 
         Sum=1024, 
         Var=2048, 
         VarP=4096
     }
    /// <summary>
    /// Data grouping
    /// </summary>
    [Flags]
    public enum eDateGroupBy
    {
        Years = 1,
        Quarters = 2,
        Months = 4,
        Days = 8,
        Hours = 16,
        Minutes = 32,
        Seconds = 64
    }
    /// <summary>
    /// Sorting
    /// </summary>
    public enum eSortType
    {
        None,
        Ascending,
        Descending
    }
    /// <summary>
    /// A pivot table field.
    /// </summary>
    public class ExcelPivotTableField : XmlHelper
    {
        internal ExcelPivotTable _table;
        internal ExcelPivotTableField(XmlNamespaceManager ns, XmlNode topNode,ExcelPivotTable table, int index, int baseIndex) :
            base(ns, topNode)
        {
            Index = index;
            BaseIndex = baseIndex;
            _table = table;
        }
        public int Index
        {
            get;
            set;
        }
        internal int BaseIndex
        {
            get;
            set;
        }
        /// <summary>
        /// Name of the field
        /// </summary>
        public string Name 
        { 
            get
            {
                string v = GetXmlNodeString("@name");
                if (v == "")
                {
                    return _cacheFieldHelper.GetXmlNodeString("@name");
                }
                else
                {
                    return GetXmlNodeString("@name");
                }
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// Compact mode
        /// </summary>
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
        /// <summary>
        /// A boolean that indicates whether the items in this field should be shown in Outline form
        /// </summary>
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
        /// <summary>
        /// The custom text that is displayed for the subtotals label
        /// </summary>
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
        /// <summary>
        /// A boolean that indicates whether to show all items for this field
        /// </summary>
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
        /// <summary>
        /// The type of sort that is applied to this field
        /// </summary>
        public eSortType Sort
        {
            get
            {
                string v = GetXmlNodeString("@sortType");
                return v == "" ? eSortType.None : (eSortType)Enum.Parse(typeof(eSortType), v, true);
            }
            set
            {
                if (value == eSortType.None)
                {
                    DeleteNode("@sortType");
                }
                else
                {
                    SetXmlNodeString("@sortType", value.ToString().ToLower(CultureInfo.InvariantCulture));
                }
            }
        }
        /// <summary>
        /// A boolean that indicates whether manual filter is in inclusive mode
        /// </summary>
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
         /// <summary>
         /// Enumeration of the different subtotal operations that can be applied to page, row or column fields
         /// </summary>
         public eSubTotalFunctions SubTotalFunctions
         {
            get
            {
                eSubTotalFunctions ret = 0;
                XmlNodeList nl = TopNode.SelectNodes("d:items/d:item/@t", NameSpaceManager);
                if (nl.Count == 0) return eSubTotalFunctions.None;
                foreach (XmlAttribute item in nl)
                {
                    try
                    {
                        ret |= (eSubTotalFunctions)Enum.Parse(typeof(eSubTotalFunctions), item.Value, true);
                    }
                    catch (ArgumentException ex)
                    {
                        throw new ArgumentException("Unable to parse value of " + item.Value + " to a valid pivot table subtotal function", ex);
                    }
                }
                return ret;
             }
             set
             {
                 if ((value & eSubTotalFunctions.None) == eSubTotalFunctions.None && (value != eSubTotalFunctions.None))
                 {
                     throw (new ArgumentException("Value None can not be combined with other values."));
                 }
                 if ((value & eSubTotalFunctions.Default) == eSubTotalFunctions.Default && (value != eSubTotalFunctions.Default))
                 {
                     throw (new ArgumentException("Value Default can not be combined with other values."));
                 }


                 // remove old attribute                 
                 XmlNodeList nl = TopNode.SelectNodes("d:items/d:item/@t", NameSpaceManager);
                 if (nl.Count > 0)
                 {
                     foreach (XmlAttribute item in nl)
                     {
                         DeleteNode("@" + item.Value + "Subtotal");
                         item.OwnerElement.ParentNode.RemoveChild(item.OwnerElement);
                     }
                 }
                 
 
                 if (value==eSubTotalFunctions.None)
                 {
                     // for no subtotals, set defaultSubtotal to off
                     SetXmlNodeBool("@defaultSubtotal", false);
                     TopNode.InnerXml = "";
                 }
                 else
                 {
                     string innerXml = "";
                     int count = 0;
                     foreach (eSubTotalFunctions e in Enum.GetValues(typeof(eSubTotalFunctions)))
                    {
                        if ((value & e) == e)
                        {
                            var newTotalType = e.ToString();
                            var totalType = char.ToLower(newTotalType[0], CultureInfo.InvariantCulture) + newTotalType.Substring(1);
                            // add new attribute
                            SetXmlNodeBool("@" + totalType + "Subtotal", true);
                            innerXml += "<item t=\"" + totalType + "\" />";
                            count++;
                        }
                    }
                    TopNode.InnerXml = string.Format("<items count=\"{0}\">{1}</items>", count, innerXml);
                 }
             }
         }
        /// <summary>
        /// Type of axis
        /// </summary>
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
        /// <summary>
        /// If the field is a row field
        /// </summary>
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
                    if (BaseIndex == Index)
                    {
                        TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
                    }
                    else
                    {
                        TopNode.InnerXml = "<items count=\"0\"></items>";
                    }
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
        /// <summary>
        /// If the field is a column field
        /// </summary>
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
                    if (BaseIndex == Index)
                    {
                        TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
                    }
                    else
                    {
                        TopNode.InnerXml = "<items count=\"0\"></items>";
                    }
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
        /// <summary>
        /// If the field is a datafield
        /// </summary>
        public bool IsDataField
        {
            get
            {
                return GetXmlNodeBool("@dataField", false);
            }
        }
        /// <summary>
        /// If the field is a page field.
        /// </summary>
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
                    var dataFieldsNode = TopNode.SelectSingleNode("../../d:pageFields", NameSpaceManager);
                    if (dataFieldsNode == null)
                    {
                        _table.CreateNode("d:pageFields");
                        dataFieldsNode = TopNode.SelectSingleNode("../../d:pageFields", NameSpaceManager);
                    }

                    TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";

                    XmlElement node = AppendField(dataFieldsNode, Index, "pageField", "fld");
                    _pageFieldSettings = new ExcelPivotTablePageFieldSettings(NameSpaceManager, node, this, Index);
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
        internal ExcelPivotTablePageFieldSettings _pageFieldSettings = null;
        public ExcelPivotTablePageFieldSettings PageFieldSettings
        {
            get
            {
                return _pageFieldSettings;
            }
        }
        internal eDateGroupBy DateGrouping
        {
            get;
            set;
        }
        ExcelPivotTableFieldGroup _grouping=null;
        /// <summary>
        /// Grouping settings. 
        /// Null if the field has no grouping otherwise ExcelPivotTableFieldNumericGroup or ExcelPivotTableFieldNumericGroup.
        /// </summary>        
        public ExcelPivotTableFieldGroup Grouping
        {
            get
            {
                return _grouping;
            }
        }
        #region Private & internal Methods
        internal XmlElement AppendField(XmlNode rowsNode, int index, string fieldNodeText, string indexAttrText)
        {
            XmlElement prevField = null, newElement;
            foreach (XmlElement field in rowsNode.ChildNodes)
            {
                string x = field.GetAttribute(indexAttrText);
                int fieldIndex;
                if(int.TryParse(x, out fieldIndex))
                {
                    if (fieldIndex == index)    //Row already exists
                    {
                        return field;
                    }
                    //else if (fieldIndex > index)
                    //{
                    //    newElement = rowsNode.OwnerDocument.CreateElement(fieldNodeText, ExcelPackage.schemaMain);
                    //    newElement.SetAttribute(indexAttrText, index.ToString());
                    //    rowsNode.InsertAfter(newElement, field);
                    //}
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
            var groupNode = cacheField.SelectSingleNode("d:fieldGroup", NameSpaceManager);
            if (groupNode!=null)
            {
                var groupBy = groupNode.SelectSingleNode("d:rangePr/@groupBy", NameSpaceManager);
                if (groupBy==null)
                {
                    _grouping = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, cacheField);
                }
                else
                {
                    DateGrouping=(eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), groupBy.Value, true);
                    _grouping = new ExcelPivotTableFieldDateGroup(NameSpaceManager, groupNode);
                }
            }
        }
        #endregion
        #region Grouping
        internal ExcelPivotTableFieldDateGroup SetDateGroup(eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
        {
            ExcelPivotTableFieldDateGroup group;
            group = new ExcelPivotTableFieldDateGroup(NameSpaceManager, _cacheFieldHelper.TopNode);
            _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsDate", true);
            _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsNonDate", false);
            _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);

            group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"{1}\" /><groupItems /></fieldGroup>", BaseIndex, GroupBy.ToString().ToLower(CultureInfo.InvariantCulture));

            if (StartDate.Year < 1900)
            {
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", "1900-01-01T00:00:00");
            }
            else
            {
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", StartDate.ToString("s", CultureInfo.InvariantCulture));
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoStart", "0");
            }

            if (EndDate==DateTime.MaxValue)
            {
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", "9999-12-31T00:00:00");
            }
            else
            {
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", EndDate.ToString("s", CultureInfo.InvariantCulture));
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@autoEnd", "0");
            }

            int items = AddDateGroupItems(group, GroupBy, StartDate, EndDate, interval);
            AddFieldItems(items);

            _grouping = group;
            return group;
        }
        internal ExcelPivotTableFieldNumericGroup SetNumericGroup(double start, double end, double interval)
        {
            ExcelPivotTableFieldNumericGroup group;
            group = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, _cacheFieldHelper.TopNode);
            _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsNumber", true);
            _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsInteger", true);
            _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);
            _cacheFieldHelper.SetXmlNodeBool("d:sharedItems/@containsString", false);

            group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"{1}\" endNum=\"{2}\" groupInterval=\"{3}\"/><groupItems /></fieldGroup>", BaseIndex, start.ToString(CultureInfo.InvariantCulture), end.ToString(CultureInfo.InvariantCulture), interval.ToString(CultureInfo.InvariantCulture));
            int items = AddNumericGroupItems(group, start, end, interval);
            AddFieldItems(items);

            _grouping = group;
            return group;
        }

        private int AddNumericGroupItems(ExcelPivotTableFieldNumericGroup group, double start, double end, double interval)
        {
            if (interval < 0)
            {
                throw (new Exception("The interval must be a positiv"));
            }
            if (start > end)
            {
                throw(new Exception("Then End number must be larger than the Start number"));
            }

            XmlElement groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
            int items = 2;
            //First date
            double index=start;
            double nextIndex=start+interval;
            AddGroupItem(groupItems, "<" + start.ToString(CultureInfo.InvariantCulture));

            while (index < end)
            {
                AddGroupItem(groupItems, string.Format("{0}-{1}", index.ToString(CultureInfo.InvariantCulture), nextIndex.ToString(CultureInfo.InvariantCulture)));
                index=nextIndex;
                nextIndex+=interval;
                items++;
            }
            AddGroupItem(groupItems, ">" + nextIndex.ToString(CultureInfo.InvariantCulture));
            return items;
        }

        private void AddFieldItems(int items)
        {
            XmlElement prevNode = null;
            XmlElement itemsNode = TopNode.SelectSingleNode("d:items", NameSpaceManager) as XmlElement;
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

        private int AddDateGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
        {
            XmlElement groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
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
                    items += 24;
                    break;
                case eDateGroupBy.Days:
                    if (interval == 1)
                    {
                        DateTime dt = new DateTime(2008, 1, 1); //pick a year with 366 days
                        while (dt.Year == 2008)
                        {
                            AddGroupItem(groupItems, dt.ToString("dd-MMM"));
                            dt = dt.AddDays(1);
                        }
                        items += 366;
                    }
                    else
                    {
                        DateTime dt = StartDate;
                        items = 0;
                        while (dt < EndDate)
                        {
                            AddGroupItem(groupItems, dt.ToString("dd-MMM"));
                            dt = dt.AddDays(interval);
                            items++;
                        }
                    }
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
                    if(StartDate.Year>=1900 && EndDate!=DateTime.MaxValue)
                    {
                        for (int year = StartDate.Year; year <= EndDate.Year; year++)
                        {
                            AddGroupItem(groupItems, year.ToString());
                        }
                        items += EndDate.Year - StartDate.Year+1;
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
                AddGroupItem(groupItems, string.Format("{0:00}", i));
            }
        }

        private void AddGroupItem(XmlElement groupItems, string value)
        {
            var s = groupItems.OwnerDocument.CreateElement("s", ExcelPackage.schemaMain);
            s.SetAttribute("v", value);
            groupItems.AppendChild(s);
        }
        #endregion
        internal ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem> _items=null;
        /// <summary>
        /// Pivottable field Items. Used for grouping.
        /// </summary>
        public ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem> Items
        {
            get
            {
                if (_items == null)
                {
                    _items = new ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem>(_table);
                    foreach (XmlNode node in TopNode.SelectNodes("d:items//d:item", NameSpaceManager))
                    {
                        var item = new ExcelPivotTableFieldItem(NameSpaceManager, node,this);
                        if (item.T == "")
                        {
                            _items.AddInternal(item);
                        }
                    }
                    //if (_grouping is ExcelPivotTableFieldDateGroup)
                    //{
                    //    ExcelPivotTableFieldDateGroup dtgrp = ((ExcelPivotTableFieldDateGroup)_grouping);

                    //    ExcelPivotTableFieldItem minItem=null, maxItem=null;
                    //    foreach (var item in _items)
                    //    {
                    //        if (item.X == 0)
                    //        {
                    //            minItem = item;
                    //        }
                    //        else if (maxItem == null || maxItem.X < item.X)
                    //        {
                    //            maxItem = item;
                    //        }
                    //    }
                    //    if (dtgrp.AutoStart)
                    //    {
                    //        _items._list.Remove(minItem);
                    //    }
                    //    if (dtgrp.AutoEnd)
                    //    {
                    //        _items._list.Remove(maxItem);
                    //    }

                    //}
                }
                return _items;
            }
        }
        /// <summary>
        /// Add numberic grouping to the field
        /// </summary>
        /// <param name="Start">Start value</param>
        /// <param name="End">End value</param>
        /// <param name="Interval">Interval</param>
        public void AddNumericGrouping(double Start, double End, double Interval)
        {
            ValidateGrouping();
            SetNumericGroup(Start, End, Interval);
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="groupBy">Group by</param>
        public void AddDateGrouping(eDateGroupBy groupBy)
        {
            AddDateGrouping(groupBy, DateTime.MinValue, DateTime.MaxValue,1);
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="groupBy">Group by</param>
        /// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
        /// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
        public void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate)
        {
            AddDateGrouping(groupBy, startDate, endDate, 1);
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="days">Number of days when grouping on days</param>
        /// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
        /// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
        public void AddDateGrouping(int days, DateTime startDate, DateTime endDate)
        {
            AddDateGrouping(eDateGroupBy.Days, startDate, endDate, days);
        }
        private void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int groupInterval)
        {
            if (groupInterval < 1 || groupInterval >= Int16.MaxValue)
            {
                throw (new ArgumentOutOfRangeException("Group interval is out of range"));
            }
            if (groupInterval > 1 && groupBy != eDateGroupBy.Days)
            {
                throw (new ArgumentException("Group interval is can only be used when groupBy is Days"));
            }
            ValidateGrouping();

            bool firstField = true;            
            List<ExcelPivotTableField> fields=new List<ExcelPivotTableField>();
            //Seconds
            if ((groupBy & eDateGroupBy.Seconds) == eDateGroupBy.Seconds)
            {
                fields.Add(AddField(eDateGroupBy.Seconds, startDate, endDate, ref firstField));
            }
            //Minutes
            if ((groupBy & eDateGroupBy.Minutes) == eDateGroupBy.Minutes)
            {
                fields.Add(AddField(eDateGroupBy.Minutes, startDate, endDate, ref firstField));
            }
            //Hours
            if ((groupBy & eDateGroupBy.Hours) == eDateGroupBy.Hours)
            {
                fields.Add(AddField(eDateGroupBy.Hours, startDate, endDate, ref firstField));
            }
            //Days
            if ((groupBy & eDateGroupBy.Days) == eDateGroupBy.Days)
            {
                fields.Add(AddField(eDateGroupBy.Days, startDate, endDate, ref firstField, groupInterval));
            }
            //Month
            if ((groupBy & eDateGroupBy.Months) == eDateGroupBy.Months)
            {
                fields.Add(AddField(eDateGroupBy.Months, startDate, endDate, ref firstField));
            }
            //Quarters
            if ((groupBy & eDateGroupBy.Quarters) == eDateGroupBy.Quarters)
            {
                fields.Add(AddField(eDateGroupBy.Quarters, startDate, endDate, ref firstField));
            }
            //Years
            if ((groupBy & eDateGroupBy.Years) == eDateGroupBy.Years)
            {
                fields.Add(AddField(eDateGroupBy.Years, startDate, endDate, ref firstField));
            }

            if (fields.Count > 1) _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/@par", (_table.Fields.Count - 1).ToString());
            if (groupInterval != 1)
            {
                _cacheFieldHelper.SetXmlNodeString("d:fieldGroup/d:rangePr/@groupInterval", groupInterval.ToString());
            }
            else
            {
                _cacheFieldHelper.DeleteNode("d:fieldGroup/d:rangePr/@groupInterval");
            }
            _items = null;
        }

        private void ValidateGrouping()
        {
            if (!(IsColumnField || IsRowField))
            {
                throw (new Exception("Field must be a row or column field"));
            }
            foreach (var field in _table.Fields)
            {
                if (field.Grouping != null)
                {
                    throw (new Exception("Grouping already exists"));
                }
            }
        }
        private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref  bool firstField)
        {
            return AddField(groupBy, startDate, endDate, ref firstField,1);
        }
        private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref  bool firstField, int interval)
        {
            if (firstField == false)
            {
                //Pivot field
                var topNode = _table.PivotTableXml.SelectSingleNode("//d:pivotFields", _table.NameSpaceManager);
                var fieldNode = _table.PivotTableXml.CreateElement("pivotField", ExcelPackage.schemaMain);
                fieldNode.SetAttribute("compact", "0");
                fieldNode.SetAttribute("outline", "0");
                fieldNode.SetAttribute("showAll", "0");
                fieldNode.SetAttribute("defaultSubtotal", "0");
                topNode.AppendChild(fieldNode);

                var field = new ExcelPivotTableField(_table.NameSpaceManager, fieldNode, _table, _table.Fields.Count, Index);
                field.DateGrouping = groupBy;

                XmlNode rowColFields;
                if (IsRowField)
                {
                    rowColFields=TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);
                }
                else
                {
                    rowColFields = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);
                }

                int fieldIndex, index = 0;
                foreach (XmlElement rowfield in rowColFields.ChildNodes)
                {
                    if (int.TryParse(rowfield.GetAttribute("x"), out fieldIndex))
                    {
                        if (_table.Fields[fieldIndex].BaseIndex == BaseIndex)
                        {
                            var newElement = rowColFields.OwnerDocument.CreateElement("field", ExcelPackage.schemaMain);
                            newElement.SetAttribute("x", field.Index.ToString());
                            rowColFields.InsertBefore(newElement, rowfield);
                            break;
                        }
                    }
                    index++;
                }

                if (IsRowField)
                {
                    _table.RowFields.Insert(field, index);
                }
                else
                {
                    _table.ColumnFields.Insert(field, index);
                }
                
                _table.Fields.AddInternal(field);

                AddCacheField(field, startDate, endDate, interval);
                return field;
            }
            else
            {
                firstField = false;
                DateGrouping = groupBy;
                Compact = false;
                SetDateGroup(groupBy, startDate, endDate, interval);
                return this;
            }
        }
        private void AddCacheField(ExcelPivotTableField field, DateTime startDate, DateTime endDate, int interval)
        {
            //Add Cache definition field.
            var cacheTopNode = _table.CacheDefinition.CacheDefinitionXml.SelectSingleNode("//d:cacheFields", _table.NameSpaceManager);
            var cacheFieldNode = _table.CacheDefinition.CacheDefinitionXml.CreateElement("cacheField", ExcelPackage.schemaMain);

            cacheFieldNode.SetAttribute("name", field.DateGrouping.ToString());
            cacheFieldNode.SetAttribute("databaseField", "0");
            cacheTopNode.AppendChild(cacheFieldNode);
            field.SetCacheFieldNode(cacheFieldNode);

            field.SetDateGroup(field.DateGrouping, startDate, endDate, interval);
        }
    }
}
