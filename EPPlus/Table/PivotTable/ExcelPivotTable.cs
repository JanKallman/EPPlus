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
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// An Excel Pivottable
    /// </summary>
    public class ExcelPivotTable : XmlHelper
    {
        internal ExcelPivotTable(Packaging.ZipPackageRelationship rel, ExcelWorksheet sheet) : 
            base(sheet.NameSpaceManager)
        {
            WorkSheet = sheet;
            PivotTableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            Relationship = rel;
            var pck = sheet._package.Package;
            Part=pck.GetPart(PivotTableUri);

            PivotTableXml = new XmlDocument();
            LoadXmlSafe(PivotTableXml, Part.GetStream());
            init();
            TopNode = PivotTableXml.DocumentElement;
            Address = new ExcelAddressBase(GetXmlNodeString("d:location/@ref"));

            _cacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this);
            LoadFields();

            //Add row fields.
            foreach (XmlElement rowElem in TopNode.SelectNodes("d:rowFields/d:field", NameSpaceManager))
            {
                int x;
                if (int.TryParse(rowElem.GetAttribute("x"), out x) && x >= 0)
                {
                    RowFields.AddInternal(Fields[x]);
                }
                else
                {
                    rowElem.ParentNode.RemoveChild(rowElem);
                }
            }

            ////Add column fields.
            foreach (XmlElement colElem in TopNode.SelectNodes("d:colFields/d:field", NameSpaceManager))
            {
                int x;
                if(int.TryParse(colElem.GetAttribute("x"),out x) && x >= 0)
                {
                    ColumnFields.AddInternal(Fields[x]);
                }
                else
                {
                    colElem.ParentNode.RemoveChild(colElem);
                }
            }

            //Add Page elements
            //int index = 0;
            foreach (XmlElement pageElem in TopNode.SelectNodes("d:pageFields/d:pageField", NameSpaceManager))
            {
                int fld;
                if (int.TryParse(pageElem.GetAttribute("fld"), out fld) && fld >= 0)
                {
                    var field = Fields[fld];
                    field._pageFieldSettings = new ExcelPivotTablePageFieldSettings(NameSpaceManager, pageElem, field, fld);
                    PageFields.AddInternal(field);
                }
            }

            //Add data elements
            //index = 0;
            foreach (XmlElement dataElem in TopNode.SelectNodes("d:dataFields/d:dataField", NameSpaceManager))
            {
                int fld;
                if (int.TryParse(dataElem.GetAttribute("fld"), out fld) && fld >= 0)
                {
                    var field = Fields[fld];
                    var dataField = new ExcelPivotTableDataField(NameSpaceManager, dataElem, field);
                    DataFields.AddInternal(dataField);
                }
            }
        }
        /// <summary>
        /// Add a new pivottable
        /// </summary>
        /// <param name="sheet">The worksheet</param>
        /// <param name="address">the address of the pivottable</param>
        /// <param name="sourceAddress">The address of the Source data</param>
        /// <param name="name"></param>
        /// <param name="tblId"></param>
        internal ExcelPivotTable(ExcelWorksheet sheet, ExcelAddressBase address,ExcelRangeBase sourceAddress, string name, int tblId) : 
            base(sheet.NameSpaceManager)
	    {
            WorkSheet = sheet;
            Address = address;
            var pck = sheet._package.Package;

            PivotTableXml = new XmlDocument();
            LoadXmlSafe(PivotTableXml, GetStartXml(name, tblId, address, sourceAddress), Encoding.UTF8);
            TopNode = PivotTableXml.DocumentElement;
            PivotTableUri =  GetNewUri(pck, "/xl/pivotTables/pivotTable{0}.xml", tblId);
            init();

            Part = pck.CreatePart(PivotTableUri, ExcelPackage.schemaPivotTable);
            PivotTableXml.Save(Part.GetStream());
            
            //Worksheet-Pivottable relationship
            Relationship = sheet.Part.CreateRelationship(UriHelper.ResolvePartUri(sheet.WorksheetUri, PivotTableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");

            _cacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, sourceAddress, tblId);
            _cacheDefinition.Relationship=Part.CreateRelationship(UriHelper.ResolvePartUri(PivotTableUri, _cacheDefinition.CacheDefinitionUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");

            sheet.Workbook.AddPivotTable(CacheID.ToString(), _cacheDefinition.CacheDefinitionUri);

            LoadFields();

            using (var r=sheet.Cells[address.Address])
            {
                r.Clear();
            }
        }
        private void init()
        {
            SchemaNodeOrder = new string[] { "location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "pageItems", "dataFields", "dataItems", "formats", "pivotTableStyleInfo" };
        }
        private void LoadFields()
        {
            //Fields.Clear();
            //int ix=0;
            //foreach(XmlElement fieldNode in PivotXml.SelectNodes("//d:pivotFields/d:pivotField",NameSpaceManager))
            //{
            //    Fields.AddInternal(new ExcelPivotTableField(NameSpaceManager, fieldNode, this, ix++));
            //}

            int index = 0;
            //Add fields.
            foreach (XmlElement fieldElem in TopNode.SelectNodes("d:pivotFields/d:pivotField", NameSpaceManager))
            {
                var fld = new ExcelPivotTableField(NameSpaceManager, fieldElem, this, index, index++);
                Fields.AddInternal(fld);
            }

            //Add fields.
            index = 0;
            foreach (XmlElement fieldElem in _cacheDefinition.TopNode.SelectNodes("d:cacheFields/d:cacheField", NameSpaceManager))
            {
                var fld = Fields[index++];
                fld.SetCacheFieldNode(fieldElem);
            }


        }
        private string GetStartXml(string name, int id, ExcelAddressBase address, ExcelAddressBase sourceAddress)
        {
            string xml = string.Format("<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"{0}\" cacheId=\"{1}\" dataOnRows=\"1\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" dataCaption=\"Data\"  createdVersion=\"4\" showMemberPropertyTips=\"0\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" indent=\"0\" compact=\"0\" compactData=\"0\" gridDropZones=\"1\">", name, id);

            xml += string.Format("<location ref=\"{0}\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" /> ", address.FirstAddress);
            xml += string.Format("<pivotFields count=\"{0}\">", sourceAddress._toCol-sourceAddress._fromCol+1);
            for (int col = sourceAddress._fromCol; col <= sourceAddress._toCol; col++)
            {
                xml += "<pivotField showAll=\"0\" />"; //compact=\"0\" outline=\"0\" subtotalTop=\"0\" includeNewItemsInFilter=\"1\"     
            }

            xml += "</pivotFields>";
            xml += "<pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\" />";
            xml += "</pivotTableDefinition>";
            return xml;
        }
        internal Packaging.ZipPackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Provides access to the XML data representing the pivottable in the package.
        /// </summary>
        public XmlDocument PivotTableXml { get; private set; }
        /// <summary>
        /// The package internal URI to the pivottable Xml Document.
        /// </summary>
        public Uri PivotTableUri
        {
            get;
            internal set;
        }
        internal Packaging.ZipPackageRelationship Relationship
        {
            get;
            set;
        }
        const string ID_PATH = "@id";
        internal int Id
        {
            get
            {
                return GetXmlNodeInt(ID_PATH);
            }
            set
            {
                SetXmlNodeString(ID_PATH, value.ToString());
            }
        }
        const string NAME_PATH = "@name";
        const string DISPLAY_NAME_PATH = "@displayName";
        /// <summary>
        /// Name of the pivottable object in Excel
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString(NAME_PATH);
            }
            set
            {
                if (WorkSheet.Workbook.ExistsTableName(value))
                {
                    throw (new ArgumentException("PivotTable name is not unique"));
                }
                string prevName = Name;
                if (WorkSheet.Tables._tableNames.ContainsKey(prevName))
                {
                    int ix = WorkSheet.Tables._tableNames[prevName];
                    WorkSheet.Tables._tableNames.Remove(prevName);
                    WorkSheet.Tables._tableNames.Add(value, ix);
                }
                SetXmlNodeString(NAME_PATH, value);
                SetXmlNodeString(DISPLAY_NAME_PATH, cleanDisplayName(value));
            }
        }        
        ExcelPivotCacheDefinition _cacheDefinition = null;
        /// <summary>
        /// Reference to the pivot table cache definition object
        /// </summary>
        public ExcelPivotCacheDefinition CacheDefinition
        {
            get
            {
                if (_cacheDefinition == null)
                {
                    _cacheDefinition = new ExcelPivotCacheDefinition(NameSpaceManager, this, null, 1);
                }
                return _cacheDefinition;
            }
        }
        private string cleanDisplayName(string name)
        {
            return Regex.Replace(name, @"[^\w\.-_]", "_");
        }
        #region "Public Properties"

        /// <summary>
        /// The worksheet where the pivottable is located
        /// </summary>
        public ExcelWorksheet WorkSheet
        {
            get;
            set;
        }
        /// <summary>
        /// The location of the pivot table
        /// </summary>
        public ExcelAddressBase Address
        {
            get;
            internal set;
        }
        /// <summary>
        /// If multiple datafields are displayed in the row area or the column area
        /// </summary>
        public bool DataOnRows 
        { 
            get
            {
                return GetXmlNodeBool("@dataOnRows");
            }
            set
            {
                SetXmlNodeBool("@dataOnRows",value);
            }
        }
        /// <summary>
        /// if true apply legacy table autoformat number format properties.
        /// </summary>
        public bool ApplyNumberFormats 
        { 
            get
            {
                return GetXmlNodeBool("@applyNumberFormats");
            }
            set
            {
                SetXmlNodeBool("@applyNumberFormats",value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat border properties
        /// </summary>
        public bool ApplyBorderFormats 
        { 
            get
            {
                return GetXmlNodeBool("@applyBorderFormats");
            }
            set
            {
                SetXmlNodeBool("@applyBorderFormats",value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat font properties
        /// </summary>
        public bool ApplyFontFormats
        { 
            get
            {
                return GetXmlNodeBool("@applyFontFormats");
            }
            set
            {
                SetXmlNodeBool("@applyFontFormats",value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat pattern properties
        /// </summary>
        public bool ApplyPatternFormats
        { 
            get
            {
                return GetXmlNodeBool("@applyPatternFormats");
            }
            set
            {
                SetXmlNodeBool("@applyPatternFormats",value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat width/height properties.
        /// </summary>
        public bool ApplyWidthHeightFormats
        { 
            get
            {
                return GetXmlNodeBool("@applyWidthHeightFormats");
            }
            set
            {
                SetXmlNodeBool("@applyWidthHeightFormats",value);
            }
        }
        /// <summary>
        /// Show member property information
        /// </summary>
        public bool ShowMemberPropertyTips
        { 
            get
            {
                return GetXmlNodeBool("@showMemberPropertyTips");
            }
            set
            {
                SetXmlNodeBool("@showMemberPropertyTips",value);
            }
        }
        /// <summary>
        /// Show the drill indicators
        /// </summary>
        public bool ShowCalcMember
        {
            get
            {
                return GetXmlNodeBool("@showCalcMbrs");
            }
            set
            {
                SetXmlNodeBool("@showCalcMbrs", value);
            }
        }       
        /// <summary>
        /// If the user is prevented from drilling down on a PivotItem or aggregate value
        /// </summary>
        public bool EnableDrill
        {
            get
            {
                return GetXmlNodeBool("@enableDrill", true);
            }
            set
            {
                SetXmlNodeBool("@enableDrill", value);
            }
        }
        /// <summary>
        /// Show the drill down buttons
        /// </summary>
        public bool ShowDrill
        {
            get
            {
                return GetXmlNodeBool("@showDrill", true);
            }
            set
            {
                SetXmlNodeBool("@showDrill", value);
            }
        }
        /// <summary>
        /// If the tooltips should be displayed for PivotTable data cells.
        /// </summary>
        public bool ShowDataTips
        {
            get
            {
                return GetXmlNodeBool("@showDataTips", true);
            }
            set
            {
                SetXmlNodeBool("@showDataTips", value, true);
            }
        }
        /// <summary>
        /// If the row and column titles from the PivotTable should be printed.
        /// </summary>
        public bool FieldPrintTitles
        {
            get
            {
                return GetXmlNodeBool("@fieldPrintTitles");
            }
            set
            {
                SetXmlNodeBool("@fieldPrintTitles", value);
            }
        }
        /// <summary>
        /// If the row and column titles from the PivotTable should be printed.
        /// </summary>
        public bool ItemPrintTitles
        {
            get
            {
                return GetXmlNodeBool("@itemPrintTitles");
            }
            set
            {
                SetXmlNodeBool("@itemPrintTitles", value);
            }
        }
        /// <summary>
        /// If the grand totals should be displayed for the PivotTable columns
        /// </summary>
        public bool ColumGrandTotals
        {
            get
            {
                return GetXmlNodeBool("@colGrandTotals");
            }
            set
            {
                SetXmlNodeBool("@colGrandTotals", value);
            }
        }        
        /// <summary>
        /// If the grand totals should be displayed for the PivotTable rows
        /// </summary>
        public bool RowGrandTotals
        {
            get
            {
                return GetXmlNodeBool("@rowGrandTotals");
            }
            set
            {
                SetXmlNodeBool("@rowGrandTotals", value);
            }
        }        
        /// <summary>
        /// If the drill indicators expand collapse buttons should be printed.
        /// </summary>
        public bool PrintDrill
        {
            get
            {
                return GetXmlNodeBool("@printDrill");
            }
            set
            {
                SetXmlNodeBool("@printDrill", value);
            }
        }        
        /// <summary>
        /// Indicates whether to show error messages in cells.
        /// </summary>
        public bool ShowError
        {
            get
            {
                return GetXmlNodeBool("@showError");
            }
            set
            {
                SetXmlNodeBool("@showError", value);
            }
        }        
        /// <summary>
        /// The string to be displayed in cells that contain errors.
        /// </summary>
        public string ErrorCaption
        {
            get
            {
                return GetXmlNodeString("@errorCaption");
            }
            set
            {
                SetXmlNodeString("@errorCaption", value);
            }
        }        
        /// <summary>
        /// Specifies the name of the value area field header in the PivotTable. 
        /// This caption is shown when the PivotTable when two or more fields are in the values area.
        /// </summary>
        public string DataCaption
        {
            get
            {
                return GetXmlNodeString("@dataCaption");
            }
            set
            {
                SetXmlNodeString("@dataCaption", value);
            }
        }        
        /// <summary>
        /// Show field headers
        /// </summary>
        public bool ShowHeaders
        {
            get
            {
                return GetXmlNodeBool("@showHeaders");
            }
            set
            {
                SetXmlNodeBool("@showHeaders", value);
            }
        }
        /// <summary>
        /// The number of page fields to display before starting another row or column
        /// </summary>
        public int PageWrap
        {
            get
            {
                return GetXmlNodeInt("@pageWrap");
            }
            set
            {
                if(value<0)
                {
                    throw new Exception("Value can't be negative");
                }
                SetXmlNodeString("@pageWrap", value.ToString());
            }
        }
        /// <summary>
        /// A boolean that indicates whether legacy auto formatting has been applied to the PivotTable view
        /// </summary>
        public bool UseAutoFormatting
        { 
            get
            {
                return GetXmlNodeBool("@useAutoFormatting");
            }
            set
            {
                SetXmlNodeBool("@useAutoFormatting",value);
            }
        } 
        /// <summary>
        /// A boolean that indicates whether the in-grid drop zones should be displayed at runtime, and whether classic layout is applied
        /// </summary>
        public bool GridDropZones
        { 
            get
            {
                return GetXmlNodeBool("@gridDropZones");
            }
            set
            {
                SetXmlNodeBool("@gridDropZones",value);
            }
        }
        /// <summary>
        /// Specifies the indentation increment for compact axis and can be used to set the Report Layout to Compact Form
        /// </summary>
        public int Indent
        { 
            get
            {
                return GetXmlNodeInt("@indent");
            }
            set
            {
                SetXmlNodeString("@indent",value.ToString());
            }
        }
        /// <summary>
        /// A boolean that indicates whether data fields in the PivotTable should be displayed in outline form
        /// </summary>
        public bool OutlineData
        {
            get
            {
                return GetXmlNodeBool("@outlineData");
            }
            set
            {
                SetXmlNodeBool("@outlineData", value);
            }
        }
        /// <summary>
        /// a boolean that indicates whether new fields should have their outline flag set to true
        /// </summary>
        public bool Outline
        {
            get
            {
                return GetXmlNodeBool("@outline");
            }
            set
            {
                SetXmlNodeBool("@outline", value);
            }
        }
        /// <summary>
        /// A boolean that indicates whether the fields of a PivotTable can have multiple filters set on them
        /// </summary>
        public bool MultipleFieldFilters
        {
            get
            {
                return GetXmlNodeBool("@multipleFieldFilters");
            }
            set
            {
                SetXmlNodeBool("@multipleFieldFilters", value);
            }
        }
        /// <summary>
        /// A boolean that indicates whether new fields should have their compact flag set to true
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
        /// A boolean that indicates whether the field next to the data field in the PivotTable should be displayed in the same column of the spreadsheet
        /// </summary>
        public bool CompactData
        { 
            get
            {
                return GetXmlNodeBool("@compactData");
            }
            set
            {
                SetXmlNodeBool("@compactData",value);
            }
        }
        /// <summary>
        /// Specifies the string to be displayed for grand totals.
        /// </summary>
        public string GrandTotalCaption
        {
            get
            {
                return GetXmlNodeString("@grandTotalCaption");
            }
            set
            {
                SetXmlNodeString("@grandTotalCaption", value);
            }
        }
        /// <summary>
        /// Specifies the string to be displayed in row header in compact mode.
        /// </summary>
        public string RowHeaderCaption 
        {
            get
            {
                return GetXmlNodeString("@rowHeaderCaption");
            }
            set
            {
                SetXmlNodeString("@rowHeaderCaption", value);                
            }
        }
        /// <summary>
        /// Specifies the string to be displayed in cells with no value
        /// </summary>
        public string MissingCaption
        {
            get
            {
                return GetXmlNodeString("@missingCaption");
            }
            set
            {
                SetXmlNodeString("@missingCaption", value);                
            }
        }
        const string FIRSTHEADERROW_PATH="d:location/@firstHeaderRow";
        /// <summary>
        /// Specifies the first row of the PivotTable header, relative to the top left cell in the ref value
        /// </summary>
        public int FirstHeaderRow
        {
            get
            {
                return GetXmlNodeInt(FIRSTHEADERROW_PATH);
            }
            set
            {
                SetXmlNodeString(FIRSTHEADERROW_PATH, value.ToString());
            }
        }
        const string FIRSTDATAROW_PATH = "d:location/@firstDataRow";
        /// <summary>
        /// Specifies the first column of the PivotTable data, relative to the top left cell in the ref value
        /// </summary>
        public int FirstDataRow
        {
            get
            {
                return GetXmlNodeInt(FIRSTDATAROW_PATH);
            }
            set
            {
                SetXmlNodeString(FIRSTDATAROW_PATH, value.ToString());
            }
        }
        const string FIRSTDATACOL_PATH = "d:location/@firstDataCol";
        /// <summary>
        /// Specifies the first column of the PivotTable data, relative to the top left cell in the ref value
        /// </summary>
        public int FirstDataCol
        {
            get
            {
                return GetXmlNodeInt(FIRSTDATACOL_PATH);
            }
            set
            {
                SetXmlNodeString(FIRSTDATACOL_PATH, value.ToString());
            }
        }
        ExcelPivotTableFieldCollection _fields = null;
        /// <summary>
        /// The fields in the table 
        /// </summary>
        public ExcelPivotTableFieldCollection Fields
        {
            get
            {
                if (_fields == null)
                {
                    _fields = new ExcelPivotTableFieldCollection(this, "");
                }
                return _fields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _rowFields = null;
        /// <summary>
        /// Row label fields 
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection RowFields
        {
            get
            {
                if (_rowFields == null)
                {
                    _rowFields = new ExcelPivotTableRowColumnFieldCollection(this, "rowFields");
                }
                return _rowFields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _columnFields = null;
        /// <summary>
        /// Column label fields 
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection ColumnFields
        {
            get
            {
                if (_columnFields == null)
                {
                    _columnFields = new ExcelPivotTableRowColumnFieldCollection(this, "colFields");
                }
                return _columnFields;
            }
        }
        ExcelPivotTableDataFieldCollection _dataFields = null;
        /// <summary>
        /// Value fields 
        /// </summary>
        public ExcelPivotTableDataFieldCollection DataFields
        {
            get
            {
                if (_dataFields == null)
                {
                    _dataFields = new ExcelPivotTableDataFieldCollection(this);
                }
                return _dataFields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _pageFields = null;
        /// <summary>
        /// Report filter fields
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection PageFields
        {
            get
            {
                if (_pageFields == null)
                {
                    _pageFields = new ExcelPivotTableRowColumnFieldCollection(this, "pageFields");
                }
                return _pageFields;
            }
        }
        const string STYLENAME_PATH = "d:pivotTableStyleInfo/@name";
        /// <summary>
        /// Pivot style name. Used for custom styles
        /// </summary>
        public string StyleName
        {
            get
            {
                return GetXmlNodeString(StyleName);
            }
            set
            {
                if (value.StartsWith("PivotStyle"))
                {
                    try
                    {
                        _tableStyle = (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10, value.Length - 10), true);
                    }
                    catch
                    {
                        _tableStyle = TableStyles.Custom;
                    }
                }
                else if (value == "None")
                {
                    _tableStyle = TableStyles.None;
                    value = "";
                }
                else
                {
                    _tableStyle = TableStyles.Custom;
                }
                SetXmlNodeString(STYLENAME_PATH, value, true);
            }
        }
        TableStyles _tableStyle = Table.TableStyles.Medium6;
        /// <summary>
        /// The table style. If this property is cusom the style from the StyleName propery is used.
        /// </summary>
        public TableStyles TableStyle
        {
            get
            {
                return _tableStyle;
            }
            set
            {
                _tableStyle=value;
                if (value != TableStyles.Custom)
                {
                    SetXmlNodeString(STYLENAME_PATH, "PivotStyle" + value.ToString());
                }
            }
        }

        #endregion
        #region "Internal Properties"
        internal int CacheID 
        { 
                get
                {
                    return GetXmlNodeInt("@cacheId");
                }
                set
                {
                    SetXmlNodeString("@cacheId",value.ToString());
                }
        }

        #endregion

    }
}
