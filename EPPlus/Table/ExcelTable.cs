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
 * Jan Källman		Added		30-AUG-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Table
{
    /// <summary>
    /// Table style Enum
    /// </summary>
    public enum TableStyles
    {
        None,
        Custom,
        Light1,
        Light2,
        Light3,
        Light4,
        Light5,
        Light6,
        Light7,
        Light8,
        Light9,
        Light10,
        Light11,
        Light12,
        Light13,
        Light14,
        Light15,
        Light16,
        Light17,
        Light18,
        Light19,
        Light20,
        Light21,
        Medium1,
        Medium2,
        Medium3,
        Medium4,
        Medium5,
        Medium6,
        Medium7,
        Medium8,
        Medium9,
        Medium10,
        Medium11,
        Medium12,
        Medium13,
        Medium14,
        Medium15,
        Medium16,
        Medium17,
        Medium18,
        Medium19,
        Medium20,
        Medium21,
        Medium22,
        Medium23,
        Medium24,
        Medium25,
        Medium26,
        Medium27,
        Medium28,
        Dark1,
        Dark2,
        Dark3,
        Dark4,
        Dark5,
        Dark6,
        Dark7,
        Dark8,
        Dark9,
        Dark10,
        Dark11,
    }
    /// <summary>
    /// An Excel Table
    /// </summary>
    public class ExcelTable : XmlHelper
    {
        internal ExcelTable(PackageRelationship rel, ExcelWorksheet sheet) : 
            base(sheet.NameSpaceManager)
        {
            WorkSheet = sheet;
            TableUri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            RelationshipID = rel.Id;
            var pck = sheet._package.Package;
            Part=pck.GetPart(TableUri);

            TableXml = new XmlDocument();
            LoadXmlSafe(TableXml, Part.GetStream());
            init();
            Address = new ExcelAddressBase(GetXmlNodeString("@ref"));
        }
        internal ExcelTable(ExcelWorksheet sheet, ExcelAddressBase address, string name, int tblId) : 
            base(sheet.NameSpaceManager)
	    {
            WorkSheet = sheet;
            Address = address;
            TableXml = new XmlDocument();
            LoadXmlSafe(TableXml, GetStartXml(name, tblId), Encoding.UTF8); 
            TopNode = TableXml.DocumentElement;

            init();

            //If the table is just one row we can not have a header.
            if (address._fromRow == address._toRow)
            {
                ShowHeader = false;
            }
        }

        private void init()
        {
            TopNode = TableXml.DocumentElement;
            SchemaNodeOrder = new string[] { "autoFilter", "tableColumns", "tableStyleInfo" };
        }
        private string GetStartXml(string name, int tblId)
        {
            string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?>";
            xml += string.Format("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{0}\" name=\"{1}\" displayName=\"{2}\" ref=\"{3}\" headerRowCount=\"1\">",
            tblId,
            name,
            cleanDisplayName(name),
            Address.Address);
            xml += string.Format("<autoFilter ref=\"{0}\" />", Address.Address);

            int cols=Address._toCol-Address._fromCol+1;
            xml += string.Format("<tableColumns count=\"{0}\">",cols);
            var names = new Dictionary<string, string>();            
            for(int i=1;i<=cols;i++)
            {
                var cell = WorkSheet.Cells[Address._fromRow, Address._fromCol+i-1];
                string colName;
                if (cell.Value == null || names.ContainsKey(cell.Value.ToString()))
                {
                    //Get an unique name
                    int a=i;
                    do
                    {
                        colName = string.Format("Column{0}", a++);
                    }
                    while (names.ContainsKey(colName));
                }
                else
                {
                    colName = System.Security.SecurityElement.Escape(cell.Value.ToString());
                }
                names.Add(colName, colName);
                xml += string.Format("<tableColumn id=\"{0}\" name=\"{1}\" />", i,colName);
            }
            xml += "</tableColumns>";
            xml += "<tableStyleInfo name=\"TableStyleMedium9\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\" /> ";
            xml += "</table>";

            return xml;
        }
        private string cleanDisplayName(string name) 
        {
            return Regex.Replace(name, @"[^\w\.-_]", "_");
        }
        internal PackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Provides access to the XML data representing the table in the package.
        /// </summary>
        public XmlDocument TableXml
        {
            get;
            set;
        }
        /// <summary>
        /// The package internal URI to the Table Xml Document.
        /// </summary>
        public Uri TableUri
        {
            get;
            internal set;
        }
        internal string RelationshipID
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
        /// The name of the table object in Excel
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString(NAME_PATH);
            }
            set 
            {
                if(WorkSheet.Workbook.ExistsTableName(value))
                {
                    throw (new ArgumentException("Tablename is not unique"));
                }
                string prevName = Name;
                if (WorkSheet.Tables._tableNames.ContainsKey(prevName))
                {
                    int ix=WorkSheet.Tables._tableNames[prevName];
                    WorkSheet.Tables._tableNames.Remove(prevName);
                    WorkSheet.Tables._tableNames.Add(value,ix);
                }
                SetXmlNodeString(NAME_PATH, value);
                SetXmlNodeString(DISPLAY_NAME_PATH, cleanDisplayName(value));
            }
        }
        /// <summary>
        /// The worksheet of the table
        /// </summary>
        public ExcelWorksheet WorkSheet
        {
            get;
            set;
        }
        /// <summary>
        /// The address of the table
        /// </summary>
        public ExcelAddressBase Address
        {
            get;
            internal set;
        }
        ExcelTableColumnCollection _cols = null;
        /// <summary>
        /// Collection of the columns in the table
        /// </summary>
        public ExcelTableColumnCollection Columns
        {
            get
            {
                if(_cols==null)
                {
                    _cols = new ExcelTableColumnCollection(this);
                }
                return _cols;
            }
        }
        TableStyles _tableStyle = TableStyles.Medium6;
        /// <summary>
        /// The table style. If this property is cusom, the style from the StyleName propery is used.
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
                    SetXmlNodeString(STYLENAME_PATH, "TableStyle" + value.ToString());
                }
            }
        }
        const string HEADERROWCOUNT_PATH = "@headerRowCount";
        const string AUTOFILTER_PATH = "d:autoFilter/@ref";
        /// <summary>
        /// If the header row is visible or not
        /// </summary>
        public bool ShowHeader
        {
            get
            {
                return GetXmlNodeInt(HEADERROWCOUNT_PATH)!=0;
            }
            set
            {
                if (Address._toRow - Address._fromRow < 1 && value ||
                    Address._toRow - Address._fromRow == 1 && value && ShowTotal)
                {
                    throw (new Exception("Cant set ShowHeader-property. Table has too few rows"));
                }

                if(value)
                {
                    DeleteNode(HEADERROWCOUNT_PATH);
                    WriteAutoFilter(ShowTotal);
                }
                else
                {
                    SetXmlNodeString(HEADERROWCOUNT_PATH, "0");
                    DeleteAllNode(AUTOFILTER_PATH);
                }
            }
        }
        internal ExcelAddressBase AutoFilterAddress
        {
            get
            {
                string a=GetXmlNodeString(AUTOFILTER_PATH);
                if (a == "")
                {
                    return null;
                }
                else
                {
                    return new ExcelAddressBase(a);
                }
            }
        }
        private void WriteAutoFilter(bool showTotal)
        {
            string autofilterAddress;
            if (ShowHeader)
            {
                if (showTotal)
                {
                    autofilterAddress = ExcelCellBase.GetAddress(Address._fromRow, Address._fromCol, Address._toRow - 1, Address._toCol);
                }
                else
                {
                    autofilterAddress = Address.Address;
                }
                SetXmlNodeString(AUTOFILTER_PATH, autofilterAddress);
            }
        }
        /// <summary>
        /// If the header row has an autofilter
        /// </summary>
        public bool ShowFilter 
        { 
            get
            {
                return ShowHeader && AutoFilterAddress != null;
            }
            set
            {
                if (ShowHeader)
                {
                    if (value)
                    {
                        WriteAutoFilter(ShowTotal);
                    }
                    else 
                    {
                        DeleteAllNode(AUTOFILTER_PATH);
                    }
                }
                else if(value)
                {
                    throw(new InvalidOperationException("Filter can only be applied when ShowHeader is set to true"));
                }
            }
        }
        const string TOTALSROWCOUNT_PATH = "@totalsRowCount";
        const string TOTALSROWSHOWN_PATH = "@totalsRowShown";
        /// <summary>
        /// If the total row is visible or not
        /// </summary>
        public bool ShowTotal
        {
            get
            {
                return GetXmlNodeInt(TOTALSROWCOUNT_PATH) == 1;
            }
            set
            {
                if (value != ShowTotal)
                {
                    if (value)
                    {
                        Address=new ExcelAddress(WorkSheet.Name, ExcelAddressBase.GetAddress(Address.Start.Row, Address.Start.Column, Address.End.Row+1, Address.End.Column));
                    }
                    else
                    {
                        Address = new ExcelAddress(WorkSheet.Name, ExcelAddressBase.GetAddress(Address.Start.Row, Address.Start.Column, Address.End.Row - 1, Address.End.Column));
                    }
                    SetXmlNodeString("@ref", Address.Address);
                    if (value)
                    {
                        SetXmlNodeString(TOTALSROWCOUNT_PATH, "1");
                    }
                    else
                    {
                        DeleteNode(TOTALSROWCOUNT_PATH);
                    }
                    WriteAutoFilter(value);
                }
            }
        }
        const string STYLENAME_PATH = "d:tableStyleInfo/@name";
        /// <summary>
        /// The style name for custum styles
        /// </summary>
        public string StyleName
        {
            get
            {
                return GetXmlNodeString(STYLENAME_PATH);
            }
            set
            {
                if (value.StartsWith("TableStyle"))
                {
                    try
                    {
                        _tableStyle = (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10,value.Length-10), true);
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
                SetXmlNodeString(STYLENAME_PATH,value,true);
            }
        }
        const string SHOWFIRSTCOLUMN_PATH = "d:tableStyleInfo/@showFirstColumn";
        /// <summary>
        /// Display special formatting for the first row
        /// </summary>
        public bool ShowFirstColumn
        {
            get
            {
                return GetXmlNodeBool(SHOWFIRSTCOLUMN_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWFIRSTCOLUMN_PATH, value, false);
            }   
        }
        const string SHOWLASTCOLUMN_PATH = "d:tableStyleInfo/@showLastColumn";
        /// <summary>
        /// Display special formatting for the last row
        /// </summary>
        public bool ShowLastColumn
        {
            get
            {
                return GetXmlNodeBool(SHOWLASTCOLUMN_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWLASTCOLUMN_PATH, value, false);
            }
        }
        const string SHOWROWSTRIPES_PATH = "d:tableStyleInfo/@showRowStripes";
        /// <summary>
        /// Display banded rows
        /// </summary>
        public bool ShowRowStripes
        {
            get
            {
                return GetXmlNodeBool(SHOWROWSTRIPES_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWROWSTRIPES_PATH, value, false);
            }
        }
        const string SHOWCOLUMNSTRIPES_PATH = "d:tableStyleInfo/@showColumnStripes";
        /// <summary>
        /// Display banded columns
        /// </summary>
        public bool ShowColumnStripes
        {
            get
            {
                return GetXmlNodeBool(SHOWCOLUMNSTRIPES_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWCOLUMNSTRIPES_PATH, value, false);
            }
        }

        const string TOTALSROWCELLSTYLE_PATH = "@totalsRowCellStyle";
        /// <summary>
        /// Named style used for the total row
        /// </summary>
        public string TotalsRowCellStyle
        {
            get
            {
                return GetXmlNodeString(TOTALSROWCELLSTYLE_PATH);
            }
            set
            {
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
                {
                    throw (new Exception(string.Format("Named style {0} does not exist.", value)));
                }
                SetXmlNodeString(TopNode, TOTALSROWCELLSTYLE_PATH, value, true);

                if (ShowTotal)
                {
                    WorkSheet.Cells[Address._toRow, Address._fromCol, Address._toRow, Address._toCol].StyleName = value;
                }
            }
        }
        const string DATACELLSTYLE_PATH = "@dataCellStyle";
        /// <summary>
        /// Named style used for the data cells
        /// </summary>
        public string DataCellStyleName
        {
            get
            {
                return GetXmlNodeString(DATACELLSTYLE_PATH);
            }
            set
            {
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
                {
                    throw (new Exception(string.Format("Named style {0} does not exist.", value)));
                }
                SetXmlNodeString(TopNode, DATACELLSTYLE_PATH, value, true);

                int fromRow = Address._fromRow + (ShowHeader ? 1 : 0),
                    toRow = Address._toRow - (ShowTotal ? 1 : 0);

                if (fromRow < toRow)
                {
                    WorkSheet.Cells[fromRow, Address._fromCol, toRow, Address._toCol].StyleName = value;
                }
            }
        }
        const string HEADERROWCELLSTYLE_PATH = "@headerRowCellStyle";
        /// <summary>
        /// Named style used for the header row
        /// </summary>
        public string HeaderRowCellStyle
        {
            get
            {
                return GetXmlNodeString(HEADERROWCELLSTYLE_PATH);
            }
            set
            {
                if (WorkSheet.Workbook.Styles.NamedStyles.FindIndexByID(value) < 0)
                {
                    throw (new Exception(string.Format("Named style {0} does not exist.", value)));
                }
                SetXmlNodeString(TopNode, HEADERROWCELLSTYLE_PATH, value, true);

                if (ShowHeader)
                {
                    WorkSheet.Cells[Address._fromRow, Address._fromCol, Address._fromRow, Address._toCol].StyleName = value;
                }

            }
        }        
    }
}
