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
using System.Data.SqlClient;
using System.Text;
using System.Xml;
using System.Linq;
using OfficeOpenXml.Utils;
namespace OfficeOpenXml.Table.PivotTable
{
    public enum eSourceType
    {
        /// <summary>
        /// Indicates that the cache contains data that consolidates ranges.
        /// </summary>
        Consolidation,
        /// <summary>
        /// Indicates that the cache contains data from an external data source.
        /// </summary>
        External,
        /// <summary>
        /// Indicates that the cache contains a scenario summary report
        /// </summary>
        Scenario,
        /// <summary>
        /// Indicates that the cache contains worksheet data
        /// </summary>
        Worksheet
    }
    /// <summary>
    /// Cache definition. This class defines the source data. Note that one cache definition can be shared between many pivot tables.
    /// </summary>
    public class ExcelPivotCacheDefinition : XmlHelper
    {
        internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable) :
            base(ns, null)
        {
            foreach (var r in pivotTable.Part.GetRelationshipsByType(ExcelPackage.schemaRelationships + "/pivotCacheDefinition"))
            {
                Relationship = r;
            }
            CacheDefinitionUri = UriHelper.ResolvePartUri(Relationship.SourceUri, Relationship.TargetUri);

            var pck = pivotTable.WorkSheet._package.Package;
            Part = pck.GetPart(CacheDefinitionUri);
            CacheDefinitionXml = new XmlDocument();
            LoadXmlSafe(CacheDefinitionXml, Part.GetStream());

            TopNode = CacheDefinitionXml.DocumentElement;
            PivotTable = pivotTable;
            if (CacheSource == eSourceType.Worksheet)
            {
                var worksheetName = GetXmlNodeString(_sourceWorksheetPath);
                if (pivotTable.WorkSheet.Workbook.Worksheets.Any(t => t.Name == worksheetName))
                {
                    _sourceRange = pivotTable.WorkSheet.Workbook.Worksheets[worksheetName].Cells[GetXmlNodeString(_sourceAddressPath)];
                }
            }
        }
        internal ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable, ExcelRangeBase sourceAddress, int tblId) :
            base(ns, null)
        {
            PivotTable = pivotTable;

            var pck = pivotTable.WorkSheet._package.Package;
            
            //CacheDefinition
            CacheDefinitionXml = new XmlDocument();
            LoadXmlSafe(CacheDefinitionXml, GetStartXml(sourceAddress), Encoding.UTF8);
            CacheDefinitionUri = GetNewUri(pck, "/xl/pivotCache/pivotCacheDefinition{0}.xml", tblId); 
            Part = pck.CreatePart(CacheDefinitionUri, ExcelPackage.schemaPivotCacheDefinition);
            TopNode = CacheDefinitionXml.DocumentElement;

            //CacheRecord. Create an empty one.
            CacheRecordUri = GetNewUri(pck, "/xl/pivotCache/pivotCacheRecords{0}.xml", tblId); 
            var cacheRecord = new XmlDocument();
            cacheRecord.LoadXml("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");
            var recPart = pck.CreatePart(CacheRecordUri, ExcelPackage.schemaPivotCacheRecords);
            cacheRecord.Save(recPart.GetStream());

            RecordRelationship = Part.CreateRelationship(UriHelper.ResolvePartUri(CacheDefinitionUri, CacheRecordUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");
            RecordRelationshipID = RecordRelationship.Id;

            CacheDefinitionXml.Save(Part.GetStream());
        }        
        /// <summary>
        /// Reference to the internal package part
        /// </summary>
        internal Packaging.ZipPackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Provides access to the XML data representing the cache definition in the package.
        /// </summary>
        public XmlDocument CacheDefinitionXml { get; private set; }
        /// <summary>
        /// The package internal URI to the pivottable cache definition Xml Document.
        /// </summary>
        public Uri CacheDefinitionUri
        {
            get;
            internal set;
        }
        internal Uri CacheRecordUri
        {
            get;
            set;
        }
        internal Packaging.ZipPackageRelationship Relationship
        {
            get;
            set;
        }
        internal Packaging.ZipPackageRelationship RecordRelationship
        {
            get;
            set;
        }
        internal string RecordRelationshipID 
        {
            get
            {
                return GetXmlNodeString("@r:id");
            }
            set
            {
                SetXmlNodeString("@r:id", value);
            }
        }
        /// <summary>
        /// Referece to the PivoTable object
        /// </summary>
        public ExcelPivotTable PivotTable
        {
            get;
            private set;
        }
        
        const string _sourceWorksheetPath="d:cacheSource/d:worksheetSource/@sheet";
        const string _sourceNamePath = "d:cacheSource/d:worksheetSource/@name";
        const string _sourceAddressPath = "d:cacheSource/d:worksheetSource/@ref";
        internal ExcelRangeBase _sourceRange = null;
        /// <summary>
        /// The source data range when the pivottable has a worksheet datasource. 
        /// The number of columns in the range must be intact if this property is changed.
        /// The range must be in the same workbook as the pivottable.
        /// </summary>
        public ExcelRangeBase SourceRange
        {
            get
            {
                if (_sourceRange == null)
                {
                    if (CacheSource == eSourceType.Worksheet)
                    {
                        var ws = PivotTable.WorkSheet.Workbook.Worksheets[GetXmlNodeString(_sourceWorksheetPath)];
                        if (ws == null) //Not worksheet, check name or table name
                        {
                            var name = GetXmlNodeString(_sourceNamePath);
                            foreach (var n in PivotTable.WorkSheet.Workbook.Names)
                            {
                                if(name.Equals(n.Name,StringComparison.InvariantCultureIgnoreCase))
                                {
                                    _sourceRange = n;
                                    return _sourceRange;
                                }
                            }
                            foreach (var w in PivotTable.WorkSheet.Workbook.Worksheets)
                            {
                                if (w.Tables._tableNames.ContainsKey(name))
                                {
                                    _sourceRange = w.Cells[w.Tables[name].Address.Address];
                                    break;
                                }
                                foreach (var n in w.Names)
                                {
                                    if (name.Equals(n.Name, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        _sourceRange = n;
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            _sourceRange = ws.Cells[GetXmlNodeString(_sourceAddressPath)];
                        }
                    }
                    else
                    {
                        throw (new ArgumentException("The cachesource is not a worksheet"));
                    }
                }
                return _sourceRange;
            }
            set
            {
                if (PivotTable.WorkSheet.Workbook != value.Worksheet.Workbook)
                {
                    throw (new ArgumentException("Range must be in the same package as the pivottable"));
                }

                var sr=SourceRange;
                if (value.End.Column - value.Start.Column != sr.End.Column - sr.Start.Column)
                {
                    throw (new ArgumentException("Can not change the number of columns(fields) in the SourceRange"));
                }

                SetXmlNodeString(_sourceWorksheetPath, value.Worksheet.Name);
                SetXmlNodeString(_sourceAddressPath, value.FirstAddress);
                _sourceRange = value;
            }
        }
        /// <summary>
        /// Type of source data
        /// </summary>
        public eSourceType CacheSource
        {
            get
            {
                var s=GetXmlNodeString("d:cacheSource/@type");
                if (s == "")
                {
                    return eSourceType.Worksheet;
                }
                else
                {
                    return (eSourceType)Enum.Parse(typeof(eSourceType), s, true);
                }
            }
        }
        private string GetStartXml(ExcelRangeBase sourceAddress)
        {
            string xml="<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"\" refreshOnLoad=\"1\" refreshedBy=\"SomeUser\" refreshedDate=\"40504.582403125001\" createdVersion=\"1\" refreshedVersion=\"3\" recordCount=\"5\" upgradeOnRefresh=\"1\">";

            xml += "<cacheSource type=\"worksheet\">";
            xml += string.Format("<worksheetSource ref=\"{0}\" sheet=\"{1}\" /> ", sourceAddress.Address, sourceAddress.WorkSheet);
            xml += "</cacheSource>";
            xml += string.Format("<cacheFields count=\"{0}\">", sourceAddress._toCol - sourceAddress._fromCol + 1);
            var sourceWorksheet = PivotTable.WorkSheet.Workbook.Worksheets[sourceAddress.WorkSheet];
            for (int col = sourceAddress._fromCol; col <= sourceAddress._toCol; col++)
            {
                if (sourceWorksheet == null || sourceWorksheet._values.GetValue(sourceAddress._fromRow, col) == null || sourceWorksheet._values.GetValue(sourceAddress._fromRow, col).ToString().Trim() == "")
                {
                    xml += string.Format("<cacheField name=\"Column{0}\" numFmtId=\"0\">", col - sourceAddress._fromCol + 1);
                }
                else
                {
                    xml += string.Format("<cacheField name=\"{0}\" numFmtId=\"0\">", sourceWorksheet._values.GetValue(sourceAddress._fromRow, col));
                }
                //xml += "<sharedItems containsNonDate=\"0\" containsString=\"0\" containsBlank=\"1\" /> ";
                xml += "<sharedItems containsBlank=\"1\" /> ";
                xml += "</cacheField>";
            }
            xml += "</cacheFields>";
            xml += "</pivotCacheDefinition>";

            return xml;
        }
    }
}
