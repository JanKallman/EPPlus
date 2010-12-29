using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;

namespace OfficeOpenXml.Table
{
    public class ExcelPivotCacheDefinition : XmlHelper
    {
        public ExcelPivotCacheDefinition(XmlNamespaceManager ns, ExcelPivotTable pivotTable, ExcelAddressBase sourceAddress, int tblId) :
            base(ns, null)
        {
            SourceAddress = sourceAddress;
            PivotTable = pivotTable;

            var pck = pivotTable.WorkSheet.xlPackage.Package;
            
            //CacheDefinition
            CacheDefinitionXml = new XmlDocument();
            CacheDefinitionXml.LoadXml(GetStartXml());
            CacheDefinitionUri=new Uri(string.Format("/xl/pivotCache/pivotCacheDefinition{0}.xml", tblId), UriKind.Relative);
            Part = pck.CreatePart(CacheDefinitionUri, ExcelPackage.schemaPivotCacheDefinition);
            TopNode = CacheDefinitionXml.DocumentElement;

            //CacheRecord. Create an empty one.
            CacheRecordUri = new Uri(string.Format("/xl/pivotCache/pivotCacheRecords{0}.xml", tblId), UriKind.Relative);
            var cacheRecord = new XmlDocument();
            cacheRecord.LoadXml("<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />");
            var recPart = pck.CreatePart(CacheRecordUri, ExcelPackage.schemaPivotCacheRecords);
            cacheRecord.Save(recPart.GetStream());

            RecordRelationship = Part.CreateRelationship(PackUriHelper.ResolvePartUri(CacheDefinitionUri, CacheRecordUri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");
            RecordRelationshipID = RecordRelationship.Id;

            CacheDefinitionXml.Save(Part.GetStream());
        }        
        internal PackagePart Part
        {
            get;
            set;
        }
        public XmlDocument CacheDefinitionXml { get; private set; }
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
        internal PackageRelationship Relationship
        {
            get;
            set;
        }
        internal PackageRelationship RecordRelationship
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
        /// <summary>
        /// The address to the Source data
        /// </summary>
        internal ExcelAddressBase SourceAddress
        {
            get;
            private set;
        }
        private string GetStartXml()
        {
            string xml="<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"\" refreshOnLoad=\"1\" refreshedBy=\"SomeUser\" refreshedDate=\"40504.582403125001\" createdVersion=\"1\" refreshedVersion=\"3\" recordCount=\"5\" upgradeOnRefresh=\"1\">";

            xml += "<cacheSource type=\"worksheet\">";
            xml += string.Format("<worksheetSource ref=\"{0}\" sheet=\"{1}\" /> ", SourceAddress.Address, SourceAddress.WorkSheet);
            xml += "</cacheSource>";
            xml += string.Format("<cacheFields count=\"{0}\">",SourceAddress._toCol - SourceAddress._fromCol + 1);
            var sourceWorksheet = PivotTable.WorkSheet.Workbook.Worksheets[SourceAddress.WorkSheet];
            for (int col = SourceAddress._fromCol; col <= SourceAddress._toCol; col++)
            {
                if (sourceWorksheet==null || sourceWorksheet.Cell(SourceAddress._fromRow, col).Value == null || sourceWorksheet.Cell(SourceAddress._fromRow, col).Value.ToString().Trim() == "")
                {
                    xml += string.Format("<cacheField name=\"Column{0}\" numFmtId=\"0\">", col - SourceAddress._fromCol+1);
                }
                else
                {
                    xml += string.Format("<cacheField name=\"{0}\" numFmtId=\"0\">", sourceWorksheet.Cell(SourceAddress._fromRow, col).Value);
                }
                xml += "<sharedItems containsNonDate=\"0\" containsString=\"0\" containsBlank=\"1\" /> ";
                xml += "</cacheField>";
            }
            xml += "</cacheFields>";
            xml += "</pivotCacheDefinition>";

            return xml;
        }
    }
}
