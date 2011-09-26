using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using System.IO.Packaging;

namespace OfficeOpenXml.Drawing.Vml
{
    public class ExcelVmlDrawingBaseCollection
    {        
        internal ExcelVmlDrawingBaseCollection(ExcelPackage pck, ExcelWorksheet ws, Uri uri)
        {
            VmlDrawingXml = new XmlDocument();
            VmlDrawingXml.PreserveWhitespace = false;
            
            NameTable nt=new NameTable();
            NameSpaceManager = new XmlNamespaceManager(nt);
            NameSpaceManager.AddNamespace("v", ExcelPackage.schemaMicrosoftVml);
            NameSpaceManager.AddNamespace("o", ExcelPackage.schemaMicrosoftOffice);
            NameSpaceManager.AddNamespace("x", ExcelPackage.schemaMicrosoftExcel);
            Uri = uri;
            if (uri == null)
            {
                Part = null;
            }
            else
            {
                Part=pck.Package.GetPart(uri);
                VmlDrawingXml.Load(Part.GetStream());
            }
        }
        internal XmlDocument VmlDrawingXml { get; set; }
        internal Uri Uri { get; set; }
        internal string RelId { get; set; }
        internal PackagePart Part { get; set; }
        internal XmlNamespaceManager NameSpaceManager { get; set; }
    }
}
