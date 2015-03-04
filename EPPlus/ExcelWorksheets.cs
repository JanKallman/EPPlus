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
 * Jan Källman		    Initial Release		       2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;
using System.IO;
using System.Linq;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Packaging.Ionic.Zlib;
using OfficeOpenXml.Utils;
using OfficeOpenXml.VBA;
namespace OfficeOpenXml
{
	/// <summary>
	/// The collection of worksheets for the workbook
	/// </summary>
	public class ExcelWorksheets : XmlHelper, IEnumerable<ExcelWorksheet>, IDisposable
	{
		#region Private Properties
        private ExcelPackage _pck;
        private Dictionary<int, ExcelWorksheet> _worksheets;
		private XmlNamespaceManager _namespaceManager;
		#endregion
		#region ExcelWorksheets Constructor
		internal ExcelWorksheets(ExcelPackage pck, XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
		{
			_pck = pck;
            _namespaceManager = nsm;
			_worksheets = new Dictionary<int, ExcelWorksheet>();
			int positionID = 1;

            foreach (XmlNode sheetNode in topNode.ChildNodes)
			{
                if (sheetNode.NodeType == XmlNodeType.Element)
                {
                    string name = sheetNode.Attributes["name"].Value;
                    //Get the relationship id
                    string relId = sheetNode.Attributes["r:id"].Value;
                    int sheetID = Convert.ToInt32(sheetNode.Attributes["sheetId"].Value);

                    //Hidden property
                    eWorkSheetHidden hidden = eWorkSheetHidden.Visible;
                    XmlNode attr = sheetNode.Attributes["state"];
                    if (attr != null)
                        hidden = TranslateHidden(attr.Value);

                    var sheetRelation = pck.Workbook.Part.GetRelationship(relId);
                    Uri uriWorksheet = UriHelper.ResolvePartUri(pck.Workbook.WorkbookUri, sheetRelation.TargetUri);

                    //add the worksheet
                    if (sheetRelation.RelationshipType.EndsWith("chartsheet"))
                    {
                        _worksheets.Add(positionID, new ExcelChartsheet(_namespaceManager, _pck, relId, uriWorksheet, name, sheetID, positionID, hidden));
                    }
                    else
                    {
                        _worksheets.Add(positionID, new ExcelWorksheet(_namespaceManager, _pck, relId, uriWorksheet, name, sheetID, positionID, hidden));
                    }
                    positionID++;
                }
			}
		}

        private eWorkSheetHidden TranslateHidden(string value)
        {
            switch (value)
            {
                case "hidden":
                    return eWorkSheetHidden.Hidden;
                case "veryHidden":
                    return eWorkSheetHidden.VeryHidden;
                default:
                    return eWorkSheetHidden.Visible;
            }
        }
		#endregion
		#region ExcelWorksheets Public Properties
		/// <summary>
		/// Returns the number of worksheets in the workbook
		/// </summary>
		public int Count
		{
			get { return (_worksheets.Count); }
		}
		#endregion
        private const string ERR_DUP_WORKSHEET = "A worksheet with this name already exists in the workbook";
        internal const string WORKSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        internal const string CHARTSHEET_CONTENTTYPE = @"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
		#region ExcelWorksheets Public Methods
		/// <summary>
        /// Foreach support
		/// </summary>
		/// <returns>An enumerator</returns>
		public IEnumerator<ExcelWorksheet> GetEnumerator()
		{
			return (_worksheets.Values.GetEnumerator());
        }
        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return (_worksheets.Values.GetEnumerator());
        }

        #endregion
		#region Add Worksheet
		/// <summary>
		/// Adds a new blank worksheet.
		/// </summary>
		/// <param name="Name">The name of the workbook</param>
		public ExcelWorksheet Add(string Name)
		{
            ExcelWorksheet worksheet = AddSheet(Name,false, null);
			return worksheet;
		}
        private ExcelWorksheet AddSheet(string Name, bool isChart, eChartType? chartType)
        {
            int sheetID;
            Uri uriWorksheet;
            lock (_worksheets)
            {
                Name = ValidateFixSheetName(Name);
                if (GetByName(Name) != null)
                {
                    throw (new InvalidOperationException(ERR_DUP_WORKSHEET + " : " + Name));
                }
                GetSheetURI(ref Name, out sheetID, out uriWorksheet, isChart);
                Packaging.ZipPackagePart worksheetPart = _pck.Package.CreatePart(uriWorksheet, isChart ? CHARTSHEET_CONTENTTYPE : WORKSHEET_CONTENTTYPE, _pck.Compression);

                //Create the new, empty worksheet and save it to the package
                StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
                XmlDocument worksheetXml = CreateNewWorksheet(isChart);
                worksheetXml.Save(streamWorksheet);
                _pck.Package.Flush();

                string rel = CreateWorkbookRel(Name, sheetID, uriWorksheet, isChart);

                int positionID = _worksheets.Count + 1;
                ExcelWorksheet worksheet;
                if (isChart)
                {
                    worksheet = new ExcelChartsheet(_namespaceManager, _pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible, (eChartType)chartType);
                }
                else
                {
                    worksheet = new ExcelWorksheet(_namespaceManager, _pck, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible);
                }

                _worksheets.Add(positionID, worksheet);
#if !MONO
                if (_pck.Workbook.VbaProject != null)
                {
                    var name = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(worksheet);
                    _pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(worksheet.CodeNameChange) { Name = name, Code = "", Attributes = _pck.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
                    worksheet.CodeModuleName = name;

                }
#endif
                return worksheet;
            }
        }
        /// <summary>
        /// Adds a copy of a worksheet
        /// </summary>
        /// <param name="Name">The name of the workbook</param>
        /// <param name="Copy">The worksheet to be copied</param>
        public ExcelWorksheet Add(string Name, ExcelWorksheet Copy)
        {
            lock (_worksheets)
            {
                int sheetID;
                Uri uriWorksheet;
                if (Copy is ExcelChartsheet)
                {
                    throw (new ArgumentException("Can not copy a chartsheet"));
                }
                if (GetByName(Name) != null)
                {
                    throw (new InvalidOperationException(ERR_DUP_WORKSHEET));
                }

                GetSheetURI(ref Name, out sheetID, out uriWorksheet, false);

                //Create a copy of the worksheet XML
                Packaging.ZipPackagePart worksheetPart = _pck.Package.CreatePart(uriWorksheet, WORKSHEET_CONTENTTYPE, _pck.Compression);
                StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
                XmlDocument worksheetXml = new XmlDocument();
                worksheetXml.LoadXml(Copy.WorksheetXml.OuterXml);
                worksheetXml.Save(streamWorksheet);
                //streamWorksheet.Close();
                _pck.Package.Flush();


                //Create a relation to the workbook
                string relID = CreateWorkbookRel(Name, sheetID, uriWorksheet, false);
                ExcelWorksheet added = new ExcelWorksheet(_namespaceManager, _pck, relID, uriWorksheet, Name, sheetID, _worksheets.Count + 1, eWorkSheetHidden.Visible);

                //Copy comments
                if (Copy.Comments.Count > 0)
                {
                    CopyComment(Copy, added);
                }
                else if (Copy.VmlDrawingsComments.Count > 0)    //Vml drawings are copied as part of the comments. 
                {
                    CopyVmlDrawing(Copy, added);
                }

                //Copy HeaderFooter
                CopyHeaderFooterPictures(Copy, added);

                //Copy all relationships 
                //CopyRelationShips(Copy, added);
                if (Copy.Drawings.Count > 0)
                {
                    CopyDrawing(Copy, added);
                }
                if (Copy.Tables.Count > 0)
                {
                    CopyTable(Copy, added);
                }
                if (Copy.PivotTables.Count > 0)
                {
                    CopyPivotTable(Copy, added);
                }
                if (Copy.Names.Count > 0)
                {
                    CopySheetNames(Copy, added);
                }

                //Copy all cells
                CloneCells(Copy, added);

                //Copy the VBA code
#if !MONO
                if (_pck.Workbook.VbaProject != null)
                {
                    var name = _pck.Workbook.VbaProject.GetModuleNameFromWorksheet(added);
                    _pck.Workbook.VbaProject.Modules.Add(new ExcelVBAModule(added.CodeNameChange) { Name = name, Code = Copy.CodeModule.Code, Attributes = _pck.Workbook.VbaProject.GetDocumentAttributes(Name, "0{00020820-0000-0000-C000-000000000046}"), Type = eModuleType.Document, HelpContext = 0 });
                    Copy.CodeModuleName = name;
                }
#endif

                _worksheets.Add(_worksheets.Count + 1, added);

                //Remove any relation to printersettings.
                XmlNode pageSetup = added.WorksheetXml.SelectSingleNode("//d:pageSetup", _namespaceManager);
                if (pageSetup != null)
                {
                    XmlAttribute attr = (XmlAttribute)pageSetup.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships);
                    if (attr != null)
                    {
                        relID = attr.Value;
                        // first delete the attribute from the XML
                        pageSetup.Attributes.Remove(attr);
                    }
                }
                return added;
            }
        }
        public ExcelChartsheet AddChart(string Name, eChartType chartType)
        {
            return (ExcelChartsheet)AddSheet(Name, true, chartType);
        }
        private void CopySheetNames(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            foreach (var name in Copy.Names)
            {
                ExcelNamedRange newName;
                if (!name.IsName)
                {
                    if (name.WorkSheet == Copy.Name)
                    {
                        newName = added.Names.Add(name.Name, added.Cells[name.FirstAddress]);
                    }
                    else
                    {
                        newName = added.Names.Add(name.Name, added.Workbook.Worksheets[name.WorkSheet].Cells[name.FirstAddress]);
                    }
                }
                else if (!string.IsNullOrEmpty(name.NameFormula))
                {
                    newName=added.Names.AddFormula(name.Name, name.Formula);
                }
                else
                {
                    newName=added.Names.AddValue(name.Name, name.Value);
                }
               newName.NameComment = name.NameComment;
            }
        }

        private void CopyTable(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            string prevName = "";
            //First copy the table XML
            foreach (var tbl in Copy.Tables)
            {
                string xml=tbl.TableXml.OuterXml;
                int Id = _pck.Workbook._nextTableID++;
                string name;
                if (prevName == "")
                {
                    name = Copy.Tables.GetNewTableName();
                }
                else
                {
                    int ix = int.Parse(prevName.Substring(5)) + 1;
                    name = string.Format("Table{0}", ix);
                    while (_pck.Workbook.ExistsPivotTableName(name))
                    {
                        name = string.Format("Table{0}", ++ix);
                    }
                }
                prevName = name;
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);
                xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
                xmlDoc.SelectSingleNode("//d:table/@name", tbl.NameSpaceManager).Value = name;
                xmlDoc.SelectSingleNode("//d:table/@displayName", tbl.NameSpaceManager).Value = name;
                xml = xmlDoc.OuterXml;

                var uriTbl = new Uri(string.Format("/xl/tables/table{0}.xml", Id), UriKind.Relative);
                var part = _pck.Package.CreatePart(uriTbl, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", _pck.Compression);
                StreamWriter streamTbl = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                streamTbl.Write(xml);
                //streamTbl.Close();
                streamTbl.Flush();

                //create the relationship and add the ID to the worksheet xml.
                var rel = added.Part.CreateRelationship(UriHelper.GetRelativeUri(added.WorksheetUri,uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");

                if (tbl.RelationshipID == null)
                {
                    var topNode = added.WorksheetXml.SelectSingleNode("//d:tableParts", tbl.NameSpaceManager);
                    if (topNode == null)
                    {
                        added.CreateNode("d:tableParts");
                        topNode = added.WorksheetXml.SelectSingleNode("//d:tableParts", tbl.NameSpaceManager);
                    }
                    XmlElement elem = added.WorksheetXml.CreateElement("tablePart", ExcelPackage.schemaMain);
                    topNode.AppendChild(elem);
                    elem.SetAttribute("id",ExcelPackage.schemaRelationships, rel.Id);
                }
                else
                {
                    XmlAttribute relAtt;
                    relAtt = added.WorksheetXml.SelectSingleNode(string.Format("//d:tableParts/d:tablePart/@r:id[.='{0}']", tbl.RelationshipID), tbl.NameSpaceManager) as XmlAttribute;
                    relAtt.Value = rel.Id;
                }
            }
        }
        private void CopyPivotTable(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            string prevName = "";
            foreach (var tbl in Copy.PivotTables)
            {
                string xml = tbl.PivotTableXml.OuterXml;
                int Id = _pck.Workbook._nextPivotTableID++;

                string name;
                if (prevName == "")
                {
                    name = Copy.PivotTables.GetNewTableName();
                }
                else
                {
                    int ix=int.Parse(prevName.Substring(10))+1;
                    name = string.Format("PivotTable{0}", ix);
                    while (_pck.Workbook.ExistsPivotTableName(name))
                    {
                        name = string.Format("PivotTable{0}", ++ix);
                    }
                }
                prevName=name;
                XmlDocument xmlDoc = new XmlDocument();
                //TODO: Fix save pivottable here
                //Copy.Save();    //Save the worksheet first
                xmlDoc.LoadXml(xml);
                //xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
                xmlDoc.SelectSingleNode("//d:pivotTableDefinition/@name", tbl.NameSpaceManager).Value = name;
                xml = xmlDoc.OuterXml;

                var uriTbl = new Uri(string.Format("/xl/pivotTables/pivotTable{0}.xml", Id), UriKind.Relative);
                var partTbl = _pck.Package.CreatePart(uriTbl, ExcelPackage.schemaPivotTable , _pck.Compression);
                StreamWriter streamTbl = new StreamWriter(partTbl.GetStream(FileMode.Create, FileAccess.Write));
                streamTbl.Write(xml);
                //streamTbl.Close();
                streamTbl.Flush();

                xml = tbl.CacheDefinition.CacheDefinitionXml.OuterXml;                
                var uriCd = new Uri(string.Format("/xl/pivotCache/pivotcachedefinition{0}.xml", Id), UriKind.Relative);
                while (_pck.Package.PartExists(uriCd))
                {
                    uriCd = new Uri(string.Format("/xl/pivotCache/pivotcachedefinition{0}.xml", ++Id), UriKind.Relative);
                }

                var partCd = _pck.Package.CreatePart(uriCd, ExcelPackage.schemaPivotCacheDefinition, _pck.Compression);
                StreamWriter streamCd = new StreamWriter(partCd.GetStream(FileMode.Create, FileAccess.Write));
                streamCd.Write(xml);
                streamCd.Flush();

                xml = "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" count=\"0\" />";
                var uriRec = new Uri(string.Format("/xl/pivotCache/pivotrecords{0}.xml", Id), UriKind.Relative);
                while (_pck.Package.PartExists(uriRec))
                {
                    uriRec = new Uri(string.Format("/xl/pivotCache/pivotrecords{0}.xml", ++Id), UriKind.Relative);
                }
                var partRec = _pck.Package.CreatePart(uriRec, ExcelPackage.schemaPivotCacheRecords, _pck.Compression);
                StreamWriter streamRec = new StreamWriter(partRec.GetStream(FileMode.Create, FileAccess.Write));
                streamRec.Write(xml);
                streamRec.Flush();

                //create the relationship and add the ID to the worksheet xml.
                added.Part.CreateRelationship(UriHelper.ResolvePartUri(added.WorksheetUri, uriTbl), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");
                partTbl.CreateRelationship(UriHelper.ResolvePartUri(tbl.Relationship.SourceUri, uriCd), tbl.CacheDefinition.Relationship.TargetMode, tbl.CacheDefinition.Relationship.RelationshipType);
                partCd.CreateRelationship(UriHelper.ResolvePartUri(uriCd, uriRec), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheRecords");
            }
        }
        private void CopyHeaderFooterPictures(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            if (Copy._headerFooter == null) return;
            //Copy the texts
            CopyText(Copy.HeaderFooter._oddHeader, added.HeaderFooter.OddHeader);
            CopyText(Copy.HeaderFooter._oddFooter, added.HeaderFooter.OddFooter);
            CopyText(Copy.HeaderFooter._evenHeader, added.HeaderFooter.EvenHeader);
            CopyText(Copy.HeaderFooter._evenFooter, added.HeaderFooter.EvenFooter);
            CopyText(Copy.HeaderFooter._firstHeader, added.HeaderFooter.FirstHeader);
            CopyText(Copy.HeaderFooter._firstFooter, added.HeaderFooter.FirstFooter);
            
            //Copy any images;
            if (Copy.HeaderFooter.Pictures.Count > 0)
            {
                Uri source = Copy.HeaderFooter.Pictures.Uri;
                Uri dest = XmlHelper.GetNewUri(_pck.Package, @"/xl/drawings/vmlDrawing{0}.vml");
                
                //var part = _pck.Package.CreatePart(dest, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
                foreach (ExcelVmlDrawingPicture pic in Copy.HeaderFooter.Pictures)
                {
                    var item = added.HeaderFooter.Pictures.Add(pic.Id, pic.ImageUri, pic.Title, pic.Width, pic.Height);
                    foreach (XmlAttribute att in pic.TopNode.Attributes)
                    {
                        (item.TopNode as XmlElement).SetAttribute(att.Name, att.Value);
                    }
                    item.TopNode.InnerXml = pic.TopNode.InnerXml;
                }
            }
        }

        private void CopyText(ExcelHeaderFooterText from, ExcelHeaderFooterText to)
        {
            if (from == null) return;
            to.LeftAlignedText=from.LeftAlignedText;
            to.CenteredText = from.CenteredText;
            to.RightAlignedText = from.RightAlignedText;
        }
        private void CloneCells(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            bool sameWorkbook=(Copy.Workbook == _pck.Workbook);

            bool doAdjust = _pck.DoAdjustDrawings;
            _pck.DoAdjustDrawings = false;
            added.MergedCells.List.AddRange(Copy.MergedCells.List);
            //Formulas
            //foreach (IRangeID f in Copy._formulaCells)
            //{
            //    added._formulaCells.Add(f);
            //}
            //Shared Formulas
            foreach (int key in Copy._sharedFormulas.Keys)
            {
                added._sharedFormulas.Add(key, Copy._sharedFormulas[key]);
            }
            
            Dictionary<int, int> styleCashe = new Dictionary<int, int>();
            //Cells
            int row,col;
            var val = new CellsStoreEnumerator<object>(Copy._values);
            //object f=null;
            //foreach (var addr in val)
            while(val.Next())
            {                
                //row=(int)addr>>32;
                //col=(int)addr&32;
                row = val.Row;
                col = val.Column;
                //added._cells.Add(cell.Clone(added));
                int styleID=0;
                if (row == 0) //Column
                {
                    var c = Copy._values.GetValue(row, col) as ExcelColumn;
                    if (c != null)
                    {
                        var clone = c.Clone(added, c.ColumnMin);
                        clone.StyleID = c.StyleID;
                        added._values.SetValue(row, col, clone);
                        styleID = c.StyleID;
                    }
                }
                else if (col == 0) //Row
                {
                    var r=Copy.Row(row);
                    if (r != null)
                    {
                        r.Clone(added);
                        styleID = r.StyleID;
                        //added._values.SetValue(row, col, r.Clone(added));                                                
                    }
                    
                }
                else
                {
                   styleID = CopyValues(Copy, added, row, col);
                }
                if (!sameWorkbook)
                {
                    if (styleCashe.ContainsKey(styleID))
                    {
                        added._styles.SetValue(row, col, styleCashe[styleID]);
                    }
                    else
                    {
                        var s = added.Workbook.Styles.CloneStyle(Copy.Workbook.Styles, styleID);
                        styleCashe.Add(styleID, s);
                        added._styles.SetValue(row, col, s);
                    }
                }
            }
            added._package.DoAdjustDrawings = doAdjust;
        }

        private int CopyValues(ExcelWorksheet Copy, ExcelWorksheet added, int row, int col)
        {
            added._values.SetValue(row, col, Copy._values.GetValue(row, col));
            var t = Copy._types.GetValue(row, col);
            if (t != null)
            {
                added._types.SetValue(row, col, t);
            }
            byte fl=0;
            if (Copy._flags.Exists(row,col,ref fl))
            {
                added._flags.SetValue(row, col, fl);
            }

            var v = Copy._formulas.GetValue(row, col);
            if (v != null)
            {
                added.SetFormula(row, col, v);
            }
            var s = Copy._styles.GetValue(row, col);
            if (s != 0)
            {
                added._styles.SetValue(row, col, s);
            }
            var f = Copy._formulas.GetValue(row, col);
            if (f != null)
            {
                added._formulas.SetValue(row, col, f);
            }
            return s;
        }

        private void CopyComment(ExcelWorksheet Copy, ExcelWorksheet workSheet)
        {            
            //First copy the drawing XML
            string xml = Copy.Comments.CommentXml.InnerXml;
            var uriComment = new Uri(string.Format("/xl/comments{0}.xml", workSheet.SheetID), UriKind.Relative);
            if (_pck.Package.PartExists(uriComment))
            {
                uriComment = XmlHelper.GetNewUri(_pck.Package, "/xl/drawings/vmldrawing{0}.vml");
            }

            var part = _pck.Package.CreatePart(uriComment, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _pck.Compression);

            StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            //streamDrawing.Close();
            streamDrawing.Flush();

            //Add the relationship ID to the worksheet xml.
            var commentRelation = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri,uriComment), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");

            xml = Copy.VmlDrawingsComments.VmlDrawingXml.InnerXml;

            var uriVml = new Uri(string.Format("/xl/drawings/vmldrawing{0}.vml", workSheet.SheetID), UriKind.Relative);
            if (_pck.Package.PartExists(uriVml))
            {
                uriVml = XmlHelper.GetNewUri(_pck.Package, "/xl/drawings/vmldrawing{0}.vml");
            }

            var vmlPart = _pck.Package.CreatePart(uriVml, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
            StreamWriter streamVml = new StreamWriter(vmlPart.GetStream(FileMode.Create, FileAccess.Write));
            streamVml.Write(xml);
            //streamVml.Close();
            streamVml.Flush();

            var newVmlRel = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri,uriVml), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");

            //Add the relationship ID to the worksheet xml.
            XmlElement e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
            if (e == null)
            {
                workSheet.CreateNode("d:legacyDrawing");
                e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
            }

            e.SetAttribute("id", ExcelPackage.schemaRelationships, newVmlRel.Id);
        }
        private void CopyDrawing(ExcelWorksheet Copy, ExcelWorksheet workSheet/*, PackageRelationship r*/)
        {
            
            //Check if the worksheet has drawings
            //if(_xlPackage.Package.PartExists(r.TargetUri))
            //{
                //First copy the drawing XML                
                string xml = Copy.Drawings.DrawingXml.OuterXml;            
                var uriDraw=new Uri(string.Format("/xl/drawings/drawing{0}.xml", workSheet.SheetID),  UriKind.Relative);
                var part= _pck.Package.CreatePart(uriDraw,"application/vnd.openxmlformats-officedocument.drawing+xml", _pck.Compression);
                StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                streamDrawing.Write(xml);
                //streamDrawing.Close();
                streamDrawing.Flush();

                XmlDocument drawXml = new XmlDocument();
                drawXml.LoadXml(xml);
                //Add the relationship ID to the worksheet xml.
                var drawRelation = workSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri,uriDraw), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
                XmlElement e = workSheet.WorksheetXml.SelectSingleNode("//d:drawing", _namespaceManager) as XmlElement;
                e.SetAttribute("id",ExcelPackage.schemaRelationships, drawRelation.Id);

                foreach (ExcelDrawing draw in Copy.Drawings)
                {
                    if (draw is ExcelChart)
                    {
                        ExcelChart chart = draw as ExcelChart;
                        xml = chart.ChartXml.InnerXml;

                        var UriChart = XmlHelper.GetNewUri(_pck.Package, "/xl/charts/chart{0}.xml");
                        var chartPart = _pck.Package.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _pck.Compression);
                        StreamWriter streamChart = new StreamWriter(chartPart.GetStream(FileMode.Create, FileAccess.Write));
                        streamChart.Write(xml);
                        //streamChart.Close();
                        streamChart.Flush();
                        //Now create the new relationship to the copied chart xml
                        var prevRelID=draw.TopNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", Copy.Drawings.NameSpaceManager).Value;
                        var rel = part.CreateRelationship(UriHelper.GetRelativeUri(uriDraw,UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
                        XmlAttribute relAtt = drawXml.SelectSingleNode(string.Format("//c:chart/@r:id[.='{0}']", prevRelID), Copy.Drawings.NameSpaceManager) as XmlAttribute;
                        relAtt.Value=rel.Id;
                    }
                    else if (draw is ExcelPicture)
                    {
                        ExcelPicture pic = draw as ExcelPicture;
                        var uri = pic.UriPic;
                        if(!workSheet.Workbook._package.Package.PartExists(uri))
                        {
                            var picPart = workSheet.Workbook._package.Package.CreatePart(uri, pic.ContentType, CompressionLevel.None);
                            pic.Image.Save(picPart.GetStream(FileMode.Create, FileAccess.Write), pic.ImageFormat);
                        }

                        var prevRelID = draw.TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", Copy.Drawings.NameSpaceManager).Value;
                        var rel = part.CreateRelationship(UriHelper.GetRelativeUri(workSheet.WorksheetUri, uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                        XmlAttribute relAtt = drawXml.SelectSingleNode(string.Format("//xdr:pic/xdr:blipFill/a:blip/@r:embed[.='{0}']", prevRelID), Copy.Drawings.NameSpaceManager) as XmlAttribute;
                        relAtt.Value = rel.Id;
                    }
                }
                //rewrite the drawing xml with the new relID's
                streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                streamDrawing.Write(drawXml.OuterXml);
               // streamDrawing.Close();
                streamDrawing.Flush();

            //}
        }

		private void CopyVmlDrawing(ExcelWorksheet origSheet, ExcelWorksheet newSheet)
		{
			var xml = origSheet.VmlDrawingsComments.VmlDrawingXml.OuterXml;
			var vmlUri = new Uri(string.Format("/xl/drawings/vmlDrawing{0}.vml", newSheet.SheetID), UriKind.Relative);
			var part = _pck.Package.CreatePart(vmlUri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _pck.Compression);
			using (var streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write)))
			{
				streamDrawing.Write(xml);
                streamDrawing.Flush();
            }
			
            //Add the relationship ID to the worksheet xml.
			var vmlRelation = newSheet.Part.CreateRelationship(UriHelper.GetRelativeUri(newSheet.WorksheetUri,vmlUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
			var e = newSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _namespaceManager) as XmlElement;
			if (e == null)
			{
				e = newSheet.WorksheetXml.CreateNode(XmlNodeType.Entity, "//d:legacyDrawing", _namespaceManager.LookupNamespace("d")) as XmlElement;
			}
			if (e != null)
			{
				e.SetAttribute("id", ExcelPackage.schemaRelationships, vmlRelation.Id);
			}
		}

		string CreateWorkbookRel(string Name, int sheetID, Uri uriWorksheet, bool isChart)
        {
            //Create the relationship between the workbook and the new worksheet
            var rel = _pck.Workbook.Part.CreateRelationship(UriHelper.GetRelativeUri(_pck.Workbook.WorkbookUri, uriWorksheet), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/" + (isChart ? "chartsheet" : "worksheet"));
            _pck.Package.Flush();

            //Create the new sheet node
            XmlElement worksheetNode = _pck.Workbook.WorkbookXml.CreateElement("sheet", ExcelPackage.schemaMain);
            worksheetNode.SetAttribute("name", Name);
            worksheetNode.SetAttribute("sheetId", sheetID.ToString());
            worksheetNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);

            TopNode.AppendChild(worksheetNode);
            return rel.Id;
        }
        private void GetSheetURI(ref string Name, out int sheetID, out Uri uriWorksheet, bool isChart)
        {
            Name = ValidateFixSheetName(Name);

            //First find maximum existing sheetID
            sheetID = 0;
            foreach(var ws in this)
            {                
                if (ws.SheetID > sheetID)
                {
                    sheetID = ws.SheetID;
                }
            }
            // we now have the max existing values, so add one
            sheetID++;

            // add the new worksheet to the package
            if (isChart)
            {
                uriWorksheet = new Uri("/xl/chartsheets/chartsheet" + sheetID.ToString() + ".xml", UriKind.Relative);
            }
            else
            {
                uriWorksheet = new Uri("/xl/worksheets/sheet" + sheetID.ToString() + ".xml", UriKind.Relative);
            }
        }

        internal string ValidateFixSheetName(string Name)
        {
            //remove invalid characters
            if (ValidateName(Name))
            {
                if (Name.IndexOf(':') > -1) Name = Name.Replace(":", " ");
                if (Name.IndexOf('/') > -1) Name = Name.Replace("/", " ");
                if (Name.IndexOf('\\') > -1) Name = Name.Replace("\\", " ");
                if (Name.IndexOf('?') > -1) Name = Name.Replace("?", " ");
                if (Name.IndexOf('[') > -1) Name = Name.Replace("[", " ");
                if (Name.IndexOf(']') > -1) Name = Name.Replace("]", " ");
            }

            if (Name.Trim() == "")
            {
                throw new ArgumentException("The worksheet can not have an empty name");
            }
            if (Name.Length > 31) Name = Name.Substring(0, 31);   //A sheet can have max 31 char's            
            return Name;
        }
        /// <summary>
        /// Validate the sheetname
        /// </summary>
        /// <param name="Name">The Name</param>
        /// <returns>True if valid</returns>
        private bool ValidateName(string Name)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(Name, @":|\?|/|\\|\[|\]");
        }

		/// <summary>
		/// Creates the XML document representing a new empty worksheet
		/// </summary>
		/// <returns></returns>
		internal XmlDocument CreateNewWorksheet(bool isChart)
		{
			XmlDocument xmlDoc = new XmlDocument();
            XmlElement elemWs = xmlDoc.CreateElement(isChart ? "chartsheet" : "worksheet", ExcelPackage.schemaMain);
            elemWs.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);
            xmlDoc.AppendChild(elemWs);


            if (isChart)
            {
                XmlElement elemSheetPr = xmlDoc.CreateElement("sheetPr", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetPr);

                XmlElement elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetViews);

                XmlElement elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
                elemSheetView.SetAttribute("workbookViewId", "0");
                elemSheetView.SetAttribute("zoomToFit", "1");

                elemSheetViews.AppendChild(elemSheetView);
            }
            else
            {
                XmlElement elemSheetViews = xmlDoc.CreateElement("sheetViews", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetViews);

                XmlElement elemSheetView = xmlDoc.CreateElement("sheetView", ExcelPackage.schemaMain);
                elemSheetView.SetAttribute("workbookViewId", "0");
                elemSheetViews.AppendChild(elemSheetView);

                XmlElement elemSheetFormatPr = xmlDoc.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
                elemSheetFormatPr.SetAttribute("defaultRowHeight", "15");
                elemWs.AppendChild(elemSheetFormatPr);

                XmlElement elemSheetData = xmlDoc.CreateElement("sheetData", ExcelPackage.schemaMain);
                elemWs.AppendChild(elemSheetData);
            }
            return xmlDoc;
		}
		#endregion
		#region Delete Worksheet
		/// <summary>
		/// Deletes a worksheet from the collection
		/// </summary>
		/// <param name="Index">The position of the worksheet in the workbook</param>
		public void Delete(int Index)
		{
			/*
            * Hack to prefetch all the drawings,
            * so that all the images are referenced, 
            * to prevent the deletion of the image file, 
            * when referenced more than once
            */
            foreach (var ws in _worksheets)
            {
                var drawings = ws.Value.Drawings; 
            }			
            
            ExcelWorksheet worksheet = _worksheets[Index];
            if (worksheet.Drawings.Count > 0)
            {
                worksheet.Drawings.ClearDrawings();
            }

            //Remove all comments
            if (!(worksheet is ExcelChartsheet) && worksheet.Comments.Count > 0)
            {
                worksheet.Comments.Clear();
            }
                        
		    //Delete any parts still with relations to the Worksheet.
            DeleteRelationsAndParts(worksheet.Part);


            //Delete the worksheet part and relation from the package 
			_pck.Workbook.Part.DeleteRelationship(worksheet.RelationshipID);

            //Delete worksheet from the workbook XML
			XmlNode sheetsNode = _pck.Workbook.WorkbookXml.SelectSingleNode("//d:workbook/d:sheets", _namespaceManager);
			if (sheetsNode != null)
			{
				XmlNode sheetNode = sheetsNode.SelectSingleNode(string.Format("./d:sheet[@sheetId={0}]", worksheet.SheetID), _namespaceManager);
				if (sheetNode != null)
				{
					sheetsNode.RemoveChild(sheetNode);
				}
			}
			_worksheets.Remove(Index);
            if (_pck.Workbook.VbaProject != null)
            {
                _pck.Workbook.VbaProject.Modules.Remove(worksheet.CodeModule);
            }
			ReindexWorksheetDictionary();
            //If the active sheet is deleted, set the first tab as active.
            if (_pck.Workbook.View.ActiveTab >= _pck.Workbook.Worksheets.Count)
            {
                _pck.Workbook.View.ActiveTab = _pck.Workbook.View.ActiveTab-1;
            }
            if (_pck.Workbook.View.ActiveTab == worksheet.SheetID)
            {
                _pck.Workbook.Worksheets[1].View.TabSelected = true;
            }
            worksheet = null;
        }

        private void DeleteRelationsAndParts(Packaging.ZipPackagePart part)
        {
            var rels = part.GetRelationships().ToList();
            for(int i=0;i<rels.Count;i++)
            {
                var rel = rels[i];
                DeleteRelationsAndParts(_pck.Package.GetPart(UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri)));
                part.DeleteRelationship(rel.Id);
            }            
            _pck.Package.DeletePart(part.Uri);
        }

		/// <summary>
		/// Deletes a worksheet from the collection
		/// </summary>
		/// <param name="name">The name of the worksheet in the workbook</param>
		public void Delete(string name)
		{
			var sheet = this[name];
			if (sheet == null)
			{
				throw new ArgumentException(string.Format("Could not find worksheet to delete '{0}'", name));
			}
			Delete(sheet.PositionID);
		}
		/// <summary>
        /// Delete a worksheet from the collection
        /// </summary>
        /// <param name="Worksheet">The worksheet to delete</param>
        public void Delete(ExcelWorksheet Worksheet)
		{
            if (Worksheet.PositionID <= _worksheets.Count && Worksheet == _worksheets[Worksheet.PositionID])
            {
                Delete(Worksheet.PositionID);
            }
            else
            {
                throw (new ArgumentException("Worksheet is not in the collection."));
            }
        }
        #endregion
		private void ReindexWorksheetDictionary()
		{
			var index = 1;
			var worksheets = new Dictionary<int, ExcelWorksheet>();
			foreach (var entry in _worksheets)
			{
				entry.Value.PositionID = index;
				worksheets.Add(index++, entry.Value);
			}
			_worksheets = worksheets;
		}

		/// <summary>
		/// Returns the worksheet at the specified position.  
		/// </summary>
		/// <param name="PositionID">The position of the worksheet. 1-base</param>
		/// <returns></returns>
		public ExcelWorksheet this[int PositionID]
		{
			get
			{
			    if (_worksheets.ContainsKey(PositionID))
			    {
			        return _worksheets[PositionID];
			    }
			    else
			    {
			        throw (new IndexOutOfRangeException("Worksheet position out of range."));
			    }
			}
		}

		/// <summary>
		/// Returns the worksheet matching the specified name
		/// </summary>
		/// <param name="Name">The name of the worksheet</param>
		/// <returns></returns>
		public ExcelWorksheet this[string Name]
		{
			get
			{
                return GetByName(Name);
			}
		}
		/// <summary>
		/// Copies the named worksheet and creates a new worksheet in the same workbook
		/// </summary>
		/// <param name="Name">The name of the existing worksheet</param>
		/// <param name="NewName">The name of the new worksheet to create</param>
		/// <returns>The new copy added to the end of the worksheets collection</returns>
		public ExcelWorksheet Copy(string Name, string NewName)
		{
            ExcelWorksheet Copy = this[Name];
            if (Copy == null)
                throw new ArgumentException(string.Format("Copy worksheet error: Could not find worksheet to copy '{0}'", Name));

            ExcelWorksheet added = Add(NewName, Copy);
            return added;
        }
		#endregion
        internal ExcelWorksheet GetBySheetID(int localSheetID)
        {
            foreach (ExcelWorksheet ws in this)
            {
                if (ws.SheetID == localSheetID)
                {
                    return ws;
                }
            }
            return null;
        }
        private ExcelWorksheet GetByName(string Name)
        {
            if (string.IsNullOrEmpty(Name)) return null;
            ExcelWorksheet xlWorksheet = null;
            foreach (ExcelWorksheet worksheet in _worksheets.Values)
            {
                if (worksheet.Name.Equals(Name, StringComparison.InvariantCultureIgnoreCase))
                    xlWorksheet = worksheet;
            }
            return (xlWorksheet);
        }
		#region MoveBefore and MoveAfter Methods
		/// <summary>
		/// Moves the source worksheet to the position before the target worksheet
		/// </summary>
		/// <param name="sourceName">The name of the source worksheet</param>
		/// <param name="targetName">The name of the target worksheet</param>
		public void MoveBefore(string sourceName, string targetName)
		{
			Move(sourceName, targetName, false);
		}

		/// <summary>
		/// Moves the source worksheet to the position before the target worksheet
		/// </summary>
		/// <param name="sourcePositionId">The id of the source worksheet</param>
		/// <param name="targetPositionId">The id of the target worksheet</param>
		public void MoveBefore(int sourcePositionId, int targetPositionId)
		{
			Move(sourcePositionId, targetPositionId, false);
		}

		/// <summary>
		/// Moves the source worksheet to the position after the target worksheet
		/// </summary>
		/// <param name="sourceName">The name of the source worksheet</param>
		/// <param name="targetName">The name of the target worksheet</param>
		public void MoveAfter(string sourceName, string targetName)
		{
			Move(sourceName, targetName, true);
		}

		/// <summary>
		/// Moves the source worksheet to the position after the target worksheet
		/// </summary>
		/// <param name="sourcePositionId">The id of the source worksheet</param>
		/// <param name="targetPositionId">The id of the target worksheet</param>
		public void MoveAfter(int sourcePositionId, int targetPositionId)
		{
			Move(sourcePositionId, targetPositionId, true);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sourceName"></param>
		public void MoveToStart(string sourceName)
		{
			var sourceSheet = this[sourceName];
			if (sourceSheet == null)
			{
				throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
			}
			Move(sourceSheet.PositionID, 1, false);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sourcePositionId"></param>
		public void MoveToStart(int sourcePositionId)
		{
			Move(sourcePositionId, 1, false);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sourceName"></param>
		public void MoveToEnd(string sourceName)
		{
			var sourceSheet = this[sourceName];
			if (sourceSheet == null)
			{
				throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
			}
			Move(sourceSheet.PositionID, _worksheets.Count, true);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="sourcePositionId"></param>
		public void MoveToEnd(int sourcePositionId)
		{
			Move(sourcePositionId, _worksheets.Count, true);
		}

		private void Move(string sourceName, string targetName, bool placeAfter)
		{
			var sourceSheet = this[sourceName];
			if (sourceSheet == null)
			{
				throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", sourceName));
			}
			var targetSheet = this[targetName];
			if (targetSheet == null)
			{
				throw new Exception(string.Format("Move worksheet error: Could not find worksheet to move '{0}'", targetName));
			}
			Move(sourceSheet.PositionID, targetSheet.PositionID, placeAfter);
		}

		private void Move(int sourcePositionId, int targetPositionId, bool placeAfter)
		{
            // Bugfix: if source and target are the same worksheet the following code will create a duplicate
            //         which will cause a corrupt workbook. /swmal 2014-05-10
		    if (sourcePositionId == targetPositionId) return;

            lock (_worksheets)
            {
                var sourceSheet = this[sourcePositionId];
                if (sourceSheet == null)
                {
                    throw new Exception(string.Format("Move worksheet error: Could not find worksheet at position '{0}'", sourcePositionId));
                }
                var targetSheet = this[targetPositionId];
                if (targetSheet == null)
                {
                    throw new Exception(string.Format("Move worksheet error: Could not find worksheet at position '{0}'", targetPositionId));
                }
                if (sourcePositionId == targetPositionId && _worksheets.Count < 2)
                {
                    return;		//--- no reason to attempt to re-arrange a single item with itself
                }

                var index = 1;
                var newOrder = new Dictionary<int, ExcelWorksheet>();
                foreach (var entry in _worksheets)
                {
                    if (entry.Key == targetPositionId)
                    {
                        if (!placeAfter)
                        {
                            sourceSheet.PositionID = index;
                            newOrder.Add(index++, sourceSheet);
                        }

                        entry.Value.PositionID = index;
                        newOrder.Add(index++, entry.Value);

                        if (placeAfter)
                        {
                            sourceSheet.PositionID = index;
                            newOrder.Add(index++, sourceSheet);
                        }
                    }
                    else if (entry.Key == sourcePositionId)
                    {
                        //--- do nothing
                    }
                    else
                    {
                        entry.Value.PositionID = index;
                        newOrder.Add(index++, entry.Value);
                    }
                }
                _worksheets = newOrder;

                MoveSheetXmlNode(sourceSheet, targetSheet, placeAfter);
            }
		}

		private void MoveSheetXmlNode(ExcelWorksheet sourceSheet, ExcelWorksheet targetSheet, bool placeAfter)
		{
            lock (TopNode.OwnerDocument)
            {
                var sourceNode = TopNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", sourceSheet.SheetID), _namespaceManager);
                var targetNode = TopNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", targetSheet.SheetID), _namespaceManager);
                if (sourceNode == null || targetNode == null)
                {
                    throw new Exception("Source SheetId and Target SheetId must be valid");
                }
                if (placeAfter)
                {
                    TopNode.InsertAfter(sourceNode, targetNode);
                }
                else
                {
                    TopNode.InsertBefore(sourceNode, targetNode);
                }
            }
		}

		#endregion
        public void Dispose()
        {            
             foreach (var sheet in this._worksheets.Values) 
             { 
                 ((IDisposable)sheet).Dispose(); 
             } 
             _worksheets = null;
             _pck = null;            
        }
    } // end class Worksheets
}
