/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
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
 * Parts of the interface of this file comes from the Excelpackage project. http://www.codeplex.com/ExcelPackage
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.IO.Packaging;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style.XmlAccess;
namespace OfficeOpenXml
{
	/// <summary>
	/// Provides enumeration through all the worksheets in the workbook
	/// </summary>
	public class ExcelWorksheets : IEnumerable<ExcelWorksheet>
	{
		#region ExcelWorksheets Private Properties
		private Dictionary<int, ExcelWorksheet> _worksheets;
		private ExcelPackage _xlPackage;
		private XmlNamespaceManager _nsManager;
		private XmlNode _worksheetsNode;
		#endregion

		#region ExcelWorksheets Constructor
		/// <summary>
		/// Creates a new instance of the ExcelWorksheets class.
		/// For internal use only!
		/// </summary>
		/// <param name="xlPackage"></param>
		protected internal ExcelWorksheets(ExcelPackage xlPackage)
		{
			_xlPackage = xlPackage;
            
			//  Create a NamespaceManager to handle the default namespace, 
			//  and create a prefix for the default namespace:
			NameTable nt = new NameTable();
			_nsManager = new XmlNamespaceManager(nt);
            _nsManager.AddNamespace(string.Empty, ExcelPackage.schemaMain);
            _nsManager.AddNamespace("d", ExcelPackage.schemaMain);
			_nsManager.AddNamespace("r", ExcelPackage.schemaRelationships);
            _nsManager.AddNamespace("c", ExcelPackage.schemaChart);

			// obtain container node for all worksheets
			_worksheetsNode = xlPackage.Workbook.WorkbookXml.SelectSingleNode("//d:sheets", _nsManager);
			if (_worksheetsNode == null)
			{
				// create new node as it did not exist
				_worksheetsNode = _xlPackage.Workbook.WorkbookXml.CreateElement("sheets", ExcelPackage.schemaMain);
				xlPackage.Workbook.WorkbookXml.DocumentElement.AppendChild(_worksheetsNode);
			}

			_worksheets = new Dictionary<int, ExcelWorksheet>();
			int positionID = 1;
			foreach (XmlNode sheetNode in _worksheetsNode.ChildNodes)
			{
				string name = sheetNode.Attributes["name"].Value;
				//  Get the relationship id attribute:
				string relId = sheetNode.Attributes["r:id"].Value;
				int sheetID = Convert.ToInt32(sheetNode.Attributes["sheetId"].Value);
				// get hidden attribute (if present)
				eWorkSheetHidden hidden = eWorkSheetHidden.Visible;
				XmlNode attr = sheetNode.Attributes["state"];
				if (attr != null)
					hidden = TranslateHidden(attr.Value);

				//string type = "";
				//attr = sheetNode.Attributes["type"];
				//if (attr != null)
				//  type = attr.Value;

				PackageRelationship sheetRelation = xlPackage.Workbook.Part.GetRelationship(relId);
				Uri uriWorksheet = PackUriHelper.ResolvePartUri(xlPackage.Workbook.WorkbookUri, sheetRelation.TargetUri);
				
				// add worksheet to our collection
                _worksheets.Add(positionID, new ExcelWorksheet(_nsManager, _xlPackage, relId, uriWorksheet, name, sheetID, positionID, hidden));
				positionID++;
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

		#region ExcelWorksheets Public Methods
		/// <summary>
		/// Returns an enumerator that allows the foreach syntax to be used to 
		/// itterate through all the worksheets
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
		/// Adds a blank worksheet with the desired name
		/// </summary>
		/// <param name="Name"></param>
		public ExcelWorksheet Add(string Name)
		{
            int sheetID;
            Uri uriWorksheet;
            GetSheetURI(ref Name, out sheetID, out uriWorksheet);
            PackagePart worksheetPart = _xlPackage.Package.CreatePart(uriWorksheet, @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", _xlPackage.Compression);

			// create the new, empty worksheet and save it to the package
			StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
			XmlDocument worksheetXml = CreateNewWorksheet();
			worksheetXml.Save(streamWorksheet);
			streamWorksheet.Close();
			_xlPackage.Package.Flush();

            string rel = CreateWorkbookRel(Name, sheetID, uriWorksheet);

			// create a reference to the new worksheet in our collection
			int positionID = _worksheets.Count + 1;
            ExcelWorksheet worksheet = new ExcelWorksheet(_nsManager, _xlPackage, rel, uriWorksheet, Name, sheetID, positionID, eWorkSheetHidden.Visible);

			_worksheets.Add(positionID, worksheet);
			return worksheet;
		}
        /// <summary>
        /// Adds a copy of a worksheet with the desired name
        /// </summary>
        /// <param name="Name"></param>
        /// <param name="Copy">The worksheet to be copied</param>
        public ExcelWorksheet Add(string Name, ExcelWorksheet Copy)
        {
            int sheetID;
            Uri uriWorksheet;

            GetSheetURI(ref Name, out sheetID, out uriWorksheet);

            //Create a copy of the worksheet XML
            PackagePart worksheetPart = _xlPackage.Package.CreatePart(uriWorksheet, @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", _xlPackage.Compression);
            StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
            XmlDocument worksheetXml = new XmlDocument();
            worksheetXml.LoadXml(Copy.WorksheetXml.OuterXml);
            worksheetXml.Save(streamWorksheet);
            streamWorksheet.Close();
            _xlPackage.Package.Flush();


            //Create a relation to the workbook
            string relID = CreateWorkbookRel(Name, sheetID, uriWorksheet);
            ExcelWorksheet added = new ExcelWorksheet(_nsManager, _xlPackage, relID, uriWorksheet, Name, sheetID, _worksheets.Count + 1, eWorkSheetHidden.Visible);

            //Copy comments
            if (Copy.Comments.Count > 0)
            {
                CopyComment(Copy, added);
            }
            else if (Copy.VmlDrawings.Count > 0)    //Vml drawings are copied as part of the comments. 
            {
                CopyVmlDrawing(Copy, added);
            }

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

            //Copy all cells
            CloneCells(Copy, added);

            _worksheets.Add(_worksheets.Count + 1, added);
            
            //Remove any relation to printersettings.
            XmlNode pageSetup = added.WorksheetXml.SelectSingleNode("//d:pageSetup", _nsManager);
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

        private void CopyTable(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            //First copy the table XML
            foreach (var tbl in Copy.Tables)
            {
                string xml=tbl.TableXml.OuterXml;
                int Id = _xlPackage.Workbook._nextTableID++;
                string name = Copy.Tables.GetNewTableName();
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xml);
                xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
                xmlDoc.SelectSingleNode("//d:table/@name", tbl.NameSpaceManager).Value = name;
                xmlDoc.SelectSingleNode("//d:table/@displayName", tbl.NameSpaceManager).Value = name;
                xml = xmlDoc.OuterXml;

                var uriTbl = new Uri(string.Format("/xl/tables/table{0}.xml", Id), UriKind.Relative);
                var part = _xlPackage.Package.CreatePart(uriTbl, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", _xlPackage.Compression);
                StreamWriter streamTbl = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                streamTbl.Write(xml);
                streamTbl.Close();

                //create the relationship and add the ID to the worksheet xml.
                var rel = added.Part.CreateRelationship(PackUriHelper.GetRelativeUri(added.WorksheetUri,uriTbl), TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");

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
            foreach (var tbl in Copy.PivotTables)
            {
                string xml = tbl.PivotTableXml.OuterXml;
                int Id = _xlPackage.Workbook._nextPivotTableID++;
                string name = Copy.PivotTables.GetNewTableName();
                XmlDocument xmlDoc = new XmlDocument();
                Copy.Save();    //Save the worksheet first
                xmlDoc.LoadXml(xml);
                //xmlDoc.SelectSingleNode("//d:table/@id", tbl.NameSpaceManager).Value = Id.ToString();
                xmlDoc.SelectSingleNode("//d:pivotTableDefinition/@name", tbl.NameSpaceManager).Value = name;
                xml = xmlDoc.OuterXml;

                var uriTbl = new Uri(string.Format("/xl/pivotTables/pivotTable{0}.xml", Id), UriKind.Relative);
                var part = _xlPackage.Package.CreatePart(uriTbl, ExcelPackage.schemaPivotTable , _xlPackage.Compression);
                StreamWriter streamTbl = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                streamTbl.Write(xml);
                streamTbl.Close();

                //create the relationship and add the ID to the worksheet xml.
                added.Part.CreateRelationship(PackUriHelper.ResolvePartUri(added.WorksheetUri, uriTbl), TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");
                part.CreateRelationship(PackUriHelper.ResolvePartUri(tbl.Relationship.SourceUri, tbl.CacheDefinition.Relationship.TargetUri), tbl.CacheDefinition.Relationship.TargetMode, tbl.CacheDefinition.Relationship.RelationshipType);
            }
        }
        private void CloneCells(ExcelWorksheet Copy, ExcelWorksheet added)
        {
            bool sameWorkbook=(Copy.Workbook == _xlPackage.Workbook);
            
            added.MergedCells.List.AddRange(Copy.MergedCells.List);
            //Formulas
            foreach (IRangeID f in Copy._formulaCells)
            {
                added._formulaCells.Add(f);
            }
            //Shared Forulas
            foreach (int key in Copy._sharedFormulas.Keys)
            {
                added._sharedFormulas.Add(key, Copy._sharedFormulas[key]);
            }
            
            Dictionary<int, int> styleCashe = new Dictionary<int, int>();
            //Cells
            foreach (ExcelCell cell in Copy._cells)
            {                
                if (sameWorkbook)   //Same workbook == same styles
                {
                    added._cells.Add(cell.Clone(added));
                }
                else
                {
                    ExcelCell addedCell=cell.Clone(added);
                    if (styleCashe.ContainsKey(cell.StyleID))
                    {
                        addedCell.StyleID = styleCashe[cell.StyleID];
                    }
                    else
                    {
                        addedCell.StyleID = added.Workbook.Styles.CloneStyle(Copy.Workbook.Styles,  cell.StyleID);
                        if (cell.StyleName != "") //Named styles
                        {
                            if (!Copy.Workbook.Styles.NamedStyles.ExistsKey(cell.StyleName))
                            {
                               var ns=Copy.Workbook.Styles.CreateNamedStyle(cell.StyleName);
                               ns.StyleXfId  = addedCell.StyleID;
                            }
                            
                        }
                        styleCashe.Add(cell.StyleID, addedCell.StyleID);
                    }
                    added._cells.Add(addedCell);
                }
            }
            //Rows
            foreach (ExcelRow row in Copy._rows)
            {
                row.Clone(added);
                if (!sameWorkbook)   //Same workbook == same styles
                {
                    ExcelRow addedRow = added.Row(row.Row) as ExcelRow;
                    if (styleCashe.ContainsKey(row.StyleID))
                    {
                        addedRow.StyleID = styleCashe[row.StyleID];
                    }
                    else
                    {
                        addedRow.StyleID = added.Workbook.Styles.CloneStyle(Copy.Workbook.Styles, addedRow.StyleID);
                        if (row.StyleName != "") //Named styles
                        {
                            if (!Copy.Workbook.Styles.NamedStyles.ExistsKey(row.StyleName))
                            {
                                var ns = Copy.Workbook.Styles.CreateNamedStyle(row.StyleName);
                                ns.StyleXfId = addedRow.StyleID;
                            }

                        }
                        styleCashe.Add(row.StyleID, addedRow.StyleID);
                    }
                }                
            }
            //Columns
            foreach (ExcelColumn col in Copy._columns)
            {
                col.Clone(added);
                if (!sameWorkbook)   //Same workbook == same styles
                {
                    ExcelColumn addedCol = added.Column(col.ColumnMin) as ExcelColumn;
                    if (styleCashe.ContainsKey(col.StyleID))
                    {
                        addedCol.StyleID = styleCashe[col.StyleID];
                    }
                    else
                    {
                        addedCol.StyleID = added.Workbook.Styles.CloneStyle(Copy.Workbook.Styles, addedCol.StyleID);
                        if (col.StyleName != "") //Named styles
                        {
                            if (!Copy.Workbook.Styles.NamedStyles.ExistsKey(col.StyleName))
                            {
                                var ns = Copy.Workbook.Styles.CreateNamedStyle(col.StyleName);
                                ns.StyleXfId = addedCol.StyleID;
                            }

                        }
                        styleCashe.Add(col.StyleID, addedCol.StyleID);
                    }
                }
            }
        }
        private void CopyComment(ExcelWorksheet Copy, ExcelWorksheet workSheet)
        {
            //First copy the drawing XML
            string xml = Copy.Comments.CommentXml.InnerXml;
            var uriComment = new Uri(string.Format("/xl/comments{0}.xml", workSheet.SheetID), UriKind.Relative);
            if (_xlPackage.Package.PartExists(uriComment))
            {
                uriComment = Copy.GetNewUri(_xlPackage.Package, "/xl/drawings/vmldrawing{0}.vml");
            }

            var part = _xlPackage.Package.CreatePart(uriComment, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _xlPackage.Compression);

            StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
            streamDrawing.Write(xml);
            streamDrawing.Close();

            //Add the relationship ID to the worksheet xml.
            PackageRelationship commentRelation = workSheet.Part.CreateRelationship(PackUriHelper.GetRelativeUri(workSheet.WorksheetUri,uriComment), TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");

            xml = Copy.VmlDrawings.VmlDrawingXml.InnerXml;

            var uriVml = new Uri(string.Format("/xl/drawings/vmldrawing{0}.vml", workSheet.SheetID), UriKind.Relative);
            if (_xlPackage.Package.PartExists(uriVml))
            {
                uriVml = Copy.GetNewUri(_xlPackage.Package, "/xl/drawings/vmldrawing{0}.vml");
            }

            var vmlPart = _xlPackage.Package.CreatePart(uriVml, "application/vnd.openxmlformats-officedocument.vmlDrawing", _xlPackage.Compression);
            StreamWriter streamVml = new StreamWriter(vmlPart.GetStream(FileMode.Create, FileAccess.Write));
            streamVml.Write(xml);
            streamVml.Close();

            PackageRelationship newVmlRel = workSheet.Part.CreateRelationship(PackUriHelper.GetRelativeUri(workSheet.WorksheetUri,uriVml), TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");

            //Add the relationship ID to the worksheet xml.
            XmlElement e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _nsManager) as XmlElement;
            if (e == null)
            {
                workSheet.CreateNode("d:legacyDrawing");
                e = workSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _nsManager) as XmlElement;
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
                var part= _xlPackage.Package.CreatePart(uriDraw,"application/vnd.openxmlformats-officedocument.drawing+xml", _xlPackage.Compression);
                StreamWriter streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                streamDrawing.Write(xml);
                streamDrawing.Close();

                XmlDocument drawXml = new XmlDocument();
                drawXml.LoadXml(xml);
                //Add the relationship ID to the worksheet xml.
                PackageRelationship drawRelation = workSheet.Part.CreateRelationship(PackUriHelper.GetRelativeUri(workSheet.WorksheetUri,uriDraw), TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
                XmlElement e = workSheet.WorksheetXml.SelectSingleNode("//d:drawing", _nsManager) as XmlElement;
                e.SetAttribute("id",ExcelPackage.schemaRelationships, drawRelation.Id);

                foreach (ExcelDrawing draw in Copy.Drawings)
                {
                    if (draw is ExcelChart)
                    {
                        ExcelChart chart = draw as ExcelChart;
                        xml = chart.ChartXml.InnerXml;

                        var UriChart = chart.GetNewUri(_xlPackage.Package, "/xl/charts/chart{0}.xml");
                        var chartPart = _xlPackage.Package.CreatePart(UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", _xlPackage.Compression);
                        StreamWriter streamChart = new StreamWriter(chartPart.GetStream(FileMode.Create, FileAccess.Write));
                        streamChart.Write(xml);
                        streamChart.Close();
                        //Now create the new relationship to the copied chart xml
                        var prevRelID=draw.TopNode.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart/@r:id", Copy.Drawings.NameSpaceManager).Value;
                        var rel = part.CreateRelationship(PackUriHelper.GetRelativeUri(uriDraw,UriChart), TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
                        XmlAttribute relAtt = drawXml.SelectSingleNode(string.Format("//c:chart/@r:id[.='{0}']", prevRelID), Copy.Drawings.NameSpaceManager) as XmlAttribute;
                        relAtt.Value=rel.Id;
                    }
                    else if (draw is ExcelPicture)
                    {
                        ExcelPicture pic = draw as ExcelPicture;
                        var uri = pic.UriPic;
                        if(!workSheet.Workbook._package.Package.PartExists(uri))
                        {
                            var picPart = workSheet.Workbook._package.Package.CreatePart(uri, pic.ContentType, CompressionOption.NotCompressed);
                            pic.Image.Save(picPart.GetStream(FileMode.Create, FileAccess.Write), pic.ImageFormat);
                        }

                        var prevRelID = draw.TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", Copy.Drawings.NameSpaceManager).Value;
                        var rel = part.CreateRelationship(PackUriHelper.GetRelativeUri(workSheet.WorksheetUri, uri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                        XmlAttribute relAtt = drawXml.SelectSingleNode(string.Format("//xdr:pic/xdr:blipFill/a:blip/@r:embed[.='{0}']", prevRelID), Copy.Drawings.NameSpaceManager) as XmlAttribute;
                        relAtt.Value = rel.Id;
                    }
                }
                //rewrite the drawing xml with the new relID's
                streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                streamDrawing.Write(drawXml.OuterXml);
                streamDrawing.Close();
            //}
        }

		private void CopyVmlDrawing(ExcelWorksheet origSheet, ExcelWorksheet newSheet)
		{
			var xml = origSheet.VmlDrawings.VmlDrawingXml.OuterXml;
			var vmlUri = new Uri(string.Format("/xl/drawings/vmlDrawing{0}.vml", newSheet.SheetID), UriKind.Relative);
			var part = _xlPackage.Package.CreatePart(vmlUri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _xlPackage.Compression);
			using (var streamDrawing = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write)))
			{
				streamDrawing.Write(xml);
			}

			//Add the relationship ID to the worksheet xml.
			PackageRelationship vmlRelation = newSheet.Part.CreateRelationship(PackUriHelper.GetRelativeUri(newSheet.WorksheetUri,vmlUri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
			var e = newSheet.WorksheetXml.SelectSingleNode("//d:legacyDrawing", _nsManager) as XmlElement;
			if (e == null)
			{
				e = newSheet.WorksheetXml.CreateNode(XmlNodeType.Entity, "//d:legacyDrawing", _nsManager.LookupNamespace("d")) as XmlElement;
			}
			if (e != null)
			{
				e.SetAttribute("id", ExcelPackage.schemaRelationships, vmlRelation.Id);
			}
		}

		string CreateWorkbookRel(string Name, int sheetID, Uri uriWorksheet)
        {
            // create the relationship between the workbook and the new worksheet
            PackageRelationship rel = _xlPackage.Workbook.Part.CreateRelationship(PackUriHelper.GetRelativeUri(_xlPackage.Workbook.WorkbookUri, uriWorksheet), TargetMode.Internal, ExcelPackage.schemaRelationships + "/worksheet");
            _xlPackage.Package.Flush();

            // now create the new worksheet tag and set name/SheetId attributes in the workbook.xml
            XmlElement worksheetNode = _xlPackage.Workbook.WorkbookXml.CreateElement("sheet", ExcelPackage.schemaMain);
            // create the new sheet node
            worksheetNode.SetAttribute("name", Name);
            worksheetNode.SetAttribute("sheetId", sheetID.ToString());
            // set the r:id attribute
            worksheetNode.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);
            // insert the sheet tag with all attributes set as above
            _worksheetsNode.AppendChild(worksheetNode);
            return rel.Id;
        }
        private void GetSheetURI(ref string Name, out int sheetID, out Uri uriWorksheet)
        {
            //remove invalid characters
            if (ValidateName(Name))
            {
                if (Name.IndexOf(':') > -1) Name = Name.Replace(":"," ");
                if (Name.IndexOf('/') > -1) Name = Name.Replace("/"," ");
                if (Name.IndexOf('\\') > -1) Name = Name.Replace("\\"," ");
                if (Name.IndexOf('?') > -1) Name = Name.Replace("?"," ");
                if (Name.IndexOf('[') > -1) Name = Name.Replace("["," ");
                if (Name.IndexOf(']') > -1) Name = Name.Replace("]"," ");
            }

            if (Name.Trim() == "")
            {
                throw new ArgumentException("Add worksheet Error: attempting to create worksheet with an empty name");
            }
            if (Name.Length > 31) Name = Name.Substring(0, 31);   //A sheet can have max 31 char's

            // first find maximum existing sheetID
            // also check the name is unique - if not throw an error
            sheetID = 0;
            foreach (XmlNode sheet in _worksheetsNode.ChildNodes)
            {
                XmlAttribute attr = (XmlAttribute)sheet.Attributes.GetNamedItem("sheetId");
                if (attr != null)
                {
                    int curID = int.Parse(attr.Value);
                    if (curID > sheetID)
                        sheetID = curID;
                }
                attr = (XmlAttribute)sheet.Attributes.GetNamedItem("name");
                if (attr != null)
                {
                    if (attr.Value == Name)
                        throw new ArgumentException("Add worksheet Error: attempting to create worksheet with duplicate name");
                }
            }
            // we now have the max existing values, so add one
            sheetID++;

            // add the new worksheet to the package
            uriWorksheet = new Uri("/xl/worksheets/sheet" + sheetID.ToString() + ".xml", UriKind.Relative);
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
		protected internal XmlDocument CreateNewWorksheet()
		{
			// create the new worksheet
			XmlDocument worksheetXml = new XmlDocument();
			// XML document does not exist so create the new worksheet XML doc
			XmlElement worksheetNode = worksheetXml.CreateElement("worksheet", ExcelPackage.schemaMain);
			worksheetNode.SetAttribute("xmlns:r", ExcelPackage.schemaRelationships);
			worksheetXml.AppendChild(worksheetNode);
			// create the sheetViews tag
			XmlElement tagSheetViews = worksheetXml.CreateElement("sheetViews", ExcelPackage.schemaMain);
			worksheetNode.AppendChild(tagSheetViews);
			// create the sheet View tag
			XmlElement tagSheetView = worksheetXml.CreateElement("sheetView", ExcelPackage.schemaMain);
			tagSheetView.SetAttribute("workbookViewId", "0");
			tagSheetViews.AppendChild(tagSheetView);

            // create the empty sheetData tag (must be present, but can be empty)
			XmlElement tagSheetData = worksheetXml.CreateElement("sheetData", ExcelPackage.schemaMain);
			worksheetNode.AppendChild(tagSheetData);
			return worksheetXml;
		}
		#endregion

		#region Delete Worksheet
		/// <summary>
		/// Delete a worksheet from the workbook package
		/// </summary>
		/// <param name="Index">The position of the worksheet in the workbook</param>
		public void Delete(int Index)
		{
			if (_worksheets.Count == 1)
				throw new Exception("Error: You are attempting to delete the last worksheet in the workbook.  One worksheet MUST be present in the workbook!");
			ExcelWorksheet worksheet = _worksheets[Index];

			// delete the worksheet from the package 
			_xlPackage.Package.DeletePart(worksheet.WorksheetUri);

			// delete the relationship from the package 
			_xlPackage.Workbook.Part.DeleteRelationship(worksheet.RelationshipID);

			// delete worksheet from the workbook XML
			XmlNode sheetsNode = _xlPackage.Workbook.WorkbookXml.SelectSingleNode("//d:workbook/d:sheets", _nsManager);
			if (sheetsNode != null)
			{
				XmlNode sheetNode = sheetsNode.SelectSingleNode(string.Format("./d:sheet[@sheetId={0}]", worksheet.SheetID), _nsManager);
				if (sheetNode != null)
				{
					sheetsNode.RemoveChild(sheetNode);
				}
			}
			// delete worksheet from the Dictionary object
			_worksheets.Remove(Index);

			ReindexWorksheetDictionary();
		}

		/// <summary>
		/// Delete a worksheet from the workbook package
		/// </summary>
		/// <param name="name">The name of the worksheet in the workbook</param>
		public void Delete(string name)
		{
			var sheet = this[name];
			if (sheet == null)
			{
				throw new Exception(string.Format("Could not find worksheet to delete '{0}'", name));
			}
			Delete(sheet.PositionID);
		}

		/// <summary>
        /// Delete a worksheet from the workbook
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
                return (_worksheets[PositionID]);
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
                if (string.IsNullOrEmpty(Name)) return null;
                ExcelWorksheet xlWorksheet = null;
				foreach (ExcelWorksheet worksheet in _worksheets.Values)
				{
                    if (worksheet.Name.ToLower() == Name.ToLower())
						xlWorksheet = worksheet;
				}
				return (xlWorksheet);
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
			if (_worksheets.Count < 2)
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

		private void MoveSheetXmlNode(ExcelWorksheet sourceSheet, ExcelWorksheet targetSheet, bool placeAfter)
		{
			var sourceNode = _worksheetsNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", sourceSheet.SheetID), _nsManager);
			var targetNode = _worksheetsNode.SelectSingleNode(string.Format("d:sheet[@sheetId = '{0}']", targetSheet.SheetID), _nsManager);
			if (sourceNode == null || targetNode == null)
			{
				throw new Exception("Source SheetId and Target SheetId must be valid");
			}
			if (placeAfter)
			{
				_worksheetsNode.InsertAfter(sourceNode, targetNode);
			}
			else
			{
				_worksheetsNode.InsertBefore(sourceNode, targetNode);
			}
		}

		#endregion
	} // end class Worksheets
}

