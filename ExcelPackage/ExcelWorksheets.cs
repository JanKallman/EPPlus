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
            //Copy worksheet XML
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

            //Copy all relationships 
            //CopyRelationShips(Copy, added);
            if (Copy.Drawings.Count > 0)
            {
                CopyDrawing(Copy, added);
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

        private void CloneCells(ExcelWorksheet Copy, ExcelWorksheet added)
        {
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
            //Cells
            foreach (ExcelCell cell in Copy._cells)
            {                
                added._cells.Add(cell.Clone(added));
            }
            //Rows
            foreach (ExcelRow row in Copy._rows)
            {
                row.Clone(added);
            }
            //Columns
            foreach (ExcelColumn col in Copy._columns)
            {
                col.Clone(added);
            }
        }

        private void CopyRelationShips(ExcelWorksheet Copy, ExcelWorksheet workSheet)
        {
            foreach (var r in Copy.Part.GetRelationships())
            {
                switch (r.RelationshipType)
                {
                    case ExcelPackage.schemaRelationships + "/drawing":
                        //CopyDrawing(Copy, workSheet, r);
                        break;
                    case ExcelPackage.schemaHyperlink:
                        //Do nothing. Hyperlinks are handled in memory.
                        break;
                    case ExcelPackage.schemaComment:
                        break;
                    case ExcelPackage.schemaImage:
                        //
                        break;
                    default:    //Other rels are not copied
                        break;
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
            PackageRelationship commentRelation = workSheet.Part.CreateRelationship(uriComment, TargetMode.Internal, ExcelPackage.schemaRelationships + "/comments");

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

            PackageRelationship newVmlRel = workSheet.Part.CreateRelationship(uriVml, TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");

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
                PackageRelationship drawRelation = workSheet.Part.CreateRelationship(uriDraw, TargetMode.Internal, ExcelPackage.schemaRelationships + "/drawing");
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
                        var rel = part.CreateRelationship(UriChart, TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
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
                        var rel = part.CreateRelationship(uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
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
        string CreateWorkbookRel(string Name, int sheetID, Uri uriWorksheet)
        {
            // create the relationship between the workbook and the new worksheet
            PackageRelationship rel = _xlPackage.Workbook.Part.CreateRelationship(uriWorksheet, TargetMode.Internal, ExcelPackage.schemaRelationships + "/worksheet");
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
                Name = Name.Remove(0, ':');
                Name = Name.Remove(0, '/');
                Name = Name.Remove(0, '\\');
                Name = Name.Remove(0, '?');
                Name = Name.Remove(0, '[');
                Name = Name.Remove(0, ']');
            }

            if (Name.Trim() == "")
            {
                throw new Exception("Add worksheet Error: attempting to create worksheet with an empty name");
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
                        throw new Exception("Add worksheet Error: attempting to create worksheet with duplicate name");
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
		/// <param name="positionID">The position of the worksheet in the workbook</param>
		public void Delete(int positionID)
		{
			if (_worksheets.Count == 1)
				throw new Exception("Error: You are attempting to delete the last worksheet in the workbook.  One worksheet MUST be present in the workbook!");
			ExcelWorksheet worksheet = _worksheets[positionID];

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
			_worksheets.Remove(positionID);
		}
		#endregion

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
				ExcelWorksheet xlWorksheet = null;
				foreach (ExcelWorksheet worksheet in _worksheets.Values)
				{
                    if (worksheet.Name.ToLower() == Name.ToLower())
						xlWorksheet = worksheet;
				}
				return (xlWorksheet);
				//throw new Exception(string.Format("ExcelWorksheets Error: Worksheet '{0}' not found!",Name));
			}
		}

		/// <summary>
		/// Copies the named worksheet and creates a new worksheet in the same workbook
		/// </summary>
		/// <param name="Name">The name of the existing worksheet</param>
		/// <param name="NewName">The name of the new worksheet to create</param>
		/// <returns></returns>
		public ExcelWorksheet Copy(string Name, string NewName)
		{
			// TODO: implement copy worksheet
			throw new Exception("The method or operation is not implemented.");
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

    } // end class Worksheets
}

