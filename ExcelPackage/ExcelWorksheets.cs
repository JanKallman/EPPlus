/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
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
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/

/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * ExcelPackage provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/ExcelPackage for details.
 * 
 * Copyright 2007 © Dr John Tunnicliffe 
 * mailto:dr.john.tunnicliffe@btinternet.com
 * All rights reserved.
 * 
 * ExcelPackage is an Open Source project provided under the 
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
 */

/*
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * John Tunnicliffe		Initial Release		01-Jan-2007
 * ******************************************************************************
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.IO.Packaging;
using OfficeOpenXml.Style;
namespace OfficeOpenXml
{
	/// <summary>
	/// Provides enumeration through all the worksheets in the workbook
	/// </summary>
	public class ExcelWorksheets : IEnumerable
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
            _nsManager.AddNamespace("d", ExcelPackage.schemaMain);
			_nsManager.AddNamespace("r", ExcelPackage.schemaRelationships);

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
				//if (sheetID != count)
				//{
				//  // renumber the sheets as they are in an odd order
				//  sheetID = count;
				//  sheetNode.Attributes["sheetId"].Value = sheetID.ToString();
				//}
				// get hidden attribute (if present)
				bool hidden = false;
				XmlNode attr = sheetNode.Attributes["hidden"];
				if (attr != null)
					hidden = Convert.ToBoolean(attr.Value);

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
		public IEnumerator GetEnumerator()
		{
			return (_worksheets.Values.GetEnumerator());
		    }

		#region Add Worksheet
		/// <summary>
		/// Adds a blank worksheet with the desired name
		/// </summary>
		/// <param name="Name"></param>
		public ExcelWorksheet Add(string Name)
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
			int sheetID = 0;
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
			Uri uriWorksheet = new Uri("/xl/worksheets/sheet" + sheetID.ToString() + ".xml", UriKind.Relative);
            PackagePart worksheetPart = _xlPackage.Package.CreatePart(uriWorksheet, @"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", CompressionOption.Maximum);

			// create the new, empty worksheet and save it to the package
			StreamWriter streamWorksheet = new StreamWriter(worksheetPart.GetStream(FileMode.Create, FileAccess.Write));
			XmlDocument worksheetXml = CreateNewWorksheet();
			worksheetXml.Save(streamWorksheet);
			streamWorksheet.Close();
			_xlPackage.Package.Flush();
			
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

			// create a reference to the new worksheet in our collection
			int positionID = _worksheets.Count + 1;
            ExcelWorksheet worksheet = new ExcelWorksheet(_nsManager, _xlPackage, rel.Id, uriWorksheet, Name, sheetID, positionID, false);

			_worksheets.Add(positionID, worksheet);
			return worksheet;
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
					if (worksheet.Name == Name)
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

