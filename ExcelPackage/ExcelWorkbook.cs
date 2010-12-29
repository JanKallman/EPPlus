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
 *  Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Xml;
using System.IO;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Text;
using System.Security;

namespace OfficeOpenXml
{
	#region Public Enum ExcelCalcMode
	/// <summary>
	/// Represents the possible workbook calculation modes
	/// </summary>
	public enum ExcelCalcMode
	{
		/// <summary>
		/// Set the calculation mode to Automatic
		/// </summary>
		Automatic,
		/// <summary>
		/// Set the calculation mode to AutomaticNoTable
		/// </summary>
		AutomaticNoTable,
		/// <summary>
		/// Set the calculation mode to Manual
		/// </summary>
		Manual
	}
	#endregion

	/// <summary>
	/// Represents the Excel workbook and provides access to all the 
	/// document properties and worksheets within the workbook.
	/// </summary>
	public sealed class ExcelWorkbook : XmlHelper
	{
        internal class SharedStringItem
        {
            internal int pos;
            internal string Text;
            internal bool isRichText = false;
        }
        #region Private Properties
		internal ExcelPackage _package;
		// we have to hard code these uris as we need them to create a workbook from scratch
		private Uri _uriWorkbook = new Uri("/xl/workbook.xml", UriKind.Relative);
		private Uri _uriSharedStrings = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
		private Uri _uriStyles = new Uri("/xl/styles.xml", UriKind.Relative);
		private Uri _uriCalcChain = new Uri("/xl/calcChain.xml", UriKind.Relative);

		private XmlDocument _xmlWorkbook;
		private XmlDocument _xmlSharedStrings;
		private XmlDocument _xmlStyles;

		private ExcelWorksheets _worksheets;
		private OfficeProperties _properties;

        private ExcelStyles _styles;
        #endregion

		#region ExcelWorkbook Constructor
		/// <summary>
		/// Creates a new instance of the ExcelWorkbook class.  For internal use only!
		/// </summary>
		/// <param name="xlPackage">The parent package</param>
        /// <param name="namespaceManager">NamespaceManager</param>
		protected internal ExcelWorkbook(ExcelPackage xlPackage, XmlNamespaceManager namespaceManager) :
            base(namespaceManager)
		{
			_package = xlPackage;
            CreateWorkbookXml();
            TopNode = WorkbookXml.DocumentElement;
            SchemaNodeOrder = new string[] { "fileVersion", "workbookPr", "workbookProtection", "bookViews", "sheets", "definedNames", "calcPr", "pivotCaches" };
            GetSharedStrings();
		}
		#endregion

        internal Dictionary<string, SharedStringItem> _sharedStrings = new Dictionary<string, SharedStringItem>(); //Used when reading cells.
        internal List<SharedStringItem> _sharedStringsList = new List<SharedStringItem>(); //Used when reading cells.
        internal ExcelNamedRangeCollection _names=new ExcelNamedRangeCollection();
        internal int _nextDrawingID = 0;
        internal int _nextTableID = 1;
        internal int _nextPivotTableID = 1;
        /// <summary>
        /// Read shared strings to list
        /// </summary>
        private void GetSharedStrings()
        {
            XmlNodeList nl = SharedStringsXml.SelectNodes("//d:sst/d:si", NameSpaceManager);
            _sharedStringsList = new List<SharedStringItem>();
            if (nl != null)
            {
                foreach (XmlNode node in nl)
                {
                    XmlNode n = node.SelectSingleNode("d:t", NameSpaceManager);
                    if (n != null)
                    {
                        _sharedStringsList.Add(new SharedStringItem(){Text= n.InnerText});
                    }
                    else
                    {
                        _sharedStringsList.Add(new SharedStringItem(){Text= node.InnerXml, isRichText=true});
                    }
                }
            }
        }
        internal void GetDefinedNames()
        {
            XmlNodeList nl = WorkbookXml.SelectNodes("//d:definedNames/d:definedName", NameSpaceManager);
            if (nl != null)
            {
                foreach (XmlElement elem in nl)
                { 
                    string fullAddress = elem.InnerText;
                    //int splitPos = fullAddress.LastIndexOf('!');
                    //string sheet = fullAddress.Substring(0, splitPos);
                    //string address = fullAddress.Substring(splitPos + 1, fullAddress.Length - splitPos - 1);
                    
                    //if(sheet[0]=='\'') sheet = sheet.Substring(1, sheet.Length-2); //remove single quotes from sheet

                    int localSheetID;
                    if(!int.TryParse(elem.GetAttribute("localSheetId"), out localSheetID))
                    {
                        localSheetID = -1;
                    }
                    ExcelAddress addr = new ExcelAddress(fullAddress);
                    ExcelNamedRange namedRange;
                    if (localSheetID > -1)
                    {
                        if (string.IsNullOrEmpty(addr._ws))
                        {
                            namedRange = Worksheets[localSheetID + 1].Names.Add(elem.GetAttribute("name"), new ExcelRangeBase(Worksheets[localSheetID + 1], fullAddress));
                        }
                        else
                        {
                            namedRange = Worksheets[localSheetID + 1].Names.Add(elem.GetAttribute("name"), new ExcelRangeBase(Worksheets[addr._ws], fullAddress));
                        }
                    }
                    else
                    {
                        namedRange = _names.Add(elem.GetAttribute("name"), new ExcelRangeBase(Worksheets[addr._ws], fullAddress));
                    }
                    if (elem.GetAttribute("hidden") == "1") namedRange.IsNameHidden = true;
                }
            }
        }
		#region Worksheets
		/// <summary>
		/// Provides access to all the worksheets in the workbook.
		/// </summary>
		public ExcelWorksheets Worksheets
		{
			get
			{
				if (_worksheets == null)
				{
					_worksheets = new ExcelWorksheets(_package);
				}
				return (_worksheets);
			}
		}
		#endregion

        /// <summary>
        /// Provides access to named ranges
        /// </summary>
        public ExcelNamedRangeCollection Names
        {
            get
            {
                return _names;
            }
        }
		#region Workbook Properties
        decimal _standardFontWidth = decimal.MinValue;
        /// <summary>
        /// Max font width for the workbook
        /// </summary>
        public decimal MaxFontWidth 
        {
            get
            {
                if (_standardFontWidth == decimal.MinValue)
                {
                    var font = Styles.Fonts[0];
                    System.Drawing.Font f = new System.Drawing.Font(font.Name, font.Size);
                    _standardFontWidth=(decimal)(System.Windows.Forms.TextRenderer.MeasureText("00", f).Width - System.Windows.Forms.TextRenderer.MeasureText("0", f).Width);
                }
                return _standardFontWidth;
            }
        }
        ExcelProtection _protection = null;
        public ExcelProtection Protection
        {
            get
            {
                if (_protection == null)
                {
                    _protection = new ExcelProtection(NameSpaceManager, TopNode);
                    _protection.SchemaNodeOrder = SchemaNodeOrder;
                }
                return _protection;
            }
        }
        ExcelWorkbookView _view = null;
        public ExcelWorkbookView View
        {
            get
            {
                if (_view == null)
                {
                    _view = new ExcelWorkbookView(NameSpaceManager, TopNode);
                }
                return _view;
            }
        }
        /// <summary>
		/// The Uri to the workbook in the package
		/// </summary>
		protected internal Uri WorkbookUri { get { return (_uriWorkbook); }	}
		/// <summary>
		/// The Uri to the styles.xml in the package
		/// </summary>
		protected internal Uri StylesUri { get { return (_uriStyles); } }
		/// <summary>
		/// The Uri to the shared strings file
		/// </summary>
		protected internal Uri SharedStringsUri { get { return (_uriSharedStrings); } }
		/// <summary>
		/// Returns a reference to the workbook's part within the package
		/// </summary>
		protected internal PackagePart Part { get { return (_package.Package.GetPart(WorkbookUri)); } }
		
		#region WorkbookXml
		/// <summary>
		/// Provides access to the XML data representing the workbook in the package.
		/// </summary>
		public XmlDocument WorkbookXml
		{
			get
			{
				if (_xmlWorkbook == null)
				{
                    CreateWorkbookXml();
				}
				return (_xmlWorkbook);
			}
		}
        /// <summary>
        /// Create or read the XML for the workbook.
        /// </summary>
        private void CreateWorkbookXml()
        {
            if (_package.Package.PartExists(WorkbookUri))
                _xmlWorkbook = _package.GetXmlFromUri(WorkbookUri);
            else
            {
                // create a new workbook part and add to the package
                PackagePart partWorkbook = _package.Package.CreatePart(WorkbookUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", _package.Compression);

                // create the workbook
                _xmlWorkbook = new XmlDocument();
                _xmlWorkbook.PreserveWhitespace = ExcelPackage.preserveWhitespace;
                // create the workbook tag
                XmlElement tagWorkbook = _xmlWorkbook.CreateElement("workbook", ExcelPackage.schemaMain);
                // Add the relationships namespace
                ExcelPackage.AddSchemaAttribute(tagWorkbook, ExcelPackage.schemaRelationships, "r");
                _xmlWorkbook.AppendChild(tagWorkbook);

                //// create the bookViews tag
                XmlElement bookViews = _xmlWorkbook.CreateElement("bookViews", ExcelPackage.schemaMain);
                tagWorkbook.AppendChild(bookViews);
                XmlElement workbookView = _xmlWorkbook.CreateElement("workbookView", ExcelPackage.schemaMain);
                bookViews.AppendChild(workbookView);

                // create the sheets tag
                XmlElement tagSheets = _xmlWorkbook.CreateElement("sheets", ExcelPackage.schemaMain);
                tagWorkbook.AppendChild(tagSheets);

                // save it to the package
                StreamWriter streamWorkbook = new StreamWriter(partWorkbook.GetStream(FileMode.Create, FileAccess.Write));
                _xmlWorkbook.Save(streamWorkbook);
                streamWorkbook.Close();
                _package.Package.Flush();
            }
        }
		#endregion

		#region SharedStrings
		/// <summary>
		/// Provides access to the XML data representing the shared strings in the package.
		/// For internal use only!
		/// </summary>
		protected internal XmlDocument SharedStringsXml
		{
			get
			{
				if (_xmlSharedStrings == null)
				{
					if (_package.Package.PartExists(SharedStringsUri))
						_xmlSharedStrings = _package.GetXmlFromUri(SharedStringsUri);
					else
					{
						// create a new sharedStrings part and add to the package
                        PackagePart partStrings = _package.Package.CreatePart(SharedStringsUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", _package.Compression);

						// create the shared strings xml doc (with no entries in it)
						_xmlSharedStrings = new XmlDocument();
                        _xmlSharedStrings.PreserveWhitespace = ExcelPackage.preserveWhitespace;
                        _xmlSharedStrings.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" ?><sst count=\"0\" uniqueCount=\"0\" xmlns=\"{0}\" />", ExcelPackage.schemaMain));

						// save it to the package
						StreamWriter streamStrings = new StreamWriter(partStrings.GetStream(FileMode.Create, FileAccess.Write));
						_xmlSharedStrings.Save(streamStrings);
						streamStrings.Close();
						_package.Package.Flush();

						// create the relationship between the workbook and the new shared strings part
						Part.CreateRelationship(SharedStringsUri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/sharedStrings");
						_package.Package.Flush();
					}
				}
				return (_xmlSharedStrings);
			}
		}
		#endregion

		#region StylesXml
		/// <summary>
		/// Provides access to the XML data representing the styles in the package. 
		/// </summary>
		public XmlDocument StylesXml
		{
			get
			{
				if (_xmlStyles == null)
				{
					if (_package.Package.PartExists(StylesUri))
						_xmlStyles = _package.GetXmlFromUri(StylesUri);
					else
					{
						// create a new styles part and add to the package
                        PackagePart partSyles = _package.Package.CreatePart(StylesUri, @"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", _package.Compression);
						// create the style sheet
						_xmlStyles = new XmlDocument();
						XmlElement tagStylesheet = _xmlStyles.CreateElement("styleSheet", ExcelPackage.schemaMain);
                        tagStylesheet.SetAttribute("xmlns:d", ExcelPackage.schemaMain);

                        _xmlStyles.AppendChild(tagStylesheet);
                        //Create the numberformat tag.
                        XmlElement tagNumFmts = _xmlStyles.CreateElement("d", "numFmts", ExcelPackage.schemaMain);
                        tagStylesheet.AppendChild(tagNumFmts);
                        // create the fonts tag
                        XmlElement tagFonts = _xmlStyles.CreateElement("d", "fonts", ExcelPackage.schemaMain);
						tagFonts.SetAttribute("count", "1");
						tagStylesheet.AppendChild(tagFonts);
						// create the font tag
                        XmlElement tagFont = _xmlStyles.CreateElement("d", "font", ExcelPackage.schemaMain);
						tagFonts.AppendChild(tagFont);
						// create the sz tag
                        XmlElement tagSz = _xmlStyles.CreateElement("d", "sz", ExcelPackage.schemaMain);
						tagSz.SetAttribute("val", "11");
						tagFont.AppendChild(tagSz);
						// create the name tag
                        XmlElement tagName = _xmlStyles.CreateElement("d", "name", ExcelPackage.schemaMain);
                        tagName.SetAttribute("val", "Calibri");
						tagFont.AppendChild(tagName);

                        //Create the Fills tag.
                        XmlElement tagFills = _xmlStyles.CreateElement("d", "fills", ExcelPackage.schemaMain);
                        tagStylesheet.AppendChild(tagFills);
                        XmlElement tagFill = _xmlStyles.CreateElement("d", "fill", ExcelPackage.schemaMain);
                        tagFills.AppendChild(tagFill);
                        XmlElement tagPatternFill = _xmlStyles.CreateElement("d", "patternFill", ExcelPackage.schemaMain);
                        tagPatternFill.SetAttribute("patternType", "none");
                        tagFill.AppendChild(tagPatternFill);

                        //Strange behavior in excel?? Needed or fill bug out.                        
                        tagFill = _xmlStyles.CreateElement("d", "fill", ExcelPackage.schemaMain);
                        tagFills.AppendChild(tagFill);
                        tagPatternFill = _xmlStyles.CreateElement("d", "patternFill", ExcelPackage.schemaMain);
                        tagPatternFill.SetAttribute("d", "patternType", "gray125");
                        tagFill.AppendChild(tagPatternFill);

                        //Create the Borders tag.
                        XmlElement tagBorders = _xmlStyles.CreateElement("d", "borders", ExcelPackage.schemaMain);
                        tagStylesheet.AppendChild(tagBorders);
                        XmlElement tagBorder = _xmlStyles.CreateElement("d", "border", ExcelPackage.schemaMain);
                        tagBorders.AppendChild(tagBorder);
                        tagBorder.AppendChild(_xmlStyles.CreateElement("d", "left", ExcelPackage.schemaMain));
                        tagBorder.AppendChild(_xmlStyles.CreateElement("d", "right", ExcelPackage.schemaMain));
                        tagBorder.AppendChild(_xmlStyles.CreateElement("d", "top", ExcelPackage.schemaMain));
                        tagBorder.AppendChild(_xmlStyles.CreateElement("d", "bottom", ExcelPackage.schemaMain));
                        tagBorder.AppendChild(_xmlStyles.CreateElement("d", "diagonal", ExcelPackage.schemaMain));
                        
                        // create the cellStyleXfs tag
                        XmlElement tagCellStyleXfs = _xmlStyles.CreateElement("d", "cellStyleXfs", ExcelPackage.schemaMain);
						tagCellStyleXfs.SetAttribute("count", "1");
						tagStylesheet.AppendChild(tagCellStyleXfs);
						// create the xf tag
                        XmlElement tagXf = _xmlStyles.CreateElement("d", "xf", ExcelPackage.schemaMain);
						tagXf.SetAttribute("numFmtId", "0");
						tagXf.SetAttribute("fontId", "0");
						tagCellStyleXfs.AppendChild(tagXf);
						// create the cellXfs tag
                        XmlElement tagCellXfs = _xmlStyles.CreateElement("d", "cellXfs", ExcelPackage.schemaMain);
						tagCellXfs.SetAttribute("count", "1");
						tagStylesheet.AppendChild(tagCellXfs);
						// create the xf tag
                        XmlElement tagXf2 = _xmlStyles.CreateElement("d", "xf", ExcelPackage.schemaMain);
						tagXf2.SetAttribute("numFmtId", "0");
						tagXf2.SetAttribute("fontId", "0");
						tagXf2.SetAttribute("xfId", "0");
						tagCellXfs.AppendChild(tagXf2);

                        //Create the CellStyles tag.
                        XmlElement tagCellStyles = _xmlStyles.CreateElement("d", "cellStyles", ExcelPackage.schemaMain);
                        tagStylesheet.AppendChild(tagCellStyles);
                        XmlElement tagCellStyle = _xmlStyles.CreateElement("d", "cellStyle", ExcelPackage.schemaMain);
                        tagCellStyle.SetAttribute("name", "Normal");
                        tagCellStyle.SetAttribute("xfId", "0");
                        tagCellStyle.SetAttribute("builtinId", "0");

                        tagCellStyles.AppendChild(tagCellStyle);
                        
                        //Save it to the package
						StreamWriter streamStyles = new StreamWriter(partSyles.GetStream(FileMode.Create, FileAccess.Write));

						_xmlStyles.Save(streamStyles);
						streamStyles.Close();
						_package.Package.Flush();

						// create the relationship between the workbook and the new shared strings part
						_package.Workbook.Part.CreateRelationship(StylesUri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/styles");
						_package.Package.Flush();
					}
				}
				return (_xmlStyles);
			}
			set
			{
				_xmlStyles = value;
			}
		}
        /// <summary>
        /// Enables access to the Workbook style collection. 
        /// </summary>
        public ExcelStyles Styles
        {
            get
            {
                if (_styles == null)
                {
                    _styles = new ExcelStyles(NameSpaceManager, StylesXml, this);
                }
                return _styles;
            }
        }
        #endregion

		#region Office Document Properties
		/// <summary>
		/// Provides access to the office document properties
		/// </summary>
		public OfficeProperties Properties
		{
			get 
			{
                if (_properties == null)
                {
                    //  Create a NamespaceManager to handle the default namespace, 
                    //  and create a prefix for the default namespace:                   
                    _properties = new OfficeProperties(_package, NameSpaceManager);
                }
				return _properties;
			}
		}
		#endregion

		#region CalcMode
        private string CALC_MODE_PATH = "d:calcPr/@calcMode";
        /// <summary>
		/// Allows you to set the calculation mode for the workbook.
		/// </summary>
		public ExcelCalcMode CalcMode
		{
			get
			{
                string calcMode = GetXmlNodeString(CALC_MODE_PATH);
				switch (calcMode)
				{
					case "autoNoTable":
						return ExcelCalcMode.AutomaticNoTable;
					case "manual":
						return ExcelCalcMode.Manual;
                    default:
                        return ExcelCalcMode.Automatic;

				}
			}
			set
			{
				switch (value)
				{
					case ExcelCalcMode.AutomaticNoTable:
                        SetXmlNodeString(CALC_MODE_PATH, "autoNoTable") ;
						break;
					case ExcelCalcMode.Manual:
                        SetXmlNodeString(CALC_MODE_PATH, "manual");
						break;
                    default:
                        SetXmlNodeString(CALC_MODE_PATH, "auto");
                        break;

				}
			}
			#endregion
		}
		#endregion

		#region Workbook Private Methods
            
		#region Save // Workbook Save
		/// <summary>
		/// Saves the workbook and all its components to the package.
		/// For internal use only!
		/// </summary>
		protected internal void Save()  // Workbook Save
		{
			// ensure we have at least one worksheet
			if (Worksheets.Count == 0)
				throw new Exception("Workbook Save Error: the workbook must contain at least one worksheet!");

			#region Delete calcChain component
			// if the calcChain component exists, we should delete it to force Excel to recreate it
			// when the spreadsheet is next opened
			if (_package.Package.PartExists(_uriCalcChain))
			{
				//  there will be a relationship with the workbook, so first delete the relationship
				Uri calcChain = new Uri("calcChain.xml", UriKind.Relative);
				foreach (PackageRelationship relationship in _package.Workbook.Part.GetRelationships())
				{
					if (relationship.TargetUri == calcChain)
					{
						_package.Workbook.Part.DeleteRelationship(relationship.Id);
						break;
					}
				}
				// delete the calcChain component
				_package.Package.DeletePart(_uriCalcChain);
			}
			#endregion

            
            UpdateDefinedNamesXml();

            // save the workbook
			if (_xmlWorkbook != null)
			{
				_package.SavePart(WorkbookUri, _xmlWorkbook);
				_package.WriteDebugFile(_xmlWorkbook, "xl", "workbook.xml");
			}

			// save the properties of the workbook
			if (_properties != null)
			{
				_properties.Save();
			}

			// save the style sheet
            Styles.UpdateXml();
			_package.SavePart(StylesUri, _xmlStyles);
			_package.WriteDebugFile(_xmlStyles, "xl", "styles.xml");

            // save all the open worksheets
            var isProtected = Protection.LockWindows || Protection.LockStructure;
            foreach (ExcelWorksheet worksheet in Worksheets)
            {
                if (isProtected)
                {
                    worksheet.View.WindowProtection = true;
                }
                worksheet.Save();
            }
            
            // save the shared strings
			if (_xmlSharedStrings != null)
			{
                UpdateSharedStringsXml();
                _package.SavePart(SharedStringsUri, _xmlSharedStrings);
				_package.WriteDebugFile(_xmlSharedStrings, "xl", "sharedstrings.xml");
			}
		}

        private void UpdateSharedStringsXml()
        {
            XmlNode top = SharedStringsXml.SelectSingleNode("//d:sst", NameSpaceManager);
            top.RemoveAll();
            StringBuilder sbXml = new StringBuilder();
            foreach (string t in _sharedStrings.Keys)
            {

                SharedStringItem ssi = _sharedStrings[t];
                if (ssi.isRichText)
                {
                    XmlNode node = SharedStringsXml.CreateElement("si", ExcelPackage.schemaMain);
                    node.InnerXml = t;
                    top.AppendChild(node);
                }
                else
                {
                    XmlNode node = SharedStringsXml.CreateElement("si", ExcelPackage.schemaMain).AppendChild(SharedStringsXml.CreateElement("t", ExcelPackage.schemaMain));
                    node.InnerText = t;
                    top.AppendChild(node.ParentNode);
                }
                //sbXml.AppendFormat("<si><t>{0}</t>/</   si>", t.Replace("<", "&lt;").Replace(">", "&gt;"));
            }
            top.Attributes.Append(SharedStringsXml.CreateAttribute("count")).Value=_sharedStrings.Count.ToString();
            top.Attributes.Append(SharedStringsXml.CreateAttribute("uniqueCount")).Value = _sharedStrings.Count.ToString();
            //top.Attributes["uniqueCount"].Value = _sharedStrings.Count.ToString();

            //top.InnerXml = sbXml.ToString();
        }
        private void UpdateDefinedNamesXml()
        {
            try
            {
                XmlNode top = WorkbookXml.SelectSingleNode("//d:definedNames", NameSpaceManager);
                if (!ExistsNames())
                {
                    if (top != null) TopNode.RemoveChild(top);
                    return;
                }
                else
                {
                    if (top == null)
                    {
                        CreateNode("d:definedNames");
                        top = WorkbookXml.SelectSingleNode("//d:definedNames", NameSpaceManager);
                    }
                    else
                    {
                        top.RemoveAll();
                    }
                    foreach (ExcelNamedRange name in _names)
                    {

                        XmlElement elem = WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
                        top.AppendChild(elem);
                        elem.SetAttribute("name", name.Name);
                        if (name.IsNameHidden) elem.SetAttribute("hidden", "1");
                        elem.InnerText = name.FullAddressAbsolute;
                    }
                }
                foreach (ExcelWorksheet ws in _worksheets)
                {
                    foreach (ExcelNamedRange name in ws.Names)
                    {

                        XmlElement elem = WorkbookXml.CreateElement("definedName", ExcelPackage.schemaMain);
                        top.AppendChild(elem);
                        elem.SetAttribute("name", name.Name);
                        elem.SetAttribute("localSheetId", name.LocalSheetId.ToString());
                        if (name.IsNameHidden) elem.SetAttribute("hidden", "1");
                        elem.InnerText = name.FullAddressAbsolute;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Internal error updating named ranges ",ex);
            }
        }
        /// <summary>
        /// Is their any names in the workbook or in the sheets.
        /// </summary>
        /// <returns>?</returns>
        private bool ExistsNames()
        {
            if (_names.Count == 0)
            {
                foreach (ExcelWorksheet ws in Worksheets)
                {
                    if(ws.Names.Count>0)
                    {
                        return true;
                    }
                }
            }
            else
            {
                return true;
            }
            return false;
        }        
        #endregion

		#endregion

        internal bool ExistsTableName(string Name)
        {
            foreach (var ws in Worksheets)
            {
                if(ws.Tables._tableNames.ContainsKey(Name))
                {
                    return true;
                }
            }
            return false;
        }
        internal bool ExistsPivotTableName(string Name)
        {
            foreach (var ws in Worksheets)
            {
                if (ws.PivotTables._pivotTableNames.ContainsKey(Name))
                {
                    return true;
                }
            }
            return false;
        }

        internal void AddPivotTable(string cacheID, Uri defUri)
        {
            CreateNode("d:pivotCaches");

            XmlElement item = WorkbookXml.CreateElement("pivotCache", ExcelPackage.schemaMain);
            item.SetAttribute("cacheId", cacheID);
            var rel = Part.CreateRelationship(PackUriHelper.ResolvePartUri(WorkbookUri, defUri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotCacheDefinition");
            item.SetAttribute("id", ExcelPackage.schemaRelationships, rel.Id);


            var pivotCashes = WorkbookXml.SelectSingleNode("//d:pivotCaches", NameSpaceManager);
            pivotCashes.AppendChild(item);
        }
    } // end Workbook
}
