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
 * Jan Källman		    Initial Release		       2011-11-02
 * Jan Källman          Total rewrite               2010-03-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Xml;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Configuration;
using OfficeOpenXml.Drawing;
using System.Diagnostics;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Text;
using System.Security;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style.XmlAccess;
using System.Text.RegularExpressions;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Table;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Table.PivotTable;
using System.ComponentModel;
using System.Drawing;
using OfficeOpenXml.Calculation;

namespace OfficeOpenXml
{
    /// <summary>
    /// Worksheet hidden enumeration
    /// </summary>
    public enum eWorkSheetHidden
    {
        /// <summary>
        /// The worksheet is visible
        /// </summary>
        Visible,
        /// <summary>
        /// The worksheet is hidden but can be shown by the user via the user interface
        /// </summary>
        Hidden,
        /// <summary>
        /// The worksheet is hidden and cannot be shown by the user via the user interface
        /// </summary>
        VeryHidden
    }
    /// <summary>
	/// Represents an Excel worksheet and provides access to its properties and methods
	/// </summary>
	public sealed class ExcelWorksheet : XmlHelper, ICalcEngineFormulaInfo, ICalcEngineValueInfo
	{
        internal class Formulas
        {
            internal int Index { get; set; }
            internal string Address { get; set; }
            internal bool IsArray { get; set; }
            public string Formula { get; set; }
            public int StartRow { get; set; }
            public int StartCol { get; set; }

        }
        /// <summary>
        /// Collection containing merged cell addresses
        /// </summary>
        /// <typeparam name="T"></typeparam>
        public class MergeCellsCollection<T> : IEnumerable<T>
        {
            private List<T> _list = new List<T>();
            internal MergeCellsCollection()
            {

            }
            internal List<T> List { get {return _list;} }
            public T this[int Index]
            {
                get
                {
                    return _list[Index];
                }
            }
            public int Count
            {
                get
                {
                    return _list.Count;
                }
            }
            #region IEnumerable<T> Members

            public IEnumerator<T> GetEnumerator()
            {
                return _list.GetEnumerator();
            }

            #endregion

            #region IEnumerable Members

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return _list.GetEnumerator();
            }

            #endregion
        }
        internal RangeCollection _cells;
        internal RangeCollection _columns;
        internal RangeCollection _rows;

        internal Dictionary<int, Formulas> _sharedFormulas = new Dictionary<int, Formulas>();
        internal Dictionary<int, Formulas> _arrayFormulas = new Dictionary<int, Formulas>();
        internal RangeCollection _formulaCells;
        internal int _minCol = ExcelPackage.MaxColumns;
        internal int _maxCol = 0;
        internal List<ulong> _hyperLinkCells;   //Used when saving the sheet
        #region Worksheet Private Properties
        internal ExcelPackage _package;
		private Uri _worksheetUri;
		private string _name;
		private int _sheetID;
        private int _positionID;
		private string _relationshipID;
		private XmlDocument _worksheetXml;
        internal ExcelWorksheetView _sheetView;
		internal ExcelHeaderFooter _headerFooter;
        #endregion
        #region ExcelWorksheet Constructor
        /// <summary>
        /// A worksheet
        /// </summary>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="excelPackage">Package</param>
        /// <param name="relID">Relationship ID</param>
        /// <param name="uriWorksheet">URI</param>
        /// <param name="sheetName">Name of the sheet</param>
        /// <param name="sheetID">Sheet id</param>
        /// <param name="positionID">Position</param>
        /// <param name="hide">hide</param>
        public ExcelWorksheet(XmlNamespaceManager ns, ExcelPackage excelPackage, string relID, 
                              Uri uriWorksheet, string sheetName, int sheetID, int positionID,
                              eWorkSheetHidden hide) :
            base(ns, null)
        {
            SchemaNodeOrder = new string[] { "sheetPr", "tabColor", "outlinePr", "pageSetUpPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "protectedRanges","scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "linePrint", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "activeXControls", "webPublishItems", "tableParts" };
            _package = excelPackage;   
            _relationshipID = relID;
            _worksheetUri = uriWorksheet;
            _name = sheetName;
            _sheetID = sheetID;
            _positionID = positionID;
            Hidden = hide;
            _names = new ExcelNamedRangeCollection(Workbook,this);
            CreateXml();
            TopNode = _worksheetXml.DocumentElement;
        }

        #endregion
        /// <summary>
        /// The Uri to the worksheet within the package
        /// </summary>
        internal Uri WorksheetUri { get { return (_worksheetUri); } }
        /// <summary>
        /// The PackagePart for the worksheet within the package
        /// </summary>
        internal PackagePart Part { get { return (_package.Package.GetPart(WorksheetUri)); } }
        /// <summary>
        /// The ID for the worksheet's relationship with the workbook in the package
        /// </summary>
        internal string RelationshipID { get { return (_relationshipID); } }
        /// <summary>
        /// The unique identifier for the worksheet.
        /// </summary>
        internal int SheetID { get { return (_sheetID); } }
        /// <summary>
        /// The position of the worksheet.
        /// </summary>
        internal int PositionID { get { return (_positionID); } set { _positionID = value; } }
		#region Worksheet Public Properties
    	/// <summary>
        /// The index in the worksheets collection
        /// </summary>
        public int Index { get { return (_positionID); } }
        /// <summary>
        /// Address for autofilter
        /// <seealso cref="ExcelRangeBase.AutoFilter" />        
        /// </summary>
        public ExcelAddressBase AutoFilterAddress
        {
            get
            {
                string address = GetXmlNodeString("d:autoFilter/@ref");
                if (address == "")
                {
                    return null;
                }
                else
                {
                    return new ExcelAddressBase(address);
                }
            }
            internal set
            {
                SetXmlNodeString("d:autoFilter/@ref", value.Address);
            }
        }

		/// <summary>
		/// Returns a ExcelWorksheetView object that allows you to set the view state properties of the worksheet
		/// </summary>
		public ExcelWorksheetView View
		{
			get
			{
				if (_sheetView == null)
				{
                    XmlNode node = TopNode.SelectSingleNode("d:sheetViews/d:sheetView", NameSpaceManager);
                    if (node == null)
                    {
                        CreateNode("d:sheetViews/d:sheetView"); //this one shouls always exist. but check anyway
                        node = TopNode.SelectSingleNode("d:sheetViews/d:sheetView", NameSpaceManager);
                    }
                    _sheetView = new ExcelWorksheetView(NameSpaceManager, node, this);
				}
				return (_sheetView);
			}
		}

		/// <summary>
		/// The worksheet's display name as it appears on the tab
		/// </summary>
		public string Name
		{
			get { return (_name); }
			set
			{
                if (value == _name) return;
                Name=_package.Workbook.Worksheets.ValidateFixSheetName(Name);
                _package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@name", _sheetID), value);
				_name = value;
            }
		}
        internal ExcelNamedRangeCollection _names;
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
		/// <summary>
		/// Indicates if the worksheet is hidden in the workbook
		/// </summary>
		public eWorkSheetHidden Hidden
		{
			get 
            {
                string state=_package.Workbook.GetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
                if (state == "hidden")
                {
                    return eWorkSheetHidden.Hidden;
                }
                else if (state == "veryHidden")
                {
                    return eWorkSheetHidden.VeryHidden;
                }
                return eWorkSheetHidden.Visible;
            }
			set
			{
                    if (value == eWorkSheetHidden.Visible)
                    {
                        _package.Workbook.DeleteNode(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
                    }
                    else
                    {
                        string v;
                        v=value.ToString();                        
                        v=v.Substring(0,1).ToLower()+v.Substring(1);
                        _package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID),v );
                    }
		    }
		}
        double _defaultRowHeight = double.NaN;
        /// <summary>
		/// Get/set the default height of all rows in the worksheet
		/// </summary>
        public double DefaultRowHeight
		{
			get 
			{
				if(double.IsNaN(_defaultRowHeight))
                {
                    _defaultRowHeight = GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight");
                    if(double.IsNaN(_defaultRowHeight))
                    {
                        _defaultRowHeight = 15; // Excel default height
                    }
                }
				return _defaultRowHeight;
			}
			set
			{
                _defaultRowHeight = value;
                SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", value.ToString(CultureInfo.InvariantCulture));
                SetXmlNodeBool("d:sheetFormatPr/@customHeight", value != 15);

                if (double.IsNaN(GetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth")))
                {
                    DefaultColWidth = 9.140625;
                }
			}
		}
        /// <summary>
        /// Get/set the default width of all rows in the worksheet
        /// </summary>
        public double DefaultColWidth
        {
            get
            {
                double ret = GetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth");
                if (double.IsNaN(ret))
                {
                    ret = 9.140625; // Excel's default height
                }
                return ret;
            }
            set
            {
                SetXmlNodeString("d:sheetFormatPr/@defaultColWidth", value.ToString(CultureInfo.InvariantCulture));

                if (double.IsNaN(GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight")))
                {
                    DefaultRowHeight = 15;
                }
            }
        }
        /** <outlinePr applyStyles="1" summaryBelow="0" summaryRight="0" /> **/
        const string outLineSummaryBelowPath = "d:sheetPr/d:outlinePr/@summaryBelow";
        /// <summary>
        /// Summary rows below details 
        /// </summary>
        public bool OutLineSummaryBelow
        { 
            get
            {
                return GetXmlNodeBool(outLineSummaryBelowPath);
            }
            set
            {
                SetXmlNodeString(outLineSummaryBelowPath, value ? "1" : "0");
            }
        }
        const string outLineSummaryRightPath = "d:sheetPr/d:outlinePr/@summaryRight";
        /// <summary>
        /// Summary rows to right of details
        /// </summary>
        public bool OutLineSummaryRight
        {
            get
            {
                return GetXmlNodeBool(outLineSummaryRightPath);
            }
            set
            {
                SetXmlNodeString(outLineSummaryRightPath, value ? "1" : "0");
            }
        }
        const string outLineApplyStylePath = "d:sheetPr/d:outlinePr/@applyStyles";
        /// <summary>
        /// Automatic styles
        /// </summary>
        public bool OutLineApplyStyle
        {
            get
            {
                return GetXmlNodeBool(outLineApplyStylePath);
            }
            set
            {
                SetXmlNodeString(outLineApplyStylePath, value ? "1" : "0");
            }
        }
        const string tabColorPath = "d:sheetPr/d:tabColor/@rgb";
        /// <summary>
        /// Color of the sheet tab
        /// </summary>
        public Color TabColor
        {
            get
            {
                string col = GetXmlNodeString(tabColorPath);
                if (col == "")
                {
                    return Color.Empty;
                }
                else
                {
                    return Color.FromArgb(int.Parse(col, System.Globalization.NumberStyles.AllowHexSpecifier));
                }
            }
            set
            {
                SetXmlNodeString(tabColorPath, value.ToArgb().ToString("X"));
            }
        }
        #region WorksheetXml
		/// <summary>
		/// The XML document holding the worksheet data.
        /// All column, row, cell, pagebreak, merged cell and hyperlink-data are loaded into memory and removed from the document when loading the document.        
		/// </summary>
		public XmlDocument WorksheetXml
		{
			get
			{
				return (_worksheetXml);
			}
		}
        internal ExcelVmlDrawingCommentCollection _vmlDrawings = null;
        /// <summary>
        /// Vml drawings. underlaying object for comments
        /// </summary>
        internal ExcelVmlDrawingCommentCollection VmlDrawingsComments
        {
            get
            {
                if (_vmlDrawings == null)
                {
                    CreateVmlCollection();
                }
                return _vmlDrawings;
            }
        }
        internal ExcelCommentCollection _comments = null;
        /// <summary>
        /// Collection of comments
        /// </summary>
        public ExcelCommentCollection Comments
        {
            get
            {
                if (_comments == null)
                {
                    CreateVmlCollection();
                    _comments = new ExcelCommentCollection(_package, this, NameSpaceManager);
                }
                return _comments;
            }
        }
        private void CreateVmlCollection()
        {
            var vmlNode = _worksheetXml.DocumentElement.SelectSingleNode("d:legacyDrawing/@r:id", NameSpaceManager);
            if (vmlNode == null)
            {
                _vmlDrawings = new ExcelVmlDrawingCommentCollection(_package, this, null);
            }
            else
            {
                if (Part.RelationshipExists(vmlNode.Value))
                {
                    var rel = Part.GetRelationship(vmlNode.Value);
                    var vmlUri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

                    _vmlDrawings = new ExcelVmlDrawingCommentCollection(_package, this, vmlUri);
                    _vmlDrawings.RelId = rel.Id;
                }
            }
        }

        private void CreateXml()
        {
            _worksheetXml = new XmlDocument();
            _worksheetXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
            PackagePart packPart = _package.Package.GetPart(WorksheetUri);
            string xml = "";

            // First Columns, rows, cells, mergecells, hyperlinks and pagebreakes are loaded from a xmlstream to optimize speed...
            bool doAdjust = _package.DoAdjustDrawings;
            _package.DoAdjustDrawings = false;
            Stream stream = packPart.GetStream();
            XmlTextReader xr = new XmlTextReader(stream);            
            
            LoadColumns(xr);    //columnXml
            long start = stream.Position;
            LoadCells(xr);
            long end = stream.Position;
            LoadMergeCells(xr);
            LoadHyperLinks(xr);
            LoadRowPageBreakes(xr);
            LoadColPageBreakes(xr);
            //...then the rest of the Xml is extracted and loaded into the WorksheetXml document.
            stream.Seek(0, SeekOrigin.Begin);
            xml = GetWorkSheetXml(stream, start, end);

            //first char is invalid sometimes?? 
            if (xml[0] != '<') 
                _worksheetXml.LoadXml(xml.Substring(1,xml.Length-1));
            else
                _worksheetXml.LoadXml(xml);

            _package.DoAdjustDrawings = doAdjust;
            ClearNodes();
        }

        private void LoadRowPageBreakes(XmlTextReader xr)
        {
            if(!ReadUntil(xr, "rowBreaks","colBreaks")) return;
            while (xr.Read())
            {
                if (xr.LocalName == "brk")
                {
                    int id;
                    if (int.TryParse(xr.GetAttribute("id"), out id))
                    {
                        Row(id).PageBreak = true;
                    }
                }
                else
                {
                    break;
                }
            }
        }
        private void LoadColPageBreakes(XmlTextReader xr)
        {
            if (!ReadUntil(xr, "colBreaks")) return;
            while (xr.Read())
            {
                if (xr.LocalName == "brk")
                {
                    int id;
                    if (int.TryParse(xr.GetAttribute("id"), out id))
                    {
                        Column(id).PageBreak = true;
                    }
                }
                else
                {
                    break;
                }
            }
        }

        private void ClearNodes()
        {
            if (_worksheetXml.SelectSingleNode("//d:cols", NameSpaceManager)!=null)
            {
                _worksheetXml.SelectSingleNode("//d:cols", NameSpaceManager).RemoveAll();
            }
            if (_worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager).RemoveAll();
            }
            if (_worksheetXml.SelectSingleNode("//d:hyperlinks", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:hyperlinks", NameSpaceManager).RemoveAll();
            }
            if (_worksheetXml.SelectSingleNode("//d:rowBreaks", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:rowBreaks", NameSpaceManager).RemoveAll();
            }
            if (_worksheetXml.SelectSingleNode("//d:colBreaks", NameSpaceManager) != null)
            {
                _worksheetXml.SelectSingleNode("//d:colBreaks", NameSpaceManager).RemoveAll();
            }
        }
        const int BLOCKSIZE=8192;
        private string GetWorkSheetXml(Stream stream, long start, long end)
        {
            StreamReader sr = new StreamReader(stream);
            int length = 0;
            char[] block;
            int pos;
            StringBuilder sb = new StringBuilder();
            Match startmMatch, endMatch;
            do
            {
                int size = stream.Length < BLOCKSIZE ? (int)stream.Length : BLOCKSIZE;
                block = new char[size];
                pos = sr.ReadBlock(block, 0, size);                
                sb.Append(block);
                length += size;
            }
            while (length < start);
            startmMatch = Regex.Match(sb.ToString(), string.Format("(<[^>]*{0}[^>]*>)", "sheetData"));
            if (!startmMatch.Success) //Not found
            {
                return sb.ToString();
            }
            else
            {
                string s = sb.ToString();
                string xml = s.Substring(0, startmMatch.Index); 
                if(startmMatch.Value.EndsWith("/>"))
                {
                    xml += s.Substring(startmMatch.Index, s.Length - startmMatch.Index);
                }
                else
                {
                    if (sr.Peek() != -1)
                    {
                        if (end - BLOCKSIZE > 0)
                        {
                            long endSeekStart = end - BLOCKSIZE - 4096 < 0 ? 0 : (end - BLOCKSIZE - 4096);
                            stream.Seek(endSeekStart, SeekOrigin.Begin);
                            int size = (int)(stream.Length-endSeekStart);
                            block = new char[size];
                            sr = new StreamReader(stream);
                            pos = sr.ReadBlock(block, 0, size);
                            sb = new StringBuilder();
                            sb.Append(block);
                            s = sb.ToString();
                        }
                    }
                    endMatch = Regex.Match(s, string.Format("(</[^>]*{0}[^>]*>)", "sheetData"));
                    xml += "<sheetData/>" + s.Substring(endMatch.Index + endMatch.Length, s.Length - (endMatch.Index + endMatch.Length));
                }
                if (sr.Peek() > -1)
                {
                    xml += sr.ReadToEnd();
                }

                return xml;
            }
        }
        private void GetBlockPos(string xml, string tag, ref int start, ref int end)
        {
            Match startmMatch, endMatch;
            startmMatch = Regex.Match(xml, string.Format("(<[^>]*{0}[^>]*>)", tag)); //"<[a-zA-Z:]*" + tag + "[?]*>");

            if (!startmMatch.Success) //Not found
            {
                start = -1;
                end = -1;
                return;
            }
            start=startmMatch.Index;
            if(startmMatch.Value.Substring(startmMatch.Value.Length-2,1)=="/")
            {
                end=startmMatch.Index+startmMatch.Length;
            }
            else
            {
                endMatch = Regex.Match(xml, string.Format("(</[^>]*{0}[^>]*>)", tag));
                if (endMatch.Success)
                {
                    end = endMatch.Index + endMatch.Length;
                }
            }
        }
        private bool ReadUntil(XmlTextReader xr,params string[] tagName)
        {
            if (xr.EOF) return false;
            while (!Array.Exists(tagName, tag => xr.LocalName.EndsWith(tag)))
            {
                xr.Read();
                if (xr.EOF) return false;
            }
            return (xr.LocalName.EndsWith(tagName[0]));
        }
        private void LoadColumns (XmlTextReader xr)//(string xml)
        {
            var colList = new List<IRangeID>();
            if (ReadUntil(xr, "cols", "sheetData"))
            {
            //if (xml != "")
            //{
                //var xr=new XmlTextReader(new StringReader(xml));
                while(xr.Read())
                {
                    if(xr.LocalName!="col") break;
                    int min = int.Parse(xr.GetAttribute("min"));

                    int style;
                    if (xr.GetAttribute("style") == null || !int.TryParse(xr.GetAttribute("style"), out style))
                    {
                        style = 0;
                    }
                    ExcelColumn col = new ExcelColumn(this, min);
                   
                    col._columnMax = int.Parse(xr.GetAttribute("max")); 
                    col.StyleID = style;
                    col.Width = xr.GetAttribute("width") == null ? 0 : double.Parse(xr.GetAttribute("width"), CultureInfo.InvariantCulture); 
                    col.BestFit = xr.GetAttribute("bestFit") != null && xr.GetAttribute("bestFit") == "1" ? true : false;
                    col.Collapsed = xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed") == "1" ? true : false;
                    col.Phonetic = xr.GetAttribute("phonetic") != null && xr.GetAttribute("phonetic") == "1" ? true : false;
                    col.OutlineLevel = xr.GetAttribute("outlineLevel") == null ? 0 : int.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture);
                    col.Hidden = xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1" ? true : false;
                    colList.Add(col);
                }
            }
            _columns = new RangeCollection(colList);
        }
        /// <summary>
        /// Read until the node is found. If not found the xmlreader is reseted.
        /// </summary>
        /// <param name="xr">The reader</param>
        /// <param name="nodeText">Text to search for</param>
        /// <param name="altNode">Alternative text to search for</param>
        /// <returns></returns>
        private static bool ReadXmlReaderUntil(XmlTextReader xr, string nodeText, string altNode)
        {
            do
            {
                if (xr.LocalName == nodeText || xr.LocalName == altNode) return true;
            }
            while(xr.Read());
            xr.Close();
            return false;
        }
        /// <summary>
        /// Load Hyperlinks
        /// </summary>
        /// <param name="xr">The reader</param>
        private void LoadHyperLinks(XmlTextReader xr)
        {
            if(!ReadUntil(xr, "hyperlinks", "rowBreaks", "colBreaks")) return;
            while (xr.Read())
            {
                if (xr.LocalName == "hyperlink")
                {
                    int fromRow, fromCol, toRow, toCol;
                    ExcelCell.GetRowColFromAddress(xr.GetAttribute("ref"), out fromRow, out fromCol, out toRow, out toCol);
                    ulong id = ExcelCell.GetCellID(_sheetID, fromRow, fromCol);
                    ExcelCell cell = _cells[id] as ExcelCell;
                    if (xr.GetAttribute("id", ExcelPackage.schemaRelationships) != null)
                    {
                        cell.HyperLinkRId = xr.GetAttribute("id", ExcelPackage.schemaRelationships);
                        cell.Hyperlink = new ExcelHyperLink(Part.GetRelationship(cell.HyperLinkRId).TargetUri.AbsoluteUri);
                        Part.DeleteRelationship(cell.HyperLinkRId); //Delete the relationship, it is recreated when we save the package.
                    }
                    else if (xr.GetAttribute("location") != null)
                    {
                        ExcelHyperLink hl = new ExcelHyperLink(xr.GetAttribute("location"), xr.GetAttribute("display"));
                        hl.RowSpann = toRow - fromRow;
                        hl.ColSpann = toCol - fromCol;
                        string tt=xr.GetAttribute("tooltip");
                        if(!string.IsNullOrEmpty(tt))
                        {
                            hl.ToolTip=tt;
                        }                        
                        cell.Hyperlink = hl;
                    }
                }
                else
                {
                    break;
                }
            }
        }
        /// <summary>
        /// Load cells
        /// </summary>
        /// <param name="xr">The reader</param>
        private void LoadCells(XmlTextReader xr)
        {
            var cellList=new List<IRangeID>();
            var rowList = new List<IRangeID>();
            var formulaList = new List<IRangeID>();
            string v="";
            ReadUntil(xr, "sheetData", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks");
            ExcelCell cell = null;
            xr.Read();
            
            while (!xr.EOF)
            {
                while (xr.NodeType == XmlNodeType.EndElement)
                {
                    xr.Read();
                }                
                if (xr.LocalName == "row")
                {
                    int row = Convert.ToInt32(xr.GetAttribute("r"));

                    if (xr.AttributeCount > 2 || (xr.AttributeCount == 2 && xr.GetAttribute("spans") != null))
                    {
                        rowList.Add(AddRow(xr, row));
                    }
                    xr.Read();
                }
                else if (xr.LocalName == "c")
                {
                    if (cell != null) cellList.Add(cell);
                    cell = new ExcelCell(this, xr.GetAttribute("r"));
                    if (xr.GetAttribute("t") != null) cell.DataType = xr.GetAttribute("t");
                    cell.StyleID = xr.GetAttribute("s") == null ? 0 : int.Parse(xr.GetAttribute("s"));
                    xr.Read();
                }
                else if (xr.LocalName == "v")
                {
                    cell._value = GetValueFromXml(cell, xr);
                    xr.Read();
                }
                else if (xr.LocalName == "f")
                {
                    string t = xr.GetAttribute("t");
                    if (t == null)
                    {
                        cell._formula = xr.ReadElementContentAsString();
                        formulaList.Add(cell);
                    }
                    else if (t == "shared")
                    {

                        string si = xr.GetAttribute("si");
                        if (si != null)
                        {
                            cell._sharedFormulaID = int.Parse(si);
                            string address = xr.GetAttribute("ref");
                            string formula = xr.ReadElementContentAsString();
                            if (formula != "")
                            {
                                _sharedFormulas.Add(cell.SharedFormulaID, new Formulas() { Index = cell.SharedFormulaID, Formula = formula, Address = address, StartRow = cell.Row, StartCol = cell.Column });
                            }
                        }
                        else
                        {
                            xr.Read();  //Something is wrong in the sheet, read next
                        }
                    }
                    else if (t == "array") //TODO: Array functions are not support yet. Read the formula for the start cell only.
                    {
                        string address = xr.GetAttribute("ref");
                        cell._formula = xr.ReadElementContentAsString();
                        cell._sharedFormulaID = GetMaxShareFunctionIndex(true); //We use the shared formula id here, just so we can use the same dictionary for both Array and Shared forulas.
                        _sharedFormulas.Add(cell._sharedFormulaID, new Formulas() { Index = cell._sharedFormulaID, Formula = cell._formula, Address = address, StartRow = cell.Row, StartCol = cell.Column, IsArray = true });
                    }
                    else // ??? some other type
                    {
                        xr.Read();  //Something is wrong in the sheet, read next
                    }

                }
                else if (xr.LocalName == "is")   //Inline string
                {
                    xr.Read();
                    if (xr.LocalName == "t")
                    {
                        cell._value = xr.ReadInnerXml();
                    }
                    else
                    {
                        cell._value = xr.ReadOuterXml();
                        cell.IsRichText = true;
                    }
                }
                else
                {
                    break;
                }
            }
            if (cell != null) cellList.Add(cell);

            _cells = new RangeCollection(cellList);
            _rows = new RangeCollection(rowList);
            _formulaCells = new RangeCollection(formulaList);
        }
        /// <summary>
        /// Load merged cells
        /// </summary>
        /// <param name="xr"></param>
        private void LoadMergeCells(XmlTextReader xr)
        {
            if(ReadUntil(xr, "mergeCells", "hyperlinks", "rowBreaks", "colBreaks") && !xr.EOF)
            {
                while (xr.Read())
                {
                    if (xr.LocalName != "mergeCell") break;

                    string address = xr.GetAttribute("ref");
                    int fromRow, fromCol, toRow, toCol;
                    ExcelCellBase.GetRowColFromAddress(address, out fromRow, out fromCol, out toRow, out toCol);
                    for (int row = fromRow; row <= toRow; row++)
                    {
                        for (int col = fromCol; col <= toCol; col++)
                        {
                            Cell(row, col).Merge = true;
                        }
                    }

                    _mergedCells.List.Add(address);
                }
            }
        }
        /// <summary>
        /// Update merged cells
        /// </summary>
        /// <param name="sw">The writer</param>
        private void UpdateMergedCells(StreamWriter sw)
        {
            sw.Write("<mergeCells>");
            foreach (string address in _mergedCells)
            {
                sw.Write("<mergeCell ref=\"{0}\" />", address);
            }
            sw.Write("</mergeCells>");
        }
        /// <summary>
        /// Reads a row from the XML reader
        /// </summary>
        /// <param name="xr">The reader</param>
        /// <param name="row">The row number</param>
        /// <returns></returns>
        private ExcelRow AddRow(XmlTextReader xr, int row)
        {
            ExcelRow r = new ExcelRow(this, row);

            r.Collapsed = xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed")== "1" ? true : false;
            if (xr.GetAttribute("ht") != null) r.Height = double.Parse(xr.GetAttribute("ht"), CultureInfo.InvariantCulture);
            r.Hidden = xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1" ? true : false; ;
            r.OutlineLevel = xr.GetAttribute("outlineLevel") == null ? 0 : int.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture); ;
            r.Phonetic = xr.GetAttribute("ph") != null && xr.GetAttribute("ph") == "1" ? true : false; ;
            r.StyleID = xr.GetAttribute("s") == null ? 0 : int.Parse(xr.GetAttribute("s"), CultureInfo.InvariantCulture);
            r.CustomHeight = xr.GetAttribute("customHeight") == null ? false : xr.GetAttribute("customHeight")=="1";
            return r;
        }

        private object GetValueFromXml(ExcelCell cell, XmlTextReader xr)
        {
            object value;
            //XmlNode vnode = colNode.SelectSingleNode("d:v", NameSpaceManager);
            //if (vnode == null) return null;

            if (cell.DataType == "s")
            {
                int ix = xr.ReadElementContentAsInt();
                value = _package.Workbook._sharedStringsList[ix].Text;
                cell.IsRichText = _package.Workbook._sharedStringsList[ix].isRichText;
            }
            else if (cell.DataType == "str")
            {
                value = xr.ReadElementContentAsString();
            }
            else if (cell.DataType == "b")
            {
                value = (xr.ReadElementContentAsString()!="0");
            }
            else
            {
                int n = cell.Style.Numberformat.NumFmtID;
                string v = xr.ReadElementContentAsString();

                if ((n >= 14 && n <= 22) || (n >= 45 && n <= 47))
                {
                    double res;
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out res))
                    {
                        value = DateTime.FromOADate(res);
                    }
                    else
                    {
                        value = "";
                    }
                }
                else
                {
                    double d;
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    {
                        value = d;
                    }
                    else
                    {
                        value = double.NaN;
                    }
                }
            }
            return value;
        }
        //private string GetSharedString(int stringID)
        //{
        //    string retValue = null;
        //    XmlNodeList stringNodes = xlPackage.Workbook.SharedStringsXml.SelectNodes(string.Format("//d:si", stringID), NameSpaceManager);
        //    XmlNode stringNode = stringNodes[stringID];
        //    if (stringNode != null)
        //        retValue = stringNode.InnerText;
        //    return (retValue);
        //}
        #endregion
		#region HeaderFooter
		/// <summary>
		/// A reference to the header and footer class which allows you to 
		/// set the header and footer for all odd, even and first pages of the worksheet
        /// </summary>
        /// <remarks>
        /// To format the text you can use the following format
        /// <list type="table">
        /// <listheader><term>Prefix</term><description>Description</description></listheader>
        /// <item><term>&amp;U</term><description>Underlined</description></item>
        /// <item><term>&amp;E</term><description>Double Underline</description></item>
        /// <item><term>&amp;K:xxxxxx</term><description>Color. ex &amp;K:FF0000 for red</description></item>
        /// <item><term>&amp;"Font,Regular Bold Italic"</term><description>Changes the font. Regular or Bold or Italic or Bold Italic can be used. ex &amp;"Arial,Bold Italic"</description></item>
        /// <item><term>&amp;nn</term><description>Change font size. nn is an integer. ex &amp;24</description></item>
        /// <item><term>&amp;G</term><description>Placeholder for images. Images can not be added by the library, but its possible to use in a template.</description></item>
        /// </list>
        /// </remarks>
        public ExcelHeaderFooter HeaderFooter
		{
			get
			{
				if (_headerFooter == null)
				{
                    XmlNode headerFooterNode = TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager);
                    if (headerFooterNode == null)
                        headerFooterNode= CreateNode("d:headerFooter");
                    _headerFooter = new ExcelHeaderFooter(NameSpaceManager, headerFooterNode, this);
				}                
				return (_headerFooter);
			}
		}
		#endregion

        #region "PrinterSettings"
        /// <summary>
        /// Printer settings
        /// </summary>
        public ExcelPrinterSettings PrinterSettings
        {
            get
            {
                var ps = new ExcelPrinterSettings(NameSpaceManager, TopNode, this);
                ps.SchemaNodeOrder = SchemaNodeOrder;
                return ps;
            }
        }
        #endregion

		#endregion // END Worksheet Public Properties

		#region Worksheet Public Methods
        
        /// <summary>
		/// Provides access to an individual cell within the worksheet.
		/// </summary>
		/// <param name="row">The row number in the worksheet</param>
		/// <param name="col">The column number in the worksheet</param>
		/// <returns></returns>		
        internal ExcelCell Cell(int row, int col)
        {
            ulong cellID=ExcelCell.GetCellID(SheetID, row, col);
            if (!_cells.ContainsKey(cellID))
            {
                _cells.Add(new ExcelCell(this, row, col));
            }
            return _cells[cellID] as ExcelCell;
        }
        /// <summary>
        /// Provides access to a range of cells
        /// </summary>  
        public ExcelRange Cells
        {
            get
            {
                return new ExcelRange(this, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            }
        }
        /// <summary>
        /// Provides access to the selected range of cells
        /// </summary>  
        public ExcelRange SelectedRange
        {
            get
            {
                return new ExcelRange(this, View.SelectedRange);
            }
        }
        MergeCellsCollection<string> _mergedCells = new MergeCellsCollection<string>();
        /// <summary>
        /// Addresses to merged ranges
        /// </summary>
        public MergeCellsCollection<string> MergedCells
        {
            get
            {
                return _mergedCells;
            }
        }
        /// <summary>
		/// Provides access to an individual row within the worksheet so you can set its properties.
		/// </summary>
		/// <param name="row">The row number in the worksheet</param>
		/// <returns></returns>
		public ExcelRow Row(int row)
		{
            ExcelRow r;
            ulong id = ExcelRow.GetRowID(_sheetID, row);
            if (_rows.ContainsKey(id))
            {
                r = _rows[id] as ExcelRow;
            }
            else
            {
                r = new ExcelRow(this, row);
                _rows.Add(r);
            }
            return r;
		}
		/// <summary>
		/// Provides access to an individual column within the worksheet so you can set its properties.
		/// </summary>
		/// <param name="col">The column number in the worksheet</param>
		/// <returns></returns>
		public ExcelColumn Column(int col)
		{
            ExcelColumn column;
            ulong id=ExcelColumn.GetColumnID(_sheetID, col);
            if (_columns.ContainsKey(id))
            {
                column = _columns[id] as ExcelColumn;
                if (column.ColumnMin != column.ColumnMax)
                {
                    int maxCol = column.ColumnMax;
                    column.ColumnMax=col;
                    ExcelColumn copy = CopyColumn(column, col+1);
                    copy.ColumnMax = maxCol;
                }
            }
            else
            {
                foreach (ExcelColumn checkColumn in _columns)
                {
                    if (col > checkColumn.ColumnMin && col <= checkColumn.ColumnMax)
                    {
                        int maxCol = checkColumn.ColumnMax;
                        checkColumn.ColumnMax = col - 1;
                        if (maxCol > col)
                        {
                            ExcelColumn newC = CopyColumn(checkColumn, col + 1);
                            newC.ColumnMax = maxCol;
                        }
                        return CopyColumn(checkColumn, col);                        
                    }
                }
                column = new ExcelColumn(this, col);
                _columns.Add(column);
             }
            return column;
		}
        /// <summary>
        /// Returns the name of the worksheet
        /// </summary>
        /// <returns>The name of the worksheet</returns>
        public override string ToString()
        {
            return Name;
        } 
        internal ExcelColumn CopyColumn(ExcelColumn c, int col)
        {
            ExcelColumn newC = new ExcelColumn(this, col);
            if (c.StyleName != "")
                newC.StyleName = c.StyleName;
            else
                newC.StyleID = c.StyleID;

            newC.Width = c.Width;
            newC.Hidden = c.Hidden;
            newC.OutlineLevel = c.OutlineLevel;
            newC.Phonetic = c.Phonetic;
            newC.BestFit = c.BestFit;
            _columns.Add(newC);
            return newC;
       }
        /// <summary>
        /// Selects a range in the worksheet. The active cell is the topmost cell.
        /// Make the current worksheet active.
        /// </summary>
        /// <param name="Address">An address range</param>
        public void Select(string Address)
        {
            Select(Address, true);
        }
        /// <summary>
        /// Selects a range in the worksheet. The actice cell is the topmost cell.
        /// </summary>
        /// <param name="Address">A range of cells</param>
        /// <param name="SelectSheet">Make the sheet active</param>
        public void Select(string Address, bool SelectSheet)
        {
            int fromCol, fromRow, toCol, toRow;
            //Get rows and columns and validate as well
            ExcelCell.GetRowColFromAddress(Address, out fromRow, out fromCol, out toRow, out toCol);

            if (SelectSheet)
            {
                View.TabSelected = true;
            }
            View.SelectedRange = Address;
            View.ActiveCell = ExcelCell.GetAddress(fromRow, fromCol);            
        }
        /// <summary>
        /// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
        /// Make the current worksheet active.
        /// </summary>
        /// <param name="Address">An address range</param>
        public void Select(ExcelAddress Address)
        {
            Select(Address, true);
        }
        /// <summary>
        /// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
        /// </summary>
        /// <param name="Address">A range of cells</param>
        /// <param name="SelectSheet">Make the sheet active</param>
        public void Select(ExcelAddress Address, bool SelectSheet)
        {

            if (SelectSheet)
            {
                View.TabSelected = true;
            }
            string selAddress = ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column) + ":" + ExcelCellBase.GetAddress(Address.End.Row, Address.End.Column);
            if (Address.Addresses != null)
            {
                foreach (var a in Address.Addresses)
                {
                    selAddress += " " + ExcelCellBase.GetAddress(a.Start.Row, a.Start.Column) + ":" + ExcelCellBase.GetAddress(a.End.Row, a.End.Column);
                }
            }
            View.SelectedRange = selAddress;
            View.ActiveCell = ExcelCell.GetAddress(Address.Start.Row, Address.Start.Column);
        }

		#region InsertRow
        /// <summary>
        /// Inserts a new row into the spreadsheet.  Existing rows below the position are 
        /// shifted down.  All formula are updated to take account of the new row.
        /// </summary>
        /// <param name="rowFrom">The position of the new row</param>
        /// <param name="rows">Number of rows to insert</param>
        public void InsertRow(int rowFrom, int rows)
        {
            InsertRow(rowFrom, rows, 0);
        }
        /// <summary>
		/// Inserts a new row into the spreadsheet.  Existing rows below the position are 
		/// shifted down.  All formula are updated to take account of the new row.
		/// </summary>
        /// <param name="rowFrom">The position of the new row</param>
        /// <param name="rows">Number of rows to insert.</param>
        /// <param name="copyStylesFromRow">Copy Styles from this row. Applied to all inserted rows</param>
		public void InsertRow(int rowFrom, int rows, int copyStylesFromRow)
		{
            //Insert the new row into the collection
            ulong copyRowID=ExcelRow.GetRowID(SheetID, copyStylesFromRow);
            List<ExcelCell> copyStylesCells=new List<ExcelCell>();
            if (copyStylesFromRow > 0)
            {
                int startIndex = _cells.IndexOf(copyRowID);
                startIndex = ~startIndex;
                while(startIndex < _cells.Count && (_cells[startIndex] as ExcelCell).Row==copyStylesFromRow)
                {
                    copyStylesCells.Add(_cells[startIndex++] as ExcelCell);
                }
            }
            ulong rowID=ExcelRow.GetRowID(SheetID, rowFrom);

            _cells.InsertRows(rowID, rows);
            _rows.InsertRows(rowID, rows);
            _formulaCells.InsertRowsUpdateIndex(rowID, rows);
            if (_comments != null) _comments._comments.InsertRowsUpdateIndex(rowID, rows);
            if (_vmlDrawings != null) _vmlDrawings._drawings.InsertRowsUpdateIndex(rowID, rows);

            foreach (ExcelCell cell in _formulaCells)
            {
                if (cell.SharedFormulaID < 0)
                {
                    cell.Formula = ExcelCell.UpdateFormulaReferences(cell.Formula, rows, 0, rowFrom, 0);
                }
                else
                {
                    throw new Exception("Shared formula error");
                }
            }

            FixSharedFormulasRows(rowFrom, rows);
            
            FixMergedCells(rowFrom, rows,false);

            //Copy the styles
            foreach (ExcelCell cell in copyStylesCells)
            {
                Cells[rowFrom, cell.Column, rowFrom + rows - 1, cell.Column].StyleID = cell.StyleID;
            }
        }
        /// <summary>
        /// Adds a value to the row of merged cells to fix for inserts or deletes
        /// </summary>
        /// <param name="position"></param>
        /// <param name="rows"></param>
        /// <param name="delete"></param>
        private void FixMergedCells(int position, int rows, bool delete)
        {
            List<int> removeIndex = new List<int>();
            for (int i = 0; i < _mergedCells.Count; i++)
            {
                ExcelAddressBase addr = new ExcelAddressBase(_mergedCells[i]), newAddr ;
                if (delete)
                {
                    newAddr=addr.DeleteRow(position, rows);
                    if (newAddr == null)
                    {
                        removeIndex.Add(i);
                        continue;
                    }
                }
                else
                {
                    newAddr = addr.AddRow(position, rows);
                }
                
                //The address has changed.
                if (newAddr._address != addr._address)
                {
                    //Set merged prop for cells
                    for (int row = newAddr._fromRow; row <= newAddr._toRow; row++)
                    {
                        for (int col = newAddr._fromCol; col <= newAddr._toCol; col++)
                        {
                            Cell(row, col).Merge = true;
                        }
                    }
                }

                _mergedCells.List[i] = newAddr._address;
            }
            for (int i = removeIndex.Count - 1; i >= 0; i--)
            {
                _mergedCells.List.RemoveAt(removeIndex[i]);
            }
        }
        private void FixSharedFormulasRows(int position, int rows)
        {
            List<Formulas> added = new List<Formulas>();
            List<Formulas> deleted = new List<Formulas>();

            foreach (int id in _sharedFormulas.Keys)
            {
                var f = _sharedFormulas[id];
                int fromCol, fromRow, toCol, toRow;

                ExcelCell.GetRowColFromAddress(f.Address, out fromRow, out fromCol, out toRow, out toCol);
                if (position >= fromRow && position+(Math.Abs(rows)) <= toRow) //Insert/delete is whithin the share formula address
                {
                    if (rows > 0) //Insert
                    {
                        f.Address = ExcelCell.GetAddress(fromRow, fromCol) + ":" + ExcelCell.GetAddress(position - 1, toCol);
                        if (toRow != fromRow)
                        {
                            Formulas newF = new Formulas();
                            newF.StartCol = f.StartCol;
                            newF.StartRow = position + rows;
                            newF.Address = ExcelCell.GetAddress(position + rows, fromCol) + ":" + ExcelCell.GetAddress(toRow + rows, toCol);
                            newF.Formula = ExcelCell.TranslateFromR1C1(ExcelCell.TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), position, f.StartCol);
                            added.Add(newF);
                        }
                    }
                    else
                    {
                        if (fromRow - rows < toRow)
                        {
                            f.Address = ExcelCell.GetAddress(fromRow, fromCol, toRow+rows, toCol);
                        }
                        else
                        {
                            f.Address = ExcelCell.GetAddress(fromRow, fromCol) + ":" + ExcelCell.GetAddress(toRow + rows, toCol);
                        }
                    }
                }
                else if (position <= toRow)
                {
                    if (rows > 0) //Insert before shift down
                    {
                        f.StartRow += rows;
                        //f.Formula = ExcelCell.UpdateFormulaReferences(f.Formula, rows, 0, position, 0); //Recalc the cells positions
                        f.Address = ExcelCell.GetAddress(fromRow + rows, fromCol) + ":" + ExcelCell.GetAddress(toRow + rows, toCol);
                    }
                    else
                    {
                        //Cells[f.Address].SetSharedFormulaID(int.MinValue);
                        if (position <= fromRow && position + Math.Abs(rows) > toRow)  //Delete the formula 
                        {
                            deleted.Add(f);
                        }
                        else
                        {
                            toRow = toRow + rows < position - 1 ? position - 1 : toRow + rows;
                            if (position <= fromRow)
                            {
                                fromRow = fromRow + rows < position ? position : fromRow + rows;
                            }

                            f.Address = ExcelCell.GetAddress(fromRow, fromCol, toRow, toCol);
                            Cells[f.Address].SetSharedFormulaID(f.Index);
                            //f.StartRow = fromRow;

                            //f.Formula = ExcelCell.UpdateFormulaReferences(f.Formula, rows, 0, position, 0);
                       
                        }
                    }
                }
            }

            AddFormulas(added, position, rows);

            //Remove formulas
            foreach (Formulas f in deleted)
            {
                _sharedFormulas.Remove(f.Index);
            }

            //Fix Formulas
            added = new List<Formulas>();
            foreach (int id in _sharedFormulas.Keys)
            {
                var f = _sharedFormulas[id];
                UpdateSharedFormulaRow(ref f, position, rows, ref added);
            }
            AddFormulas(added, position, rows);
        }

        private void AddFormulas(List<Formulas> added, int position, int rows)
        {
            //Add new formulas
            foreach (Formulas f in added)
            {
                f.Index = GetMaxShareFunctionIndex(false);
                _sharedFormulas.Add(f.Index, f);
                Cells[f.Address].SetSharedFormulaID(f.Index);
            }
        }

        private void UpdateSharedFormulaRow(ref Formulas formula, int startRow, int rows, ref List<Formulas> newFormulas)
        {
            int fromRow,fromCol, toRow, toCol;
            int newFormulasCount = newFormulas.Count;
            ExcelCellBase.GetRowColFromAddress(formula.Address, out fromRow, out fromCol, out toRow, out toCol);
            //int refSplits = Regex.Split(formula.Formula, "#REF!").GetUpperBound(0);
            string formualR1C1;
            if (rows > 0 || fromRow <= startRow)
            {
                formualR1C1 = ExcelRangeBase.TranslateToR1C1(formula.Formula, formula.StartRow, formula.StartCol);
                formula.Formula = ExcelRangeBase.TranslateFromR1C1(formualR1C1, fromRow, formula.StartCol);
            }
            else
            {
                formualR1C1 = ExcelRangeBase.TranslateToR1C1(formula.Formula, formula.StartRow-rows, formula.StartCol);
                formula.Formula = ExcelRangeBase.TranslateFromR1C1(formualR1C1, formula.StartRow, formula.StartCol);
            }
            //bool isRef = false;
            //Formulas restFormula=formula;
            string prevFormualR1C1 = formualR1C1;
            for (int row = fromRow; row <= toRow; row++)
            {
                for (int col = fromCol; col <= toCol; col++)
                {
                    string newFormula;
                    string currentFormulaR1C1;
                    if (rows > 0 || row < startRow)
                    {
                        newFormula = ExcelCellBase.UpdateFormulaReferences(ExcelCell.TranslateFromR1C1(formualR1C1, row, col), rows, 0, startRow, 0);
                        currentFormulaR1C1 = ExcelRangeBase.TranslateToR1C1(newFormula, row, col);
                    }
                    else
                    {
                        newFormula = ExcelCellBase.UpdateFormulaReferences(ExcelCell.TranslateFromR1C1(formualR1C1, row-rows, col), rows, 0, startRow, 0);
                        currentFormulaR1C1 = ExcelRangeBase.TranslateToR1C1(newFormula, row, col);
                    }
                    if (currentFormulaR1C1 != prevFormualR1C1) //newFormula.Contains("#REF!"))
                    {
                        //if (refSplits == 0 || Regex.Split(newFormula, "#REF!").GetUpperBound(0) != refSplits)
                        //{
                        //isRef = true;
                        if (row == fromRow && col == fromCol)
                        {
                            formula.Formula = newFormula;
                        }
                        else
                        {
                            if (newFormulas.Count == newFormulasCount)
                            {
                                formula.Address = ExcelCellBase.GetAddress(formula.StartRow, formula.StartCol, row - 1, col);
                            }
                            else
                            {
                                newFormulas[newFormulas.Count - 1].Address = ExcelCellBase.GetAddress(newFormulas[newFormulas.Count - 1].StartRow, newFormulas[newFormulas.Count - 1].StartCol, row - 1, col);
                            }
                            var refFormula = new Formulas();
                            refFormula.Formula = newFormula;
                            refFormula.StartRow = row;
                            refFormula.StartCol = col;
                            newFormulas.Add(refFormula);

                            //restFormula = null;
                            prevFormualR1C1 = currentFormulaR1C1;
                        }
                    }
                    //    }
                    //    else
                    //    {
                    //        isRef = false;
                    //    }
                    //}
                    //else
                    //{
                    //    isRef = false;
                    //}
                    //if (restFormula==null)
                    //{
                        //if (newFormulas.Count == newFormulasCount)
                        //{
                        //    formula.Address = ExcelCellBase.GetAddress(formula.StartRow, formula.StartCol, row - 1, col);
                        //}
                        //else
                        //{
//                            newFormulas[newFormulas.Count - 1].Address = ExcelCellBase.GetAddress(newFormulas[newFormulas.Count - 1].StartRow, newFormulas[0].StartCol, row - 1, col);
                        //}

                        //restFormula = new Formulas();
                        //restFormula.Formula = newFormula;
                        //restFormula.StartRow = row;
                        //restFormula.StartCol = col;
                        //newFormulas.Add(restFormula);
                    //}
                }
            }
            if (rows < 0 && formula.StartRow > startRow)
            {
                if (formula.StartRow + rows < startRow)
                {
                    formula.StartRow = startRow;
                }
                else
                {
                    formula.StartRow += rows;
                }
            }
            if (newFormulas.Count > newFormulasCount)
            {
                newFormulas[newFormulas.Count - 1].Address = ExcelCellBase.GetAddress(newFormulas[newFormulas.Count - 1].StartRow, newFormulas[newFormulas.Count - 1].StartCol, toRow, toCol);
            }
        }
        #endregion

        #region DeleteRow
        /// <summary>
        /// Deletes the specified row from the worksheet.
        /// </summary>
        /// <param name="rowFrom">The number of the start row to be deleted</param>
        /// <param name="rows">Number of rows to delete</param>
        public void DeleteRow(int rowFrom, int rows)
        {
            ulong rowID = ExcelRow.GetRowID(SheetID, rowFrom);

            _cells.DeleteRows(rowID, rows, true);
            _rows.DeleteRows(rowID, rows, true);
            _formulaCells.DeleteRows(rowID, rows, false);
            if (_comments != null) _comments._comments.DeleteRows(rowID, rows, false);
            if (_vmlDrawings != null) _vmlDrawings._drawings.DeleteRows(rowID, rows, false);

            foreach (ExcelCell cell in _formulaCells)
            {
                cell._formula = ExcelCell.UpdateFormulaReferences(cell.Formula, -rows, 0, rowFrom, 0);
                cell._formulaR1C1 = "";
            }
            FixSharedFormulasRows(rowFrom, -rows);
            FixMergedCells(rowFrom, rows,true);
        }
        /// <summary>
        /// Deletes the specified row from the worksheet.
        /// </summary>
        /// <param name="rowFrom">The number of the start row to be deleted</param>
        /// <param name="rows">Number of rows to delete</param>
        /// <param name="shiftOtherRowsUp">Not used. Rows are always shifted</param>
        public void DeleteRow(int rowFrom, int rows, bool shiftOtherRowsUp)
		{
            if (shiftOtherRowsUp)
            {
                DeleteRow(rowFrom, rows);
            }
            else
            {
                ulong rowID = ExcelRow.GetRowID(SheetID, rowFrom);
                _cells.DeleteRows(rowID, rows, true);
                _rows.DeleteRows(rowID, rows, true);
                _formulaCells.DeleteRows(rowID, rows, false);
                if (_comments != null) _comments._comments.DeleteRows(rowID, rows, false);
                if (_vmlDrawings != null) _vmlDrawings._drawings.DeleteRows(rowID, rows, false);
            }
        }
		#endregion

        /// <summary>
        /// Get the cell value from thw worksheet
        /// </summary>
        /// <param name="Row">The row number</param>
        /// <param name="Column">The row number</param>
        /// <returns>The value</returns>
        public object GetValue(int Row, int Column)
        {
            ulong cellID = ExcelCell.GetCellID(SheetID, Row, Column);

            if (_cells.ContainsKey(cellID))
            {
                var cell = ((ExcelCell)_cells[cellID]);
                if (cell.IsRichText)
                {
                    return (object)Cells[Row, Column].RichText.Text;
                }
                else
                {
                    return cell.Value;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get a strongly typed cell value from the worksheet
        /// </summary>
        /// <typeparam name="T">The type</typeparam>
        /// <param name="Row">The row number</param>
        /// <param name="Column">The row number</param>
        /// <returns>The value. If the value can't be converted to the specified type, the default value will be returned</returns>
        public T GetValue<T>(int Row, int Column)
        {
             ulong cellID=ExcelCell.GetCellID(SheetID, Row, Column);
                        
            if (!_cells.ContainsKey(cellID))
            {
                return default(T);
            }

            var cell=((ExcelCell)_cells[cellID]);
            if (cell.IsRichText)
            {
                return (T)(object)Cells[Row, Column].RichText.Text;
            }
            else
            {
                return GetTypedValue<T>(cell.Value);
            }
        }
        //Thanks to Michael Tran for parts of this method
        internal T GetTypedValue<T>(object v)
        {
            if (v == null)
            {
                return default(T);
            }
            Type fromType = v.GetType();
            Type toType = typeof(T);
            if (fromType == toType)
            {
                return (T)v;
            }
            var cnv = TypeDescriptor.GetConverter(fromType);
            if (toType == typeof(DateTime))    //Handle dates
            {
                if (fromType == typeof(TimeSpan))
                {
                    return ((T)(object)(new DateTime(((TimeSpan)v).Ticks)));
                }
                else if (fromType == typeof(string))
                {
                    DateTime dt;
                    if (DateTime.TryParse(v.ToString(), out dt))
                    {
                        return (T)(object)(dt);
                    }
                    else
                    {
                        return default(T);
                    }

                }
                else
                {
                    if (cnv.CanConvertTo(typeof(double)))
                    {
                        return (T)(object)(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))));
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
            else if (toType == typeof(TimeSpan))    //Handle timespan
            {
                if (fromType == typeof(DateTime))
                {
                    return ((T)(object)(new TimeSpan(((DateTime)v).Ticks)));
                }
                else if (fromType == typeof(string))
                {
                    TimeSpan ts;
                    if (TimeSpan.TryParse(v.ToString(), out ts))
                    {
                        return (T)(object)(ts); 
                    }
                    else
                    {
                        return default(T);
                    }
                }
                else
                {
                    if (cnv.CanConvertTo(typeof(double)))
                    {

                        return (T)(object)(new TimeSpan(DateTime.FromOADate((double)cnv.ConvertTo(v, typeof(double))).Ticks));
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
            else
            {
                if (cnv.CanConvertTo(toType))
                {
                    return (T)cnv.ConvertTo(v, typeof(T));
                }
                else
                {
                    if (toType.IsGenericType && toType.GetGenericTypeDefinition().Equals(typeof(Nullable<>)))
                    {
                        toType = Nullable.GetUnderlyingType(toType);
                        if (cnv.CanConvertTo(toType))
                        {
                            return (T)cnv.ConvertTo(v, typeof(T));
                        }
                    }

                    if(fromType==typeof(double) && toType==typeof(decimal))
                    {
                        return (T)(object)Convert.ToDecimal(v);
                    }
                    else if (fromType == typeof(decimal) && toType == typeof(double))
                    {
                        return (T)(object)Convert.ToDouble(v);
                    }
                    else
                    {
                        return default(T);
                    }
                }
            }
        }
        /// <summary>
        /// Set the value of a cell
        /// </summary>
        /// <param name="Row">The row number</param>
        /// <param name="Column">The column number</param>
        /// <param name="Value">The value</param>
        public void SetValue(int Row, int Column, object Value)
        {
            Cell(Row, Column).Value = Value;
        }
        /// <summary>
        /// Set the value of a cell
        /// </summary>
        /// <param name="Address">The Excel address</param>
        /// <param name="Value">The value</param>
        public void SetValue(string Address, object Value)
        {
            int row, col;
            ExcelAddressBase.GetRowCol(Address, out row, out col, true);
            if (row < 1 || col < 1 || row > ExcelPackage.MaxRows && col > ExcelPackage.MaxColumns)
            {
                throw new ArgumentOutOfRangeException("Address is invalid or out of range");
            }
            Cell(row, col).Value = Value;
        }
		#endregion // END Worksheet Public Methods

		#region Worksheet Private Methods

		#region Worksheet Save
		/// <summary>
		/// Saves the worksheet to the package.
		/// </summary>
		internal void Save()  // Worksheet Save
		{
            DeletePrinterSettings();

			if (_worksheetXml != null)
			{
                
				// save the header & footer (if defined)
				if (_headerFooter != null)
					HeaderFooter.Save();

                if (_cells.Count > 0)
                {
                    this.SetXmlNodeString("d:dimension/@ref", Dimension.Address);
                }

                SaveComments();
                HeaderFooter.SaveHeaderFooterImages();
                SaveTables();
                SavePivotTables();
                SaveXml();
			}
            
            if (Drawings.UriDrawing!=null)
            {
                PackagePart partPack = Drawings.Part;
				Drawings.DrawingXml.Save(partPack.GetStream(FileMode.Create, FileAccess.Write));
                foreach (ExcelDrawing d in Drawings)
                {
                    if (d is ExcelChart)
                    {
                        ExcelChart c = (ExcelChart)d;
                        c.ChartXml.Save(c.Part.GetStream(FileMode.Create, FileAccess.Write));
                    }
                }   
            }
		}

        /// <summary>
        /// Delete the printersettings relationship and part.
        /// </summary>
        private void DeletePrinterSettings()
        {
            //Delete the relationship from the pageSetup tag
            XmlAttribute attr = (XmlAttribute)WorksheetXml.SelectSingleNode("//d:pageSetup/@r:id", NameSpaceManager);
            if (attr != null)
            {
                string relID = attr.Value;
                //First delete the attribute from the XML
                attr.OwnerElement.Attributes.Remove(attr);
                if(Part.RelationshipExists(relID))
                {
                    PackageRelationship rel = Part.GetRelationship(relID);
                    Uri printerSettingsUri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                    Part.DeleteRelationship(rel.Id);

                    //Delete the part from the package
                    if(_package.Package.PartExists(printerSettingsUri))
                    {
                        _package.Package.DeletePart(printerSettingsUri);
                    }
                }
            }
        }
        private void SaveComments()
        {
            if (_comments != null)
            {
                if (_comments.Count == 0)
                {
                    if (_comments.Uri != null)
                    {
                        Part.DeleteRelationship(_comments.RelId);
                        _package.Package.DeletePart(_comments.Uri);                        
                    }
                }
                else
                {
                    if (_comments.Uri == null)
                    {
                        _comments.Uri=new Uri(string.Format(@"/xl/comments{0}.xml", SheetID), UriKind.Relative);                        
                    }
                    if(_comments.Part==null)
                    {
                        _comments.Part = _package.Package.CreatePart(_comments.Uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", _package.Compression);
                        var rel = Part.CreateRelationship(PackUriHelper.GetRelativeUri(WorksheetUri, _comments.Uri), TargetMode.Internal, ExcelPackage.schemaRelationships+"/comments");
                    }
                    _comments.CommentXml.Save(_comments.Part.GetStream());
                }
            }

            if (_vmlDrawings != null)
            {
                if (_vmlDrawings.Count == 0)
                {
                    if (_vmlDrawings.Uri != null)
                    {
                        Part.DeleteRelationship(_vmlDrawings.RelId);
                        _package.Package.DeletePart(_vmlDrawings.Uri);
                    }
                }
                else
                {
                    if (_vmlDrawings.Uri == null)
                    {
                        _vmlDrawings.Uri = XmlHelper.GetNewUri(_package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
                    }
                    if (_vmlDrawings.Part == null)
                    {
                        _vmlDrawings.Part = _package.Package.CreatePart(_vmlDrawings.Uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _package.Compression);
                        var rel = Part.CreateRelationship(PackUriHelper.GetRelativeUri(WorksheetUri, _vmlDrawings.Uri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
                        SetXmlNodeString("d:legacyDrawing/@r:id", rel.Id);
                        _vmlDrawings.RelId = rel.Id;
                    }
                    _vmlDrawings.VmlDrawingXml.Save(_vmlDrawings.Part.GetStream());
                }
            }
        }
        /// <summary>
        /// Save all table data
        /// </summary>
        private void SaveTables()
        {
            foreach (var tbl in Tables)
            {
                if (tbl.ShowHeader || tbl.ShowTotal)
                {
                    int colNum = tbl.Address._fromCol;
                    foreach (var col in tbl.Columns)
                    {
                        if (tbl.ShowHeader)
                        {
                            Cell(tbl.Address._fromRow, colNum).Value = col.Name;
                        }
                        if (tbl.ShowTotal)
                        {
                            if (col.TotalsRowFunction == RowFunctions.Custom)
                            {
                                Cell(tbl.Address._toRow, colNum).Formula = col.TotalsRowFormula;
                            }
                            else if (col.TotalsRowFunction != RowFunctions.None)
                            {
                                switch (col.TotalsRowFunction)
                                {
                                    case RowFunctions.Average:
                                        Cell(tbl.Address._toRow, colNum).Formula = GetTotalFunction(col, "101");
                                        break;
                                    case RowFunctions.Count:
                                        Cell(tbl.Address._toRow, colNum).Formula = GetTotalFunction(col, "102");
                                        break;
                                    case RowFunctions.CountNums:
                                        Cell(tbl.Address._toRow, colNum).Formula = GetTotalFunction(col, "103");
                                        break;
                                    case RowFunctions.Max:
                                        Cell(tbl.Address._toRow, colNum).Formula = GetTotalFunction(col, "104");
                                        break;
                                    case RowFunctions.Min:
                                        Cell(tbl.Address._toRow, colNum).Formula = GetTotalFunction(col, "105");
                                        break;
                                    case RowFunctions.StdDev:
                                        Cell(tbl.Address._toRow, colNum).Formula = GetTotalFunction(col, "107");
                                        break;
                                    case RowFunctions.Var:
                                        Cell(tbl.Address._toRow, colNum).Formula = GetTotalFunction(col, "110");
                                        break;
                                    case RowFunctions.Sum:
                                        Cell(tbl.Address._toRow, colNum).Formula = GetTotalFunction(col, "109");
                                        break;
                                    default:
                                        throw (new Exception("Unknown RowFunction enum"));
                                }
                            }
                            else
                            {
                                Cell(tbl.Address._toRow, colNum).Value = col.TotalsRowLabel;
                            }
                        }
                        if (!string.IsNullOrEmpty(col.CalculatedColumnFormula))
                        {
                            int fromRow = tbl.ShowHeader ? tbl.Address._fromRow + 1 : tbl.Address._fromRow;
                            int toRow = tbl.ShowTotal ? tbl.Address._toRow - 1 : tbl.Address._toRow;
                            for (int row = fromRow; row <= toRow; row++)
                            {
                                Cell(row, colNum).Formula = col.CalculatedColumnFormula;
                            }                            
                        }
                        colNum++;
                    }
                }                
                if (tbl.Part == null)
                {
                    tbl.TableUri = GetNewUri(_package.Package, @"/xl/tables/table{0}.xml", tbl.Id);
                    tbl.Part = _package.Package.CreatePart(tbl.TableUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", Workbook._package.Compression);
                    var stream = tbl.Part.GetStream(FileMode.Create);
                    tbl.TableXml.Save(stream);
                    var rel = Part.CreateRelationship(PackUriHelper.GetRelativeUri(WorksheetUri, tbl.TableUri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");
                    tbl.RelationshipID = rel.Id;

                    CreateNode("d:tableParts");
                    XmlNode tbls = TopNode.SelectSingleNode("d:tableParts",NameSpaceManager);

                    var tblNode = tbls.OwnerDocument.CreateElement("tablePart",ExcelPackage.schemaMain);
                    tbls.AppendChild(tblNode);
                    tblNode.SetAttribute("id",ExcelPackage.schemaRelationships, rel.Id);
                }
                else
                {
                    var stream = tbl.Part.GetStream(FileMode.Create);
                    tbl.TableXml.Save(stream);
                }
            }
        }
        private void SavePivotTables()
        {
            foreach (var pt in PivotTables)
            {
                if (pt.DataFields.Count > 1)
                {
                    XmlElement parentNode;
                    if(pt.DataOnRows==true)
                    {
                        parentNode =  pt.PivotTableXml.SelectSingleNode("//d:rowFields", pt.NameSpaceManager) as XmlElement;
                        if (parentNode == null)
                        {
                            pt.CreateNode("d:rowFields");
                            parentNode = pt.PivotTableXml.SelectSingleNode("//d:rowFields", pt.NameSpaceManager) as XmlElement;
                        }
                    }
                    else
                    {
                        parentNode =  pt.PivotTableXml.SelectSingleNode("//d:colFields", pt.NameSpaceManager) as XmlElement;
                        if (parentNode == null)
                        {
                            pt.CreateNode("d:colFields");
                            parentNode = pt.PivotTableXml.SelectSingleNode("//d:colFields", pt.NameSpaceManager) as XmlElement;
                        }
                    }

                    if (parentNode.SelectSingleNode("d:field[@ x= \"-2\"]", pt.NameSpaceManager) == null)
                    {
                        XmlElement fieldNode = pt.PivotTableXml.CreateElement("field", ExcelPackage.schemaMain);
                        fieldNode.SetAttribute("x", "-2");
                        parentNode.AppendChild(fieldNode);
                    }
                }
                pt.PivotTableXml.Save(pt.Part.GetStream(FileMode.Create));
                pt.CacheDefinition.CacheDefinitionXml.Save(pt.CacheDefinition.Part.GetStream(FileMode.Create));
            }
        }
        private static string GetTotalFunction(ExcelTableColumn col,string FunctionNum)
        {
            return string.Format("SUBTOTAL({0},[{1}])", FunctionNum, col.Name);
        }
        private void SaveXml()
        {
            //Create the nodes if they do not exist.
            CreateNode("d:cols");
            CreateNode("d:sheetData");
            CreateNode("d:mergeCells");
            CreateNode("d:hyperlinks");
            CreateNode("d:rowBreaks");
            CreateNode("d:colBreaks");

            string xml = _worksheetXml.OuterXml;
            PackagePart partPack = _package.Package.GetPart(WorksheetUri);
            StreamWriter sw=new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));

            int colStart=0, colEnd=0;
            GetBlockPos(xml, "cols", ref colStart, ref colEnd);

            sw.Write(xml.Substring(0, colStart));
            var colBreaks = new List<int>();
            if (_columns.Count > 0)
            {
                UpdateColumnData(sw);
            }

            int cellStart = colEnd, cellEnd = colEnd;
            GetBlockPos(xml, "sheetData", ref cellStart, ref cellEnd);
            sw.Write(xml.Substring(colEnd, cellStart - colEnd));
            var rowBreaks=new List<int>();
            UpdateRowCellData(sw);

            int mergeStart = cellEnd, mergeEnd = cellEnd;

            GetBlockPos(xml, "mergeCells", ref mergeStart, ref mergeEnd);
            sw.Write(xml.Substring(cellEnd, mergeStart - cellEnd));

            if (_mergedCells.Count > 0)
            {
                UpdateMergedCells(sw);
            }


            int hyperStart = mergeEnd, hyperEnd = mergeEnd;
            GetBlockPos(xml, "hyperlinks", ref hyperStart, ref hyperEnd);
            sw.Write(xml.Substring(mergeEnd, hyperStart - mergeEnd));
            if (_hyperLinkCells.Count > 0)
            {
                UpdateHyperLinks(sw);
            }

            int rowBreakStart = hyperEnd, rowBreakEnd = hyperEnd;
            GetBlockPos(xml, "rowBreaks", ref rowBreakStart, ref rowBreakEnd);
            sw.Write(xml.Substring(hyperEnd, rowBreakStart - hyperEnd));
            //if (rowBreaks.Count > 0)
            //{
                UpdateRowBreaks(sw);
            //}

            int colBreakStart = rowBreakEnd, colBreakEnd = rowBreakEnd;
            GetBlockPos(xml, "colBreaks", ref colBreakStart, ref colBreakEnd);
            sw.Write(xml.Substring(rowBreakEnd, colBreakStart - rowBreakEnd));
            //if (colBreaks.Count > 0)
            //{
                UpdateColBreaks(sw);
            //}

            sw.Write(xml.Substring(colBreakEnd, xml.Length - colBreakEnd));
            sw.Flush();
        }
        private void UpdateColBreaks(StreamWriter sw)
        {
            StringBuilder breaks = new StringBuilder();
            int count = 0;
            foreach (ExcelColumn col in _columns)
            {
                if (col.PageBreak)
                {
                    breaks.AppendFormat("<brk id=\"{0}\" max=\"16383\" man=\"1\" />", col.ColumnMin);
                    count++;
                }
            }
            if (count > 0)
            {
                sw.Write(string.Format("<colBreaks count=\"{0}\" manualBreakCount=\"{0}\">{1}</colBreaks>", count, breaks.ToString()));
            }
        }

        private void UpdateRowBreaks(StreamWriter sw)
        {
            StringBuilder breaks=new StringBuilder();
            int count = 0;
            foreach(ExcelRow row in _rows)            
            {
                if (row.PageBreak)
                {
                    breaks.AppendFormat("<brk id=\"{0}\" max=\"1048575\" man=\"1\" />", row.Row);
                    count++;
                }
            }
            if (count>0)
            {
                sw.Write(string.Format("<rowBreaks count=\"{0}\" manualBreakCount=\"{0}\">{1}</rowBreaks>", count, breaks.ToString()));                
            }
        }
        /// <summary>
        /// Inserts the cols collection into the XML document
        /// </summary>
        private void UpdateColumnData(StreamWriter sw)
        {
            //ExcelColumn prevCol = null;   //commented out 11/1-12 JK 
            //foreach (ExcelColumn col in _columns)
            //{                
            //    if (prevCol != null)
            //    {
            //        if(prevCol.ColumnMax != col.ColumnMin-1)
            //        {
            //            prevCol._columnMax=col.ColumnMin-1;
            //        }
            //    }
            //    prevCol = col;
            //}
            sw.Write("<cols>");
            foreach (ExcelColumn col in _columns)
            {
                ExcelStyleCollection<ExcelXfs> cellXfs = _package.Workbook.Styles.CellXfs;

                sw.Write("<col min=\"{0}\" max=\"{1}\"", col.ColumnMin, col.ColumnMax);
                if (col.Hidden == true)
                {
                    //sbXml.Append(" width=\"0\" hidden=\"1\" customWidth=\"1\"");
                    sw.Write(" hidden=\"1\"");
                }
                else if (col.BestFit)
                {
                    sw.Write(" bestFit=\"1\"");
                }
                sw.Write(string.Format(CultureInfo.InvariantCulture, " width=\"{0}\" customWidth=\"1\"", col.Width));
                if (col.OutlineLevel > 0)
                {                    
                    sw.Write(" outlineLevel=\"{0}\" ", col.OutlineLevel);
                    if (col.Collapsed)
                    {
                        if (col.Hidden)
                        {
                            sw.Write(" collapsed=\"1\"");
                        }
                        else
                        {
                            sw.Write(" collapsed=\"1\" hidden=\"1\""); //Always hidden
                        }
                    }
                }
                if (col.Phonetic)
                {
                    sw.Write(" phonetic=\"1\"");
                }
                long styleID = col.StyleID >= 0 ? cellXfs[col.StyleID].newID : col.StyleID;
                if (styleID > 0)
                {
                    sw.Write(" style=\"{0}\"", styleID);
                }
                sw.Write(" />");

                //if (col.PageBreak)
                //{
                //    colBreaks.Add(col.ColumnMin);
                //}
            }
            sw.Write("</cols>");
        }
        /// <summary>
        /// Insert row and cells into the XML document
        /// </summary>
        private void UpdateRowCellData(StreamWriter sw)
        {
            ExcelStyleCollection<ExcelXfs> cellXfs = _package.Workbook.Styles.CellXfs;
            
            _hyperLinkCells = new List<ulong>();
            int row = -1;

            foreach (ExcelRow r in _rows)
            {
                int nextCell = ~_cells.IndexOf(r.RowID);
                if (nextCell >= _cells.Count || ((ExcelCell)_cells[nextCell]).Row!=r.Row)
                {
                    _cells.Add(r);
                }
            }

            StringBuilder sbXml = new StringBuilder();
            var ss = _package.Workbook._sharedStrings;
            sw.Write("<sheetData>");
            foreach (IRangeID r in _cells)
            {
                if (r is ExcelCell)
                {
                    ExcelCell cell = (ExcelCell)r;
                    long styleID = cell.StyleID >= 0 ? cellXfs[cell.StyleID].newID : cell.StyleID;

                    //Add the row element if it's a new row
                    if (row != cell.Row)
                    {
                        WriteRow(sw, cellXfs, row, cell.Row);
                        row = cell.Row;
                    }
                    if (cell.SharedFormulaID >= 0)
                    {
                        var f = _sharedFormulas[cell.SharedFormulaID];
                        if (f.Address.IndexOf(':') > 0)
                        {
                            if (f.StartCol == cell.Column && f.StartRow == cell.Row)
                            {
                                if (f.IsArray)
                                {
                                    sw.Write("<c r=\"{0}\" s=\"{1}\"><f ref=\"{2}\" t=\"array\">{3}</f></c>", cell.CellAddress, styleID < 0 ? 0 : styleID, f.Address, SecurityElement.Escape(f.Formula));
                                }
                                else
                                {
                                    sw.Write("<c r=\"{0}\" s=\"{1}\"><f ref=\"{2}\" t=\"shared\"  si=\"{3}\">{4}</f></c>", cell.CellAddress, styleID < 0 ? 0 : styleID, f.Address, cell.SharedFormulaID, SecurityElement.Escape(f.Formula));
                                }

                            }
                            else if (f.IsArray)
                            {
                                sw.Write("<c r=\"{0}\" s=\"{1}\" />", cell.CellAddress, styleID < 0 ? 0 : styleID);
                            }
                            else
                            {
                                sw.Write("<c r=\"{0}\" s=\"{1}\"><f t=\"shared\" si=\"{2}\" /></c>", cell.CellAddress, styleID < 0 ? 0 : styleID, cell.SharedFormulaID);
                            }
                        }
                        else
                        {
                            sw.Write("<c r=\"{0}\" s=\"{1}\">", f.Address, styleID < 0 ? 0 : styleID);
                            sw.Write("<f>{0}</f></c>", SecurityElement.Escape(f.Formula));
                        }
                    }
                    else if (cell.Formula != "")
                    {
                        sw.Write("<c r=\"{0}\" s=\"{1}\">", cell.CellAddress, styleID < 0 ? 0 : styleID);
                        sw.Write("<f>{0}</f></c>", SecurityElement.Escape(cell.Formula));
                    }
                    else
                    {
                        if (cell._value == null)
                        {
                            sw.Write("<c r=\"{0}\" s=\"{1}\" />", cell.CellAddress, styleID < 0 ? 0 : styleID);
                        }
                        else
                        {
                            if ((cell._value.GetType().IsPrimitive || cell._value is double || cell._value is decimal || cell._value is DateTime || cell._value is TimeSpan) && cell.DataType != "s")
                            {
                                string s;
                                try
                                {
                                    if (cell._value is DateTime)
                                    {
                                        s = ((DateTime)cell.Value).ToOADate().ToString(CultureInfo.InvariantCulture);
                                    }
                                    else if (cell._value is TimeSpan)
                                    {
                                        s = new DateTime(((TimeSpan)cell.Value).Ticks).ToOADate().ToString(CultureInfo.InvariantCulture); ;
                                    }
                                    else
                                    {
                                        if (cell._value is double && double.IsNaN((double)cell._value))
                                        {
                                            s = "0";
                                        }
                                        else
                                        {
                                            s = Convert.ToDouble(cell._value, CultureInfo.InvariantCulture).ToString("g15", CultureInfo.InvariantCulture);
                                        }
                                    }
                                }

                                catch
                                {
                                    s = "0";
                                }
                                if (cell._value is bool)
                                {
                                    sw.Write("<c r=\"{0}\" s=\"{1}\" t=\"b\">", cell.CellAddress, styleID < 0 ? 0 : styleID);
                                }
                                else
                                {
                                    sw.Write("<c r=\"{0}\" s=\"{1}\">", cell.CellAddress, styleID < 0 ? 0 : styleID);
                                }
                                sw.Write("<v>{0}</v></c>", s);
                            }
                            else
                            {
                                int ix;
                                if (!ss.ContainsKey(cell._value.ToString()))
                                {
                                    ix = ss.Count;
                                    ss.Add(cell._value.ToString(), new ExcelWorkbook.SharedStringItem() { isRichText = cell.IsRichText, pos = ix });
                                }
                                else
                                {
                                    ix = ss[cell.Value.ToString()].pos;
                                }
                                sw.Write("<c r=\"{0}\" s=\"{1}\" t=\"s\">", cell.CellAddress, styleID < 0 ? 0 : styleID);
                                sw.Write("<v>{0}</v></c>", ix);
                            }
                        }
                    }
                    //Update hyperlinks.
                    if (cell.Hyperlink != null)
                    {
                        _hyperLinkCells.Add(cell.CellID);
                    }
                }
                else  //ExcelRow
                {
                    int newRow=((ExcelRow)r).Row;
                    WriteRow(sw, cellXfs, row, newRow);
                    row = newRow;
                }
            }

            if (row != -1) sw.Write("</row>");
            sw.Write("</sheetData>");
        }

        private void WriteRow(StreamWriter sw, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
        {
            if (prevRow != -1) sw.Write("</row>");
            ulong rowID = ExcelRow.GetRowID(SheetID, row);
            sw.Write("<row r=\"{0}\" ", row);
            if (_rows.ContainsKey(rowID))
            {
                ExcelRow currRow = _rows[rowID] as ExcelRow;
                if (currRow.Hidden == true)
                {
                    sw.Write("ht=\"0\" hidden=\"1\" ");
                }
                else if (currRow.Height != DefaultRowHeight)
                {
                    sw.Write(string.Format(CultureInfo.InvariantCulture, "ht=\"{0}\" ", currRow.Height));
                    if (currRow.CustomHeight)
                    {
                        sw.Write("customHeight=\"1\" ");
                    }
                }

                if (currRow.StyleID > 0)
                {
                    sw.Write("s=\"{0}\" customFormat=\"1\" ", cellXfs[currRow.StyleID].newID);
                }
                if (currRow.OutlineLevel > 0)
                {
                    sw.Write("outlineLevel =\"{0}\" ", currRow.OutlineLevel);
                    if (currRow.Collapsed)
                    {
                        if (currRow.Hidden)
                        {
                            sw.Write(" collapsed=\"1\"");
                        }
                        else
                        {
                            sw.Write(" collapsed=\"1\" hidden=\"1\""); //Always hidden
                        }
                    }
                }
                if (currRow.Phonetic)
                {
                    sw.Write("ph=\"1\" ");
                }
            }
            sw.Write(">");
        }

        /// <summary>
        /// Update xml with hyperlinks 
        /// </summary>
        /// <param name="sw">The stream</param>
        private void UpdateHyperLinks(StreamWriter sw)
        {
                sw.Write("<hyperlinks>");
                Dictionary<string, string> hyps = new Dictionary<string, string>();
                foreach (ulong cellId in _hyperLinkCells)
                {
                    ExcelCell cell = _cells[cellId] as ExcelCell;
                    if (cell.Hyperlink is ExcelHyperLink && !string.IsNullOrEmpty((cell.Hyperlink as ExcelHyperLink).ReferenceAddress))
                    {
                        ExcelHyperLink hl = cell.Hyperlink as ExcelHyperLink;
                        sw.Write("<hyperlink ref=\"{0}\" location=\"{1}\" display=\"{2}\" {3}/>", 
                                Cells[cell.Row, cell.Column, cell.Row+hl.RowSpann, cell.Column+hl.ColSpann].Address, 
                                ExcelCell.GetFullAddress(Name, hl.ReferenceAddress),
                                SecurityElement.Escape(hl.Display),
                                string.IsNullOrEmpty(hl.ToolTip) ? "" : "tooltip=\"" + SecurityElement.Escape(hl.Display) + "\"");
                    }
                    else
                    {
                        string id;
                        if (hyps.ContainsKey(cell.Hyperlink.AbsoluteUri))
                        {
                            id = hyps[cell.Hyperlink.AbsoluteUri];
                        }
                        else
                        {
                            PackageRelationship relationship = Part.CreateRelationship(cell.Hyperlink, TargetMode.External, ExcelPackage.schemaHyperlink);
                            if (cell.Hyperlink is ExcelHyperLink && !string.IsNullOrEmpty((cell.Hyperlink as ExcelHyperLink).Display))
                            {
                                ExcelHyperLink hl = cell.Hyperlink as ExcelHyperLink;
                                sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\" {2}/>",cell.CellAddress, relationship.Id,                                
                                    string.IsNullOrEmpty(hl.ToolTip) ? "" : "@tooltip=\"" + SecurityElement.Escape(hl.Display) + "\"");
                            }
                            else
                            {
                                sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\" />",cell.CellAddress, relationship.Id);
                            }
                            id = relationship.Id;
                        }
                        cell.HyperLinkRId = id;
                    }
                }
                sw.Write("</hyperlinks>");
        }
        /// <summary>
        /// Create the hyperlinks node in the XML
        /// </summary>
        /// <returns></returns>
        private XmlNode CreateHyperLinkCollection()
        {
            XmlElement hl=_worksheetXml.CreateElement("hyperlinks",ExcelPackage.schemaMain);
            XmlNode prevNode = _worksheetXml.SelectSingleNode("//d:conditionalFormatting", NameSpaceManager);
            if (prevNode == null)
            {
                prevNode = _worksheetXml.SelectSingleNode("//d:mergeCells", NameSpaceManager);
                if (prevNode == null)
                {
                    prevNode = _worksheetXml.SelectSingleNode("//d:sheetData", NameSpaceManager);
                }
            }
            return _worksheetXml.DocumentElement.InsertAfter(hl, prevNode);
        }
        /// <summary>
        /// Dimension address for the worksheet. 
        /// Top left cell to Bottom right.
        /// If the worksheet has no cells, null is returned
        /// </summary>
        public ExcelAddressBase Dimension
        {
            get
            {
                if (_cells.Count > 0)
                {
                    ExcelAddressBase addr = new ExcelAddressBase((_cells[0] as ExcelCell).Row, _minCol, (_cells[_cells.Count - 1] as ExcelCell).Row, _maxCol);
                    addr._ws = Name;
                    return addr;
                }
                else
                {
                    return null;
                }
            }
        }
        ExcelSheetProtection _protection=null;
        /// <summary>
        /// Access to sheet protection properties
        /// </summary>
        public ExcelSheetProtection Protection
        {
            get
            {
                if (_protection == null)
                {
                    _protection = new ExcelSheetProtection(NameSpaceManager, TopNode, this);
                }
                return _protection;
            }
        }
        #region Drawing
        ExcelDrawings _drawings = null;
        /// <summary>
        /// Collection of drawing-objects like shapes, images and charts
        /// </summary>
        public ExcelDrawings Drawings
        {
            get
            {
                if (_drawings == null)
                {
                    _drawings = new ExcelDrawings(_package, this);
                }
                return _drawings;
            }
        }
        #endregion
        ExcelTableCollection _tables = null;
        /// <summary>
        /// Tables defined in the worksheet.
        /// </summary>
        public ExcelTableCollection Tables
        {
            get
            {
                if (_tables == null)
                {
                    _tables = new ExcelTableCollection(this);
                }
                return _tables;
            }
        }
        ExcelPivotTableCollection _pivotTables = null;
        /// <summary>
        /// Pivottables defined in the worksheet.
        /// </summary>
        public ExcelPivotTableCollection PivotTables
        {
            get
            {
                if (_pivotTables == null)
                {
                    _pivotTables = new ExcelPivotTableCollection(this);
                }
                return _pivotTables;
            }
        }
        private ExcelDataValidationCollection _dataValidation = null;
        /// <summary>
        /// DataValidation defined in the worksheet. Use the Add methods to create DataValidations and add them to the worksheet. Then
        /// set the properties on the instance returned.
        /// </summary>
        /// <seealso cref="ExcelDataValidationCollection"/>
        public ExcelDataValidationCollection DataValidations
        {
            get
            {
                if (_dataValidation == null)
                {
                    _dataValidation = new ExcelDataValidationCollection(this);
                }
                return _dataValidation;
            }
        }
        ExcelBackgroundImage _backgroundImage = null;
        /// <summary>
        /// An image displayed as the background of the worksheet.
        /// </summary>
        public ExcelBackgroundImage BackgroundImage
        {
            get
            {
                if (_backgroundImage == null)
                {
                    _backgroundImage = new ExcelBackgroundImage(NameSpaceManager, TopNode, this);
                }
                return _backgroundImage;
            }
        }
        /// <summary>
		/// Returns the style ID given a style name.  
		/// The style ID will be created if not found, but only if the style name exists!
		/// </summary>
		/// <param name="StyleName"></param>
		/// <returns></returns>
		internal int GetStyleID(string StyleName)
		{
			ExcelNamedStyleXml namedStyle=null;
            Workbook.Styles.NamedStyles.FindByID(StyleName, ref namedStyle);
            if (namedStyle.XfId == int.MinValue)
            {
                namedStyle.XfId=Workbook.Styles.CellXfs.FindIndexByID(namedStyle.Style.Id);
            }
            return namedStyle.XfId;
		}
        /// <summary>
        /// The workbook object
        /// </summary>
        public ExcelWorkbook Workbook
        {
            get
            {
                return _package.Workbook;
            }
        }
		#endregion
        #endregion  // END Worksheet Private Methods

        /// <summary>
        /// Get the next ID from a shared formula or an Array formula
        /// Sharedforumlas will have an id from 0-x. Array formula ids start from 0x4000001-. 
        /// </summary>
        /// <param name="isArray">If the formula is an array formula</param>
        /// <returns></returns>
        internal int GetMaxShareFunctionIndex(bool isArray)
        {
            int i=_sharedFormulas.Count + 1;
            if (isArray)
                i |= 0x40000000;

            while(_sharedFormulas.ContainsKey(i))
            {
                i++;
            }
            return i;
        }
        internal void SetHFLegacyDrawingRel(string relID)
        {
            SetXmlNodeString("d:legacyDrawingHF/@r:id", relID);
        }

        #region Formulas
        Dictionary<string, string> ICalcEngineFormulaInfo.GetFormulas()
        {
            Dictionary<string, string> fs = new Dictionary<string, string>();

            //Single Cell Formulas
            foreach (var r in _formulaCells)
            {
                var f = (ExcelCell)r;
                fs.Add(f.CellAddress, f._formula);
            }

            //Shared Formulas
            foreach (var sf in _sharedFormulas.Values)
            {
                fs.Add(sf.Address, sf.Formula);
            }

            //Name formulas
            foreach (var n in _names)
            {
                if (!string.IsNullOrEmpty(n.NameFormula))
                {
                    fs.Add(n.Name, n.NameFormula);
                }
            }
            return fs;
        }



        Dictionary<string, object> ICalcEngineFormulaInfo.GetNameValues()
        {
            Dictionary<string, object> nv = new Dictionary<string, object>();
            //Name formulas
            foreach (var n in _names)
            {
                if (string.IsNullOrEmpty(n.NameFormula))
                {
                    nv.Add(n.Name, n.Value);
                }
            }
            return nv;
        }

        object ICalcEngineValueInfo.GetValue(int row, int col)
        {
            return ((ExcelCell)_cells[ExcelCell.GetCellID(SheetID, row, col)]).Value;
        }



        bool ICalcEngineValueInfo.IsHidden(int row, int col)
        {
            var colID = ExcelColumn.GetColumnID(_sheetID, col);
            if (_columns.ContainsKey(colID))
            {
                if (((ExcelColumn)_columns[col]).Width == 0)
                {
                    return true;
                }
            }
            var rowID = ExcelRow.GetRowID(_sheetID, row);
            if (_rows.ContainsKey(rowID))
            {
                if (((ExcelRow)_rows[rowID]).Height == 0)
                {
                    return true;
                }
            }
            return false;
        }



        void ICalcEngineValueInfo.SetFormulaValue(int row, int col, object value)
        {
            ((ExcelCell)_cells[ExcelCell.GetCellID(SheetID, row, col)])._value = value;
        }
        #endregion
    }  // END class Worksheet
}
