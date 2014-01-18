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
 * Jan Källman		    Initial Release		        2011-11-02
 * Jan Källman          Total rewrite               2010-03-01
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using System;
using System.Xml;
using System.Collections.Generic;
using System.IO;
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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Utils;
using Ionic.Zip;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;
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
    [Flags]
    internal enum CellFlags
    {
        Merged = 0x1,
        RichText = 0x2,
        SharedFormula = 0x4,
        ArrayFormula = 0x8
    }
    /// <summary>
    /// Represents an Excel Chartsheet and provides access to its properties and methods
    /// </summary>
    public class ExcelChartsheet : ExcelWorksheet
    {
        //ExcelDrawings draws;
        public ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hidden, eChartType chartType) :
            base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
        {
            this.Drawings.AddChart("Chart 1", chartType);
        }
        public ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hidden) :
            base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
        {
        }
        public ExcelChart Chart 
        {
            get
            {
                return (ExcelChart)Drawings[0];
            }
        }
    }
    /// <summary>
	/// Represents an Excel worksheet and provides access to its properties and methods
	/// </summary>
    public class ExcelWorksheet : XmlHelper, IDisposable
	{
        internal class Formulas
        {
            public Formulas(ISourceCodeTokenizer tokenizer)
            {
                _tokenizer = tokenizer;
            }

            private ISourceCodeTokenizer _tokenizer;
            internal int Index { get; set; }
            internal string Address { get; set; }
            internal bool IsArray { get; set; }
            public string Formula { get; set; }
            public int StartRow { get; set; }
            public int StartCol { get; set; }

            private IEnumerable<Token> Tokens {get; set;}

            internal string GetFormula(int row, int column)
            {
                if (StartRow == row && StartCol == column)
                {
                    return Formula;
                }

                if (Tokens == null)
                {
                    Tokens = _tokenizer.Tokenize(Formula);
                }
                
                string f = "";
                foreach (var token in Tokens)
                {
                    if (token.TokenType == TokenType.ExcelAddress)
                    {
                        var a = new ExcelFormulaAddress(token.Value);
                        f += a.GetOffset(row - StartRow, column - StartCol);
                    }
                    else
                    {
                        f += token.Value;
                    }
                }
                return f;
            }
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
            internal void Remove(T Item)
            {
                _list.Remove(Item);
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

        internal CellStore<object> _values;
        internal CellStore<string> _types;
        internal CellStore<int> _styles;
        internal CellStore<object> _formulas;
        internal FlagCellStore _flags;
        internal CellStore<List<Token>> _formulaTokens;

        internal CellStore<Uri> _hyperLinks;
        internal CellStore<ExcelComment> _commentsStore;

        internal Dictionary<int, Formulas> _sharedFormulas = new Dictionary<int, Formulas>();
        internal int _minCol = ExcelPackage.MaxColumns;
        internal int _maxCol = 0;
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
            SchemaNodeOrder = new string[] { "sheetPr", "tabColor", "outlinePr", "pageSetUpPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "protectedRanges","scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "linePrint", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "activeXControls", "webPublishItems", "tableParts" , "extLst" };
            _package = excelPackage;   
            _relationshipID = relID;
            _worksheetUri = uriWorksheet;
            _name = sheetName;
            _sheetID = sheetID;
            _positionID = positionID;
            Hidden = hide;
           
            /**** Cellstore ****/
            _values=new CellStore<object>();
            _types = new CellStore<string>();
            _styles = new CellStore<int>();
            _formulas = new CellStore<object>();
            _flags = new FlagCellStore();
            _commentsStore = new CellStore<ExcelComment>();
            _hyperLinks = new CellStore<Uri>();
            
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
        /// The Zip.ZipPackagePart for the worksheet within the package
        /// </summary>
        internal Packaging.ZipPackagePart Part { get { return (_package.Package.GetPart(WorksheetUri)); } }
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
                CheckSheetType();
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
                CheckSheetType();
                SetXmlNodeString("d:autoFilter/@ref", value.Address);
            }
        }

        internal void CheckSheetType()
        {
            if (this is ExcelChartsheet)
            {
                throw (new NotSupportedException("This property or method is not supported for a Chartsheet"));
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
                value=_package.Workbook.Worksheets.ValidateFixSheetName(value);
                _package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@name", _sheetID), value);
                ChangeNames(value);

                _name = value;
            }
		}

        private void ChangeNames(string value)
        {
            //Renames name in this Worksheet;
            foreach (var n in Workbook.Names)
            {
                if (string.IsNullOrEmpty(n.NameFormula) && n.NameValue==null)
                {
                    n.ChangeWorksheet(_name, value);
                }
            }
            foreach (var ws in Workbook.Worksheets)
            {
                if (!(ws is ExcelChartsheet))
                {
                    foreach (var n in ws.Names)
                    {
                        if (string.IsNullOrEmpty(n.NameFormula) && n.NameValue == null)
                        {
                            n.ChangeWorksheet(_name, value);
                        }
                    }
                }
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
                CheckSheetType();
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
                CheckSheetType();
                if (double.IsNaN(_defaultRowHeight))
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
                CheckSheetType();
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
                CheckSheetType();
                double ret = GetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth");
                if (double.IsNaN(ret))
                {
                    ret = 9.140625; // Excel's default width
                }
                return ret;
            }
            set
            {
                CheckSheetType();
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
                CheckSheetType();
                return GetXmlNodeBool(outLineSummaryBelowPath);
            }
            set
            {
                CheckSheetType();
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
                CheckSheetType();
                return GetXmlNodeBool(outLineSummaryRightPath);
            }
            set
            {
                CheckSheetType();
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
                CheckSheetType();
                return GetXmlNodeBool(outLineApplyStylePath);
            }
            set
            {
                CheckSheetType();
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
        const string codeModuleNamePath = "d:sheetPr/@codeName";
        internal string CodeModuleName
        {
            get
            {
                return GetXmlNodeString(codeModuleNamePath);
            }
            set
            {
                SetXmlNodeString(codeModuleNamePath, value);
            }
        }
        internal void CodeNameChange(string value)
        {
            CodeModuleName = value;
        }
        public VBA.ExcelVBAModule CodeModule
        {
            get
            {
                if (_package.Workbook.VbaProject != null)
                {
                    return _package.Workbook.VbaProject.Modules[CodeModuleName];
                }
                else
                {
                    return null;
                }
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
                CheckSheetType();
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
                    var vmlUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

                    _vmlDrawings = new ExcelVmlDrawingCommentCollection(_package, this, vmlUri);
                    _vmlDrawings.RelId = rel.Id;
                }
            }
        }

        private void CreateXml()
        {
            _worksheetXml = new XmlDocument();
            _worksheetXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
            Packaging.ZipPackagePart packPart = _package.Package.GetPart(WorksheetUri);
            string xml = "";

            // First Columns, rows, cells, mergecells, hyperlinks and pagebreakes are loaded from a xmlstream to optimize speed...
            bool doAdjust = _package.DoAdjustDrawings;
            _package.DoAdjustDrawings = false;
            Stream stream = packPart.GetStream();

            XmlTextReader xr = new XmlTextReader(stream);
            xr.ProhibitDtd = true;
            
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
            Encoding encoding;
            xml = GetWorkSheetXml(stream, start, end, out encoding);

            //first char is invalid sometimes?? 
            if (xml[0] != '<')
                LoadXmlSafe(_worksheetXml, xml.Substring(1, xml.Length - 1), encoding);
            else
                LoadXmlSafe(_worksheetXml, xml, encoding);

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
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        int id;
                        if (int.TryParse(xr.GetAttribute("id"), out id))
                        {
                            Row(id).PageBreak = true;
                        }   
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
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        int id;
                        if (int.TryParse(xr.GetAttribute("id"), out id))
                        {
                            Column(id).PageBreak = true;
                        }
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
        private string GetWorkSheetXml(Stream stream, long start, long end, out Encoding encoding)
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
                sb.Append(block,0,pos);
                length += size;
            }
            while (length < start + 20 && length < end);
            startmMatch = Regex.Match(sb.ToString(), string.Format("(<[^>]*{0}[^>]*>)", "sheetData"));
            if (!startmMatch.Success) //Not found
            {
                encoding = sr.CurrentEncoding;
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
                            sb.Append(block, 0, pos);
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
                
                encoding = sr.CurrentEncoding;
                return xml;
            }
        }
        private void GetBlockPos(string xml, string tag, ref int start, ref int end)
        {
            Match startmMatch, endMatch;
            startmMatch = Regex.Match(xml.Substring(start), string.Format("(<[^>]*{0}[^>]*>)", tag)); //"<[a-zA-Z:]*" + tag + "[?]*>");

            if (!startmMatch.Success) //Not found
            {
                start = -1;
                end = -1;
                return;
            }
            var startPos=startmMatch.Index+start;
            if(startmMatch.Value.Substring(startmMatch.Value.Length-2,1)=="/")
            {
                end = startPos + startmMatch.Length;
            }
            else
            {
                endMatch = Regex.Match(xml.Substring(start), string.Format("(</[^>]*{0}[^>]*>)", tag));
                if (endMatch.Success)
                {
                    end = endMatch.Index + endMatch.Length + start;
                }
            }
            start = startPos;
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
                    if (xr.LocalName != "col") break;
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        int min = int.Parse(xr.GetAttribute("min"));

                        ExcelColumn col = new ExcelColumn(this, min);

                        col.ColumnMax = int.Parse(xr.GetAttribute("max"));
                        col.Width = xr.GetAttribute("width") == null ? 0 : double.Parse(xr.GetAttribute("width"), CultureInfo.InvariantCulture);
                            col.BestFit = xr.GetAttribute("bestFit") != null && xr.GetAttribute("bestFit") == "1" ? true : false;
                            col.Collapsed = xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed") == "1" ? true : false;
                            col.Phonetic = xr.GetAttribute("phonetic") != null && xr.GetAttribute("phonetic") == "1" ? true : false;
                        col.OutlineLevel = (short)(xr.GetAttribute("outlineLevel") == null ? 0 : int.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture));
                            col.Hidden = xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1" ? true : false;
                        _values.SetValue(0, min, col);
                    
                        int style;
                        if (!(xr.GetAttribute("style") == null || !int.TryParse(xr.GetAttribute("style"), out style)))
                        {
                            _styles.SetValue(0, min, style);
                        }
                    }
                }
            }
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
            if (!ReadUntil(xr, "hyperlinks", "rowBreaks", "colBreaks")) return;
            while (xr.Read())
            {
                if (xr.LocalName == "hyperlink")
                {
                    int fromRow, fromCol, toRow, toCol;
                    ExcelCellBase.GetRowColFromAddress(xr.GetAttribute("ref"), out fromRow, out fromCol, out toRow, out toCol);
                    ExcelHyperLink hl = null;
                    if (xr.GetAttribute("id", ExcelPackage.schemaRelationships) != null)
                    {
                        var rId = xr.GetAttribute("id", ExcelPackage.schemaRelationships);
                        var uri = Part.GetRelationship(rId).TargetUri;
                        if (uri.IsAbsoluteUri)
                        {
                            hl = new ExcelHyperLink(uri.AbsoluteUri);
                        }
                        else
                        {
                            hl = new ExcelHyperLink(uri.OriginalString, UriKind.Relative);
                        }
                        hl.RId = rId;
                        Part.DeleteRelationship(rId); //Delete the relationship, it is recreated when we save the package.
                    }
                    else if (xr.GetAttribute("location") != null)
                    {
                        hl = new ExcelHyperLink(xr.GetAttribute("location"), xr.GetAttribute("display"));
                        hl.RowSpann = toRow - fromRow;
                        hl.ColSpann = toCol - fromCol;
                    }

                    string tt = xr.GetAttribute("tooltip");
                    if (!string.IsNullOrEmpty(tt))
                    {
                        hl.ToolTip = tt;
                    }
                    _hyperLinks.SetValue(fromRow, fromCol, hl);
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
            //var cellList=new List<IRangeID>();
            //var rowList = new List<IRangeID>();
            //var formulaList = new List<IRangeID>();
            ReadUntil(xr, "sheetData", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks");
            ExcelAddressBase address=null;
            string type="";
            int style=0;
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
                        //rowList.Add(AddRow(xr, row));
                        _values.SetValue(row, 0, AddRow(xr, row));
                        if(xr.GetAttribute("s") != null)
                        {
                            _styles.SetValue(row, 0, int.Parse(xr.GetAttribute("s"), CultureInfo.InvariantCulture));
                        }
                    }
                    xr.Read();
                }
                else if (xr.LocalName == "c")
                {
                    //if (cell != null) cellList.Add(cell);
                    //cell = new ExcelCell(this, xr.GetAttribute("r"));
                    address=new ExcelAddressBase(xr.GetAttribute("r"));                    
                    
                    //Datetype
                    if (xr.GetAttribute("t") != null)
                    {
                        type=xr.GetAttribute("t");
                        _types.SetValue(address._fromRow, address._fromCol, type); 
                    }
                    else
                    {
                        type="";
                    }
                    //Style
                    if(xr.GetAttribute("s") != null)
                    {
                        style=int.Parse(xr.GetAttribute("s"));
                        _styles.SetValue(address._fromRow, address._fromCol, style);
                        _values.SetValue(address._fromRow, address._fromCol, null); //TODO:Better Performance ??
                    }
                    else
                    {
                        style = 0;
                    }
                    xr.Read();
                }
                else if (xr.LocalName == "v")
                {
                    SetValueFromXml(xr, type, style, address._fromRow, address._fromCol);
                    
                    xr.Read();
                }
                else if (xr.LocalName == "f")
                {
                    string t = xr.GetAttribute("t");
                    if (t == null)
                    {
                        _formulas.SetValue(address._fromRow, address._fromCol, xr.ReadElementContentAsString());
                        _values.SetValue(address._fromRow, address._fromCol, null);
                        //formulaList.Add(cell);
                    }
                    else if (t == "shared")
                    {

                        string si = xr.GetAttribute("si");
                        if (si != null)
                        {
                            var sfIndex = int.Parse(si);
                            _formulas.SetValue(address._fromRow, address._fromCol, sfIndex);
                            _values.SetValue(address._fromRow, address._fromCol, null);
                            string fAddress = xr.GetAttribute("ref");
                            string formula = xr.ReadElementContentAsString();
                            if (formula != "")
                            {
                                _sharedFormulas.Add(sfIndex, new Formulas(SourceCodeTokenizer.Default) { Index = sfIndex, Formula = formula, Address = fAddress, StartRow = address._fromRow, StartCol = address._fromCol });
                            }
                        }
                        else
                        {
                            xr.Read();  //Something is wrong in the sheet, read next
                        }
                    }
                    else if (t == "array") //TODO: Array functions are not support yet. Read the formula for the start cell only.
                    {
                        string aAddress = xr.GetAttribute("ref");
                        string formula = xr.ReadElementContentAsString();
                        var afIndex = GetMaxShareFunctionIndex(true);
                        _formulas.SetValue(address._fromRow, address._fromCol, afIndex.ToString());
                        _values.SetValue(address._fromRow, address._fromCol, null);
                        _sharedFormulas.Add(afIndex, new Formulas(SourceCodeTokenizer.Default) { Index = afIndex, Formula = formula, Address = aAddress, StartRow = address._fromRow, StartCol = address._fromCol, IsArray = true });
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
                        _values.SetValue(address._fromRow, address._fromCol, xr.ReadInnerXml());
                        //cell._value = xr.ReadInnerXml();
                    }
                    else
                    {
                        _values.SetValue(address._fromRow, address._fromCol, xr.ReadOuterXml());
                        _types.SetValue(address._fromRow, address._fromCol, "rt");
                        _flags.SetFlagValue(address._fromRow, address._fromCol, true, CellFlags.RichText);
                        //cell.IsRichText = true;
                    }
                }
                else
                {
                    break;
                }
            }
            //if (cell != null) cellList.Add(cell);

            //_cells = new RangeCollection(cellList);
            //_rows = new RangeCollection(rowList);
            //_formulaCells = new RangeCollection(formulaList);
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
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        string address = xr.GetAttribute("ref");
                        int fromRow, fromCol, toRow, toCol;
                        ExcelCellBase.GetRowColFromAddress(address, out fromRow, out fromCol, out toRow, out toCol);
                        for (int row = fromRow; row <= toRow; row++)
                        {
                            for (int col = fromCol; col <= toCol; col++)
                            {
                            _flags.SetFlagValue(row, col, true,CellFlags.Merged);
                            }
                        }
                    _mergedCells.List.Add(address);
                    }
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
        private RowInternal AddRow(XmlTextReader xr, int row)
        {
            return new RowInternal()
            {
                Collapsed=(xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed")== "1" ? true : false),
                Height = (xr.GetAttribute("ht") == null ? -1 : double.Parse(xr.GetAttribute("ht"), CultureInfo.InvariantCulture)),
                Hidden = (xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1" ? true : false),
                Phonetic = xr.GetAttribute("ph") != null && xr.GetAttribute("ph") == "1" ? true : false,
                CustomHeight = xr.GetAttribute("customHeight") == null ? false : xr.GetAttribute("customHeight")=="1"
            };
        }

        private void SetValueFromXml(XmlTextReader xr, string type, int styleID, int row, int col)
        {
            //XmlNode vnode = colNode.SelectSingleNode("d:v", NameSpaceManager);
            //if (vnode == null) return null;
            if (type == "s")
            {
                int ix = xr.ReadElementContentAsInt();
                _values.SetValue(row, col, _package.Workbook._sharedStringsList[ix].Text);
                if (_package.Workbook._sharedStringsList[ix].isRichText)
                {
                    _flags.SetFlagValue(row, col, true, CellFlags.RichText);
                }                
            }
            else if (type == "str")
            {
                _values.SetValue(row, col, xr.ReadElementContentAsString());
            }
            else if (type == "b")
            {
                _values.SetValue(row, col, (xr.ReadElementContentAsString()!="0"));
            }
            else if (type == "e")
            {
                _values.SetValue(row, col, GetErrorType(xr.ReadElementContentAsString()));
            }
            else
            {
                string v = xr.ReadElementContentAsString();
                var nf = Workbook.Styles.CellXfs[styleID].NumberFormatId;
                if ((nf >= 14 && nf <= 22) || (nf >= 45 && nf <= 47))
                {
                    double res;
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out res))
                    {
                        _values.SetValue(row, col, DateTime.FromOADate(res));
                    }
                    else
                    {
                        _values.SetValue(row, col, "");
                    }
                }
                else
                {
                    double d;
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    {
                        _values.SetValue(row, col, d);
                    }
                    else
                    {
                        _values.SetValue(row, col, double.NaN);
                    }
                }
            }
        }

        private object GetErrorType(string v)
        {
            return ExcelErrorValue.Parse(v.ToUpper());
            //switch(v.ToUpper())
            //{
            //    case "#DIV/0!":
            //        return new ExcelErrorValue.cre(eErrorType.Div0);
            //    case "#REF!":
            //        return new ExcelErrorValue(eErrorType.Ref);
            //    case "#N/A":
            //        return new ExcelErrorValue(eErrorType.NA);
            //    case "#NAME?":
            //        return new ExcelErrorValue(eErrorType.Name);
            //    case "#NULL!":
            //        return new ExcelErrorValue(eErrorType.Null);
            //    case "#NUM!":
            //        return new ExcelErrorValue(eErrorType.Num);
            //    default:
            //        return new ExcelErrorValue(eErrorType.Value);
            //}
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
        
        ///// <summary>
        ///// Provides access to an individual cell within the worksheet.
        ///// </summary>
        ///// <param name="row">The row number in the worksheet</param>
        ///// <param name="col">The column number in the worksheet</param>
        ///// <returns></returns>		
        //internal ExcelCell Cell(int row, int col)
        //{
        //    return new ExcelCell(_values, row, col);
        //}
         /// <summary>
         /// Provides access to a range of cells
        /// </summary>  
        public ExcelRange Cells
        {
            get
            {
                CheckSheetType();
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
                CheckSheetType();
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
                CheckSheetType();
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
            //ExcelRow r;
            //ulong id = ExcelRow.GetRowID(_sheetID, row);
            //TODO: Fixa.
            //var v = _values.GetValue(row, 0);
            //if (v!=null)
            //{
            //    var ri=(RowInternal)v;
            //    r = new ExcelRow(this, row)
            //}
            //else
            //{
                //r = new ExcelRow(this, row);
                //_values.SetValue(row, 0, r);
                //_rows.Add(r);
            //}
            CheckSheetType();
            return new ExcelRow(this, row);
            //return r;
		}
		/// <summary>
		/// Provides access to an individual column within the worksheet so you can set its properties.
		/// </summary>
		/// <param name="col">The column number in the worksheet</param>
		/// <returns></returns>
		public ExcelColumn Column(int col)
		{
            CheckSheetType();
            ExcelColumn column = _values.GetValue(0, col) as ExcelColumn;
            // id=ExcelColumn.GetColumnID(_sheetID, col);
            if (column!=null)
            {
                //column = _columns[id] as ExcelColumn;
                if (column.ColumnMin != column.ColumnMax)
                {
                    int maxCol = column.ColumnMax;
                    column.ColumnMax=col;
                    ExcelColumn copy = CopyColumn(column, col + 1, maxCol);
                }
            }
            else
            {
                int r=0, c=col;
                if (_values.PrevCell(ref r, ref c))
                {
                    column = _values.GetValue(0, c) as ExcelColumn;
                    int maxCol = column.ColumnMax;
                    if (maxCol >= col)
                    {
                        column.ColumnMax = col-1;
                        if (maxCol > col)
                        {
                            ExcelColumn newC = CopyColumn(column, col + 1, maxCol);
                        }
                        return CopyColumn(column, col, col);
                    }
                }
                //foreach (ExcelColumn checkColumn in _columns)
                //{
                //    if (col > checkColumn.ColumnMin && col <= checkColumn.ColumnMax)
                //    {
                //        int maxCol = checkColumn.ColumnMax;
                //        checkColumn.ColumnMax = col - 1;
                //        if (maxCol > col)
                //        {
                //            ExcelColumn newC = CopyColumn(checkColumn, col + 1, maxCol);
                //        }
                //        return CopyColumn(checkColumn, col,col);                        
                //    }
                //}
                column = new ExcelColumn(this, col);
                _values.SetValue(0, col, column);
                //_columns.Add(column);
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
        internal ExcelColumn CopyColumn(ExcelColumn c, int col, int maxCol)
        {
            ExcelColumn newC = new ExcelColumn(this, col);
            newC.ColumnMax = maxCol;
            if (c.StyleName != "")
                newC.StyleName = c.StyleName;
            else
                newC.StyleID = c.StyleID;

            newC._hidden = c.Hidden;
            newC.OutlineLevel = c.OutlineLevel;
            newC.Phonetic = c.Phonetic;
            newC.BestFit = c.BestFit;
            //_columns.Add(newC);
            _values.SetValue(0, col, newC);
            newC.Width = c._width;
            return newC;
       }
        /// <summary>
        /// Make the current worksheet active.
        /// </summary>
        public void Select()
        {
            View.TabSelected = true;
            //Select(Address, true);
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
            CheckSheetType();
            int fromCol, fromRow, toCol, toRow;
            //Get rows and columns and validate as well
            ExcelCellBase.GetRowColFromAddress(Address, out fromRow, out fromCol, out toRow, out toCol);

            if (SelectSheet)
            {
                View.TabSelected = true;
            }
            View.SelectedRange = Address;
            View.ActiveCell = ExcelCellBase.GetAddress(fromRow, fromCol);            
        }
        /// <summary>
        /// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
        /// Make the current worksheet active.
        /// </summary>
        /// <param name="Address">An address range</param>
        public void Select(ExcelAddress Address)
        {
            CheckSheetType();
            Select(Address, true);
        }
        /// <summary>
        /// Selects a range in the worksheet. The active cell is the topmost cell of the first address.
        /// </summary>
        /// <param name="Address">A range of cells</param>
        /// <param name="SelectSheet">Make the sheet active</param>
        public void Select(ExcelAddress Address, bool SelectSheet)
        {

            CheckSheetType();
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
            View.ActiveCell = ExcelCellBase.GetAddress(Address.Start.Row, Address.Start.Column);
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
		public void  InsertRow(int rowFrom, int rows, int copyStylesFromRow)
		{
            CheckSheetType();
            var d = Dimension;
            //Check that cells aren't shifted outside the boundries
            if (d != null && d.End.Row > rowFrom && d.End.Row + rows > ExcelPackage.MaxRows)
            {
                throw (new ArgumentOutOfRangeException("Can't insert. Rows will be shifted outside the boundries of the worksheet."));
            }

            _values.Insert(rowFrom, 0, rows, 0);
            _formulas.Insert(rowFrom, 0, rows, 0);
            _styles.Insert(rowFrom, 0, rows, 0);
            _types.Insert(rowFrom, 0, rows, 0);
            _commentsStore.Insert(rowFrom, 0, rows, 0);
            _hyperLinks.Insert(rowFrom, 0, rows, 0);
            _flags.Insert(rowFrom, 0, rows, 0);

            foreach (var f in _sharedFormulas.Values)
            {
                if (f.StartRow >= rowFrom) f.StartRow += rows;
                var a = new ExcelAddressBase(f.Address);
                if (a._fromRow >= rowFrom)
                {
                    a._fromRow += rows;
                    a._toRow += rows;
                }
                else if (a._toRow >= rowFrom)
                {
                    a._toRow += rows;
                }
                f.Address = ExcelAddressBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
                f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, 0);
            }

            var cse = new CellsStoreEnumerator<object>(_formulas);
            while (cse.Next())
            {
                if (cse.Value is string)
                {
                    cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), rows, 0, rowFrom, 0);
                }
            }
            
            FixMergedCells(rowFrom, rows,false);
        }
        /// <summary>
        /// Adds a value to the row of merged cells to fix for inserts or deletes
        /// </summary>
        /// <param name="row"></param>
        /// <param name="rows"></param>
        /// <param name="delete"></param>
        private void FixMergedCells(int row, int rows, bool delete)
        {
            List<int> removeIndex = new List<int>();
            for (int i = 0; i < _mergedCells.Count; i++)
            {
                ExcelAddressBase addr = new ExcelAddressBase(_mergedCells[i]), newAddr;
                if (delete)
                {
                    newAddr = addr.DeleteRow(row, rows);
                    if (newAddr == null)
                    {
                        removeIndex.Add(i);
                        continue;
                    }
                }
                else
                {
                    newAddr = addr.AddRow(row, rows);
                }

                //The address has changed.
                if (newAddr._address != addr._address)
                {
                    //Set merged prop for cells
                    for (int r = newAddr._fromRow; r <= newAddr._toRow; r++)
                    {
                        for (int c = newAddr._fromCol; c <= newAddr._toCol; c++)
                        {
                            _flags.SetFlagValue(r, c, true, CellFlags.Merged);
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

                ExcelCellBase.GetRowColFromAddress(f.Address, out fromRow, out fromCol, out toRow, out toCol);
                if (position >= fromRow && position+(Math.Abs(rows)) <= toRow) //Insert/delete is whithin the share formula address
                {
                    if (rows > 0) //Insert
                    {
                        f.Address = ExcelCellBase.GetAddress(fromRow, fromCol) + ":" + ExcelCellBase.GetAddress(position - 1, toCol);
                        if (toRow != fromRow)
                        {
                            Formulas newF = new Formulas(SourceCodeTokenizer.Default);
                            newF.StartCol = f.StartCol;
                            newF.StartRow = position + rows;
                            newF.Address = ExcelCellBase.GetAddress(position + rows, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol);
                            newF.Formula = ExcelCellBase.TranslateFromR1C1(ExcelCellBase.TranslateToR1C1(f.Formula, f.StartRow, f.StartCol), position, f.StartCol);
                            added.Add(newF);
                        }
                    }
                    else
                    {
                        if (fromRow - rows < toRow)
                        {
                            f.Address = ExcelCellBase.GetAddress(fromRow, fromCol, toRow+rows, toCol);
                        }
                        else
                        {
                            f.Address = ExcelCellBase.GetAddress(fromRow, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol);
                        }
                    }
                }
                else if (position <= toRow)
                {
                    if (rows > 0) //Insert before shift down
                    {
                        f.StartRow += rows;
                        //f.Formula = ExcelCell.UpdateFormulaReferences(f.Formula, rows, 0, position, 0); //Recalc the cells positions
                        f.Address = ExcelCellBase.GetAddress(fromRow + rows, fromCol) + ":" + ExcelCellBase.GetAddress(toRow + rows, toCol);
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

                            f.Address = ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol);
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
                        newFormula = ExcelCellBase.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(formualR1C1, row, col), rows, 0, startRow, 0);
                        currentFormulaR1C1 = ExcelRangeBase.TranslateToR1C1(newFormula, row, col);
                    }
                    else
                    {
                        newFormula = ExcelCellBase.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(formualR1C1, row-rows, col), rows, 0, startRow, 0);
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
                            var refFormula = new Formulas(SourceCodeTokenizer.Default);
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
            CheckSheetType();

            _values.Delete(rowFrom, 1, rows, ExcelPackage.MaxColumns);
            _types.Delete(rowFrom, 1, rows, ExcelPackage.MaxColumns);
            _formulas.Delete(rowFrom, 1, rows, ExcelPackage.MaxColumns);
            _styles.Delete(rowFrom, 1, rows, ExcelPackage.MaxColumns);
            _flags.Delete(rowFrom, 1, rows, ExcelPackage.MaxColumns);
            _commentsStore.Delete(rowFrom, 1, rows, ExcelPackage.MaxColumns);
            _hyperLinks.Delete(rowFrom, 1, rows, ExcelPackage.MaxColumns);

            AdjustFormulasRow(rowFrom, rows);
            FixMergedCells(rowFrom, rows,true);
        }
        internal void AdjustFormulasRow(int rowFrom, int rows)
        {
            var delSF = new List<int>();
            foreach (var sf in _sharedFormulas.Values)
            {
                var a = new ExcelAddress(sf.Address).DeleteRow(rowFrom, rows);
                if (a==null)
                {
                    delSF.Add(sf.Index);
                }
                else
                {
                    sf.Address = a.Address;
                    sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, -rows, 0, rowFrom, 0);
                    if (sf.StartRow >= rowFrom)
                    {
                        sf.StartRow -= sf.StartRow;
                    }
                }
            }
            foreach (var ix in delSF)
            {
                _sharedFormulas.Remove(ix);
            }
            delSF = null;
            var cse = new CellsStoreEnumerator<object>(_formulas, rowFrom, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Value is string)
                {
                    cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), -rows, 0, rowFrom, 0);
                }
            }
        }
        /// <summary>
        /// Deletes the specified row from the worksheet.
        /// </summary>
        /// <param name="rowFrom">The number of the start row to be deleted</param>
        /// <param name="rows">Number of rows to delete</param>
        /// <param name="shiftOtherRowsUp">Not used. Rows are always shifted</param>
        public void DeleteRow(int rowFrom, int rows, bool shiftOtherRowsUp)
		{
            DeleteRow(rowFrom, rows);
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
            CheckSheetType();
            //ulong cellID = ExcelCellBase.GetCellID(SheetID, Row, Column);
            var v = _values.GetValue(Row, Column);
            if (v!=null)
            {
                //var cell = ((ExcelCell)_cells[cellID]);
                if (_flags.GetFlagValue(Row, Column, CellFlags.RichText))
                {
                    return (object)Cells[Row, Column].RichText.Text;
                }
                else
                {
                    return v;
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
            CheckSheetType();
            //ulong cellID=ExcelCellBase.GetCellID(SheetID, Row, Column);
            var v = _values.GetValue(Row, Column);           
            if (v==null)
            {
                return default(T);
            }

            //var cell=((ExcelCell)_cells[cellID]);
            if (_flags.GetFlagValue(Row, Column, CellFlags.RichText))
            {
                return (T)(object)Cells[Row, Column].RichText.Text;
            }
            else
            {
                return GetTypedValue<T>(v);
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
                        try
                        {
                            // Issue 14682 -- "GetValue<decimal>() won't convert strings"
                            // As suggested, after all special cases, all .NET to do it's 
                            // preferred conversion rather than simply returning the default
                            return (T)Convert.ChangeType(v, typeof(T));
                        }
                        catch (Exception)
                        {
                            // This was the previous behaviour -- no conversion is available.
                            return default(T);
                        }
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
            CheckSheetType();
            if (Row < 1 || Column < 1 || Row > ExcelPackage.MaxRows && Column > ExcelPackage.MaxColumns)
            {
                throw new ArgumentOutOfRangeException("Row or Column out of range");
            }            
            _values.SetValue(Row, Column, Value);
        }
        /// <summary>
        /// Set the value of a cell
        /// </summary>
        /// <param name="Address">The Excel address</param>
        /// <param name="Value">The value</param>
        public void SetValue(string Address, object Value)
        {
            CheckSheetType();
            int row, col;
            ExcelAddressBase.GetRowCol(Address, out row, out col, true);
            if (row < 1 || col < 1 || row > ExcelPackage.MaxRows && col > ExcelPackage.MaxColumns)
            {
                throw new ArgumentOutOfRangeException("Address is invalid or out of range");
            }
            _values.SetValue(row, col, Value);           
        }

        #region MergeCellId

        /// <summary>
        /// Get MergeCell Index No
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public int GetMergeCellId(int row, int column)
        {
            for (int i = 0; i < _mergedCells.Count; i++)
            {
                ExcelRange range = Cells[_mergedCells[i]];

                if (range.Start.Row <= row && row <= range.End.Row)
                {
                    if (range.Start.Column <= column && column <= range.End.Column)
                    {
                        return i + 1;
                    }
                }
            }
            return 0;
        }

        #endregion
		#endregion // END Worksheet Public Methods

		#region Worksheet Private Methods

		#region Worksheet Save
        internal void Save()
        {
                DeletePrinterSettings();

                if (_worksheetXml != null)
                {

                    if (!(this is ExcelChartsheet))
                    {
                        // save the header & footer (if defined)
                        if (_headerFooter != null)
                            HeaderFooter.Save();

                        var d = Dimension;
                        if (d == null)
                        {
                            this.DeleteAllNode("d:dimension/@ref");
                        }
                        else
                        {
                            this.SetXmlNodeString("d:dimension/@ref", d.Address);
                        }


                        if (Drawings.Count != null && _drawings.Count == 0)
                        {
                            //Remove node if no drawings exists.
                            DeleteNode("d:drawing");
                        }

                        SaveComments();
                        HeaderFooter.SaveHeaderFooterImages();
                        SaveTables();
                        SavePivotTables();
                    }
                }

                if (Drawings.UriDrawing!=null)
                {
                    if (Drawings.Count == 0)
                    {                                            
                        Part.DeleteRelationship(Drawings._drawingRelation.Id);
                        _package.Package.DeletePart(Drawings.UriDrawing);                    
                    }
                    else
                    {
                        Packaging.ZipPackagePart partPack = Drawings.Part;
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
        }
        internal void SaveHandler(ZipOutputStream stream, Ionic.Zlib.CompressionLevel compressionLevel, string fileName)
        {
                    //Init Zip
                    stream.CodecBufferSize = 8096;
                    stream.CompressionLevel = compressionLevel;
                    stream.PutNextEntry(fileName);

                    
                    SaveXml(stream);
        }

        

        ///// <summary>
        ///// Saves the worksheet to the package.
        ///// </summary>
        //internal void Save()  // Worksheet Save
        //{
        //    DeletePrinterSettings();

        //    if (_worksheetXml != null)
        //    {
                
        //        // save the header & footer (if defined)
        //        if (_headerFooter != null)
        //            HeaderFooter.Save();

        //        var d = Dimension;
        //        if (d == null)
        //        {
        //            this.DeleteAllNode("d:dimension/@ref");
        //        }
        //        else
        //        {
        //            this.SetXmlNodeString("d:dimension/@ref", d.Address);
        //        }
                

        //        if (_drawings != null && _drawings.Count == 0)
        //        {
        //            //Remove node if no drawings exists.
        //            DeleteNode("d:drawing");
        //        }

        //        SaveComments();
        //        HeaderFooter.SaveHeaderFooterImages();
        //        SaveTables();
        //        SavePivotTables();
        //        SaveXml();
        //    }
            
        //    if (Drawings.UriDrawing!=null)
        //    {
        //        if (Drawings.Count == 0)
        //        {                    
        //            Part.DeleteRelationship(Drawings._drawingRelation.Id);
        //            _package.Package.DeletePart(Drawings.UriDrawing);                    
        //        }
        //        else
        //        {
        //            Packaging.ZipPackagePart partPack = Drawings.Part;
        //            Drawings.DrawingXml.Save(partPack.GetStream(FileMode.Create, FileAccess.Write));
        //            foreach (ExcelDrawing d in Drawings)
        //            {
        //                if (d is ExcelChart)
        //                {
        //                    ExcelChart c = (ExcelChart)d;
        //                    c.ChartXml.Save(c.Part.GetStream(FileMode.Create, FileAccess.Write));
        //                }
        //            }
        //        }
        //    }
        //}

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
                    var rel = Part.GetRelationship(relID);
                    Uri printerSettingsUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
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
                    RemoveLegacyDrawingRel(VmlDrawingsComments.RelId);
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
                        var rel = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, _comments.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships+"/comments");
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
                        var rel = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, _vmlDrawings.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
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
                    var colVal = new HashSet<string>();
                    foreach (var col in tbl.Columns)
                    {                        
                        var n=col.Name.ToLower();
                        if(colVal.Contains(n))
                        {
                            throw(new InvalidDataException(string.Format("Table {0} Column {1} does not have a unique name.", tbl.Name, col.Name)));
                        }
                        colVal.Add(n);
                        if (tbl.ShowHeader)
                        {
                            _values.SetValue(tbl.Address._fromRow, colNum, col.Name);
                        }
                        if (tbl.ShowTotal)
                        {
                            SetTableTotalFunction(tbl, col, colNum);
                        }
                        if (!string.IsNullOrEmpty(col.CalculatedColumnFormula))
                        {
                            int fromRow = tbl.ShowHeader ? tbl.Address._fromRow + 1 : tbl.Address._fromRow;
                            int toRow = tbl.ShowTotal ? tbl.Address._toRow - 1 : tbl.Address._toRow;
                            for (int row = fromRow; row <= toRow; row++)
                            {
                                //Cell(row, colNum).Formula = col.CalculatedColumnFormula;
                                SetFormula(row, colNum, col.CalculatedColumnFormula);
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
                    var rel = Part.CreateRelationship(UriHelper.GetRelativeUri(WorksheetUri, tbl.TableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/table");
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

        internal void SetTableTotalFunction(ExcelTable tbl, ExcelTableColumn col, int colNum=-1)
        {
            if (colNum == -1)
            {
                for (int i = 0; i < tbl.Columns.Count; i++)
                {
                    if (tbl.Columns[i].Name == col.Name)
                    {
                        colNum = tbl.Address._fromCol + i;
                    }
                }
            }
            if (col.TotalsRowFunction == RowFunctions.Custom)
            {
                SetFormula(tbl.Address._toRow, colNum, col.TotalsRowFormula);
            }
            else if (col.TotalsRowFunction != RowFunctions.None)
            {
                switch (col.TotalsRowFunction)
                {
                    case RowFunctions.Average:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "101"));
                        break;
                    case RowFunctions.Count:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "102"));
                        break;
                    case RowFunctions.CountNums:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "103"));
                        break;
                    case RowFunctions.Max:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "104"));
                        break;
                    case RowFunctions.Min:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "105"));
                        break;
                    case RowFunctions.StdDev:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "107"));
                        break;
                    case RowFunctions.Var:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "110"));
                        break;
                    case RowFunctions.Sum:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "109"));
                        break;
                    default:
                        throw (new Exception("Unknown RowFunction enum"));
                }
            }
            else
            {
                _values.SetValue(tbl.Address._toRow, colNum, col.TotalsRowLabel);

            }
        }

        internal void SetFormula(int row, int col, object value)
        {
            _formulas.SetValue(row, col, value);
            if (!_values.Exists(row, col)) _values.SetValue(row, col, null);
        }
        internal void SetStyle(int row, int col, int value)
        {
            _styles.SetValue(row, col, value);
            if(!_values.Exists(row,col)) _values.SetValue(row, col, null);
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
            return string.Format("SUBTOTAL({0},{1}[{2}])", FunctionNum, col._tbl.Name, col.Name);
        }
        private void SaveXml(Stream stream)
        {
            //Create the nodes if they do not exist.
            StreamWriter sw = new StreamWriter(stream, System.Text.Encoding.Default, 65536);
            if (this is ExcelChartsheet)
            {
                sw.Write(_worksheetXml.OuterXml);
            }
            else
            {
                CreateNode("d:cols");
                CreateNode("d:sheetData");
                CreateNode("d:mergeCells");
                CreateNode("d:hyperlinks");
                CreateNode("d:rowBreaks");
                CreateNode("d:colBreaks");

                //StreamWriter sw=new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));
                var xml = _worksheetXml.OuterXml;
                int colStart = 0, colEnd = 0;
                GetBlockPos(xml, "cols", ref colStart, ref colEnd);

                sw.Write(xml.Substring(0, colStart));
                var colBreaks = new List<int>();
                //if (_columns.Count > 0)
                //{
                UpdateColumnData(sw);
                //}

                int cellStart = colEnd, cellEnd = colEnd;
                GetBlockPos(xml, "sheetData", ref cellStart, ref cellEnd);

                sw.Write(xml.Substring(colEnd, cellStart - colEnd));
                var rowBreaks = new List<int>();
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
                //if (_hyperLinkCells.Count > 0)
                //{
                UpdateHyperLinks(sw);
                // }

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
            }
            sw.Flush();
            //sw.Close();
        }
        private void UpdateColBreaks(StreamWriter sw)
        {
            StringBuilder breaks = new StringBuilder();
            int count = 0;
            var cse = new CellsStoreEnumerator<object>(_values, 0, 0, 0, ExcelPackage.MaxColumns);
            //foreach (ExcelColumn col in _columns)
            while(cse.Next())
            {
                var col=cse.Value as ExcelColumn;
                if (col != null && col.PageBreak)
                {
                    breaks.AppendFormat("<brk id=\"{0}\" max=\"16383\" man=\"1\" />", cse.Column);
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
            var cse = new CellsStoreEnumerator<object>(_values, 0, 0, ExcelPackage.MaxRows, 0);
            //foreach(ExcelRow row in _rows)            
            while(cse.Next())
            {
                var row=cse.Value as ExcelRow;
                if (row != null && row.PageBreak)
                {
                    breaks.AppendFormat("<brk id=\"{0}\" max=\"1048575\" man=\"1\" />", cse.Row);
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
            var cse = new CellsStoreEnumerator<object>(_values, 0, 1, 0, ExcelPackage.MaxColumns);
            //sw.Write("<cols>");
            //foreach (ExcelColumn col in _columns)
            bool first = true;
            while(cse.Next())
            {
                if (first)
                {
                    sw.Write("<cols>");
                    first = false;
                }
                var col = cse.Value as ExcelColumn;
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

                var styleID = col.StyleID >= 0 ? cellXfs[col.StyleID].newID : col.StyleID;
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
            if (!first)
            {
                sw.Write("</cols>");
            }
        }
        /// <summary>
        /// Insert row and cells into the XML document
        /// </summary>
        private void UpdateRowCellData(StreamWriter sw)
        {
            ExcelStyleCollection<ExcelXfs> cellXfs = _package.Workbook.Styles.CellXfs;
            
            //_hyperLinkCells = new List<ulong>();
            int row = -1;

            //foreach (ExcelRow r in _rows)
            //{
            //    int nextCell = ~_cells.IndexOf(r.RowID);
            //    if (nextCell >= 0 && (nextCell >= _cells.Count || ((ExcelCell)_cells[nextCell]).Row != r.Row))
            //    {
            //        _cells.Add(r);
            //    }
            //}

            StringBuilder sbXml = new StringBuilder();
            var ss = _package.Workbook._sharedStrings;
            var styles = _package.Workbook.Styles;
            sw.Write("<sheetData>");
            var cse = new CellsStoreEnumerator<object>(_values, 1, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            //foreach (IRangeID r in _cells)
            while(cse.Next())
            {
                if (cse.Column>0)
                {
                    //ExcelCell cell = (ExcelCell)r;
                    //long styleID = cell.StyleID >= 0 ? cellXfs[cell.StyleID].newID : cell.StyleID;
                    int styleID = cellXfs[styles.GetStyleId(this, cse.Row, cse.Column)].newID;
                    //styleID = styleID >= 0 ? cellXfs[styleID].newID : styleID;
                    //Add the row element if it's a new row
                    if (cse.Row != row)
                    {
                        WriteRow(sw, cellXfs, row, cse.Row);
                        row = cse.Row;
                    }
                    object v = cse.Value;
                    object formula = _formulas.GetValue(cse.Row, cse.Column);
                    if (formula is int)
                    {
                        int sfId = (int)formula;
                        var f = _sharedFormulas[(int)sfId];
                        if (f.Address.IndexOf(':') > 0)
                        {
                            if (f.StartCol == cse.Column && f.StartRow == cse.Row)
                            {
                                if (f.IsArray)
                                {
                                    sw.Write("<c r=\"{0}\" s=\"{1}\"><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, f.Address, SecurityElement.Escape(f.Formula), GetFormulaValue(v));
                                }
                                else
                                {
                                    sw.Write("<c r=\"{0}\" s=\"{1}\"><f ref=\"{2}\" t=\"shared\"  si=\"{3}\">{4}</f>{5}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, f.Address, sfId, SecurityElement.Escape(f.Formula), GetFormulaValue(v));
                                }

                            }
                            else if (f.IsArray)
                            {
                                sw.Write("<c r=\"{0}\" s=\"{1}\" />", cse.CellAddress, styleID < 0 ? 0 : styleID);
                            }
                            else
                            {
                                sw.Write("<c r=\"{0}\" s=\"{1}\"><f t=\"shared\" si=\"{2}\" />{3}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, sfId, GetFormulaValue(v));
                            }
                        }
                        else
                        {
                            // We can also have a single cell array formula
                            if(f.IsArray)
                            {
                                sw.Write("<c r=\"{0}\" s=\"{1}\"><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, string.Format("{0}:{1}", f.Address, f.Address), SecurityElement.Escape(f.Formula), GetFormulaValue(v));
                            }
                            else
                            {
                                sw.Write("<c r=\"{0}\" s=\"{1}\">", f.Address, styleID < 0 ? 0 : styleID);
                                sw.Write("<f>{0}</f>{1}</c>", SecurityElement.Escape(f.Formula), GetFormulaValue(v));
                            }
                        }
                    }
                    else if (formula!=null && formula.ToString()!="")
                    {
                        sw.Write("<c r=\"{0}\" s=\"{1}\" {2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(v));
                        sw.Write("<f>{0}</f>{1}</c>", SecurityElement.Escape(formula.ToString()), GetFormulaValue(v));
                    }
                    else
                    {
                        if (v == null)
                        {
                            sw.Write("<c r=\"{0}\" s=\"{1}\" />", cse.CellAddress, styleID < 0 ? 0 : styleID);
                        }
                        else
                        {
                            if ((v.GetType().IsPrimitive || v is double || v is decimal || v is DateTime || v is TimeSpan) && _types.GetValue(cse.Row,cse.Column) != "s")
                            {
                                string sv = GetValueForXml(v);
                                sw.Write("<c r=\"{0}\" s=\"{1}\" {2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(v));
                                sw.Write("<v>{0}</v></c>", sv);
                            }
                            else
                            {
                                int ix;
                                if (!ss.ContainsKey(v.ToString()))
                                {
                                    ix = ss.Count;
                                    ss.Add(v.ToString(), new ExcelWorkbook.SharedStringItem() { isRichText = _flags.GetFlagValue(cse.Row,cse.Column,CellFlags.RichText), pos = ix });
                                }
                                else
                                {
                                    ix = ss[v.ToString()].pos;
                                }
                                sw.Write("<c r=\"{0}\" s=\"{1}\" t=\"s\">", cse.CellAddress, styleID < 0 ? 0 : styleID);
                                sw.Write("<v>{0}</v></c>", ix);
                            }
                        }
                    }
                    ////Update hyperlinks.
                   //if (cell.Hyperlink != null)
                    //{
                    //    _hyperLinkCells.Add(cell.CellID);
                    //}
                }
                else  //ExcelRow
                {
                    //int newRow=((ExcelRow)cse.Value).Row;
                    WriteRow(sw, cellXfs, row, cse.Row);
                    row = cse.Row;
                }
            }

            if (row != -1) sw.Write("</row>");
            sw.Write("</sheetData>");

            
        }

        private object GetFormulaValue(object v)
        {
            if (_package.Workbook._isCalculated)
            {
                return "<v>" + GetValueForXml(v) + "</v>";
            }
            else
            {
                return "";
            }
        }

        private string GetCellType(object v)
        {
            if (v is bool)
            {
                return " t=\"b\"";
            }
            else if ((v is double && double.IsInfinity((double)v)) || v is ExcelErrorValue)
            {
                return " t=\"e\"";
            }
            else
            {
                return "";
            }
        }

        private static string GetValueForXml(object v)
        {
            string s;
            try
            {
                if (v is DateTime)
                {
                    s = ((DateTime)v).ToOADate().ToString(CultureInfo.InvariantCulture);
                }
                else if (v is TimeSpan)
                {
                    s = new DateTime(((TimeSpan)v).Ticks).ToOADate().ToString(CultureInfo.InvariantCulture); ;
                }
                else if(v.GetType().IsPrimitive || v is double || v is decimal)
                {
                    if (v is double && double.IsNaN((double)v))
                    {
                        s = "0";
                    }
                    else if (v is double && double.IsInfinity((double)v))
                    {
                        s = "#NUM!";
                    }
                    else
                    {
                        s = Convert.ToDouble(v, CultureInfo.InvariantCulture).ToString("R15", CultureInfo.InvariantCulture);
                    }
                }
                else
                {
                    s = v.ToString();
                }
            }

            catch
            {
                s = "0";
            }
            return s;
        }
        private void WriteRow(StreamWriter sw, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
        {
            if (prevRow != -1) sw.Write("</row>");
            //ulong rowID = ExcelRow.GetRowID(SheetID, row);
            sw.Write("<row r=\"{0}\" ", row);
            ExcelRow currRow = _values.GetValue(row, 0) as ExcelRow;
            if (currRow!=null)
            {
                
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

                if (currRow.OutlineLevel > 0)
                {
                    sw.Write("outlineLevel =\"{0}\" ", currRow.OutlineLevel);
                    if (currRow.Collapsed)
                    {
                        if (currRow.Hidden)
                        {
                            sw.Write(" collapsed=\"1\" ");
                        }
                        else
                        {
                            sw.Write(" collapsed=\"1\" hidden=\"1\" "); //Always hidden
                        }
                    }
                }
                if (currRow.Phonetic)
                {
                    sw.Write("ph=\"1\" ");
                }
            }
            var s = _styles.GetValue(row, 0);
            if (s > 0)
            {
                sw.Write("s=\"{0}\" customFormat=\"1\"", cellXfs[s].newID);
            }
            sw.Write(">");
        }

        /// <summary>
        /// Update xml with hyperlinks 
        /// </summary>
        /// <param name="sw">The stream</param>
        private void UpdateHyperLinks(StreamWriter sw)
        {
                Dictionary<string, string> hyps = new Dictionary<string, string>();
                var cse = new CellsStoreEnumerator<Uri>(_hyperLinks);
                bool first = true;
                //foreach (ulong cell in _hyperLinks)
                while(cse.Next())
                {
                    if (first)
                    {
                        sw.Write("<hyperlinks>");
                        first = false;
                    }
                    //int row, col;
                    var uri = _hyperLinks.GetValue(cse.Row, cse.Column);
                    //ExcelCell cell = _cells[cellId] as ExcelCell;
                    if (uri is ExcelHyperLink && !string.IsNullOrEmpty((uri as ExcelHyperLink).ReferenceAddress))
                    {
                        ExcelHyperLink hl = uri as ExcelHyperLink;
                        sw.Write("<hyperlink ref=\"{0}\" location=\"{1}\" {2}{3}/>",
                                Cells[cse.Row, cse.Column, cse.Row + hl.RowSpann, cse.Column + hl.ColSpann].Address, 
                                ExcelCellBase.GetFullAddress(Name, hl.ReferenceAddress),
                                    string.IsNullOrEmpty(hl.Display) ? "" : "display=\"" + SecurityElement.Escape(hl.Display) + "\" ",
                                    string.IsNullOrEmpty(hl.ToolTip) ? "" : "tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\" ");
                    }
                    else if( uri!=null)
                    {
                        string id;
                        Uri hyp;
                        if (uri is ExcelHyperLink)
                        {
                            hyp = ((ExcelHyperLink)uri).OriginalUri;
                        }
                        else
                        {
                            hyp = uri;
                        }
                        if (hyps.ContainsKey(hyp.OriginalString))
                        {
                            id = hyps[hyp.OriginalString];
                        }
                        else
                        {
                            var relationship = Part.CreateRelationship(hyp, Packaging.TargetMode.External, ExcelPackage.schemaHyperlink);
                            if (uri is ExcelHyperLink)
                            {
                                ExcelHyperLink hl = uri as ExcelHyperLink;
                                sw.Write("<hyperlink ref=\"{0}\" {2}{3}r:id=\"{1}\" />", ExcelCellBase.GetAddress(cse.Row, cse.Column), relationship.Id,                                
                                    string.IsNullOrEmpty(hl.Display) ? "" : "display=\"" + SecurityElement.Escape(hl.Display) + "\" ",
                                    string.IsNullOrEmpty(hl.ToolTip) ? "" : "tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\" ");
                            }
                            else
                            {
                                sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\" />", ExcelCellBase.GetAddress(cse.Row, cse.Column), relationship.Id);
                            }
                            id = relationship.Id;
                        }
                        //cell.HyperLinkRId = id;
                    }
                }
                if (!first)
                {
                    sw.Write("</hyperlinks>");
                }
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
                CheckSheetType();
                int fromRow, fromCol, toRow, toCol;
                if (_values.GetDimension(out fromRow, out fromCol, out toRow, out toCol))
                {
                    ExcelAddressBase addr = new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
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

        private ExcelProtectedRangeCollection _protectedRanges;
        public ExcelProtectedRangeCollection ProtectedRanges
        {
            get
            {
                if (_protectedRanges == null)
                    _protectedRanges = new ExcelProtectedRangeCollection(NameSpaceManager, TopNode, this);
                return _protectedRanges;
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
                CheckSheetType();
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
                CheckSheetType();
                if (_pivotTables == null)
                {
                    _pivotTables = new ExcelPivotTableCollection(this);
                }
                return _pivotTables;
            }
        }
        private ExcelConditionalFormattingCollection _conditionalFormatting = null;
        /// <summary>
        /// ConditionalFormatting defined in the worksheet. Use the Add methods to create ConditionalFormatting and add them to the worksheet. Then
        /// set the properties on the instance returned.
        /// </summary>
        /// <seealso cref="ExcelConditionalFormattingCollection"/>
        public ExcelConditionalFormattingCollection ConditionalFormatting
        {
            get
            {
                CheckSheetType();
                if (_conditionalFormatting == null)
                {
                    _conditionalFormatting = new ExcelConditionalFormattingCollection(this);
                }
                return _conditionalFormatting;
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
                CheckSheetType();
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
        internal void RemoveLegacyDrawingRel(string relID)
        {
            var n = WorksheetXml.DocumentElement.SelectSingleNode(string.Format("d:legacyDrawing[@r:id=\"{0}\"]", relID), NameSpaceManager);
            if (n != null)
            {
                n.ParentNode.RemoveChild(n);
            }
        }
        internal string GetFormula(int row, int col)
        {
            var v = _formulas.GetValue(row, col);
            if (v is int)
            {
                return _sharedFormulas[(int)v].GetFormula(row,col);
            }
            else if (v != null)
            {
                return v.ToString();
            }
            else
            {
                return "";
            }
        }
        internal string GetFormulaR1C1(int row, int col)
        {
            var v = _formulas.GetValue(row, col);
            if (v is int)
            {
                var sf = _sharedFormulas[(int)v];
                return ExcelCellBase.TranslateToR1C1(sf.Formula, sf.StartRow, sf.StartCol);
            }
            else if (v != null)
            {
                return ExcelCellBase.TranslateToR1C1(v.ToString(), row, col);
            }
            else
            {
                return "";
            }
        }



        public void Dispose()
        {
            _values.Dispose();
            _formulas.Dispose();
            _flags.Dispose();
            _hyperLinks.Dispose();
            _styles.Dispose();
            _types.Dispose();
            _commentsStore.Dispose();

            if (_formulaTokens != null) _commentsStore.Dispose();
            _values = null;
            _formulas = null;
            _flags = null;
            _hyperLinks = null;
            _styles = null;
            _types = null;
            _commentsStore = null;
            _formulaTokens = null;

            _package = null;
            _pivotTables = null;
            _protection = null;
            _sharedFormulas.Clear();
            _sharedFormulas = null;
            _sheetView = null;
            _tables = null;
            _vmlDrawings = null;
            _conditionalFormatting = null;
            _dataValidation = null;
            _drawings = null;
        }
    }  // END class Worksheet
}
