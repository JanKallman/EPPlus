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
 * Jan Källman                      Total rewrite               2010-03-01
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
namespace OfficeOpenXml
{
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
	public sealed class ExcelWorksheet : XmlHelper
	{
        internal class Formulas
        {
            internal int Index { get; set; }
            internal string Address { get; set; }
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
        internal RangeCollection _formulaCells;
        internal static CultureInfo _ci=new CultureInfo("en-US");
        internal int _minCol = ExcelPackage.MaxColumns;
        internal int _maxCol = 0;
        internal List<ulong> _hyperLinkCells;   //Used when saving the sheet
		/// <summary>
		/// Reference to the parent package
		/// For internal use only!
		/// </summary>
        #region Worksheet Private Properties
        internal ExcelPackage xlPackage;
		private Uri _worksheetUri;
		private string _name;
		private int _sheetID;
        private int _positionID;
        private eWorkSheetHidden _hidden;
		private string _relationshipID;
		private XmlDocument _worksheetXml;
        internal ExcelWorksheetView _sheetView;
		private ExcelHeaderFooter _headerFooter;
        #endregion
        #region ExcelWorksheet Constructor
        /// <summary>
        /// A worksheet
        /// </summary>
        /// <param name="ns">Namespacemanager</param>
        /// <param name="excelPackage">Package</param>
        /// <param name="uriWorksheet">URI</param>
        /// <param name="sheetName">Name of the sheet</param>
        /// <param name="SheetID">Sheet id</param>
        /// <param name="PositionID">Position</param>
        /// <param name="Hide">hide</param>
        public ExcelWorksheet(XmlNamespaceManager ns, ExcelPackage excelPackage, string relID, 
                              Uri uriWorksheet, string sheetName, int sheetID, int positionID,
                              eWorkSheetHidden hide) :
            base(ns, null)
        {
            SchemaNodeOrder = new string[] { "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "protectedRanges", "autoFilter", "customSheetViews", "mergeCells", "conditionalFormatting", "hyperlinks", "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks", "drawing", "legacyDrawingHF"};
            xlPackage = excelPackage;   
            _relationshipID = relID;
            _worksheetUri = uriWorksheet;
            _name = sheetName;
            _sheetID = sheetID;
            _positionID = positionID;
            Hidden = hide;
            _names = new ExcelNamedRangeCollection(this);
            CreateXml();
            TopNode = _worksheetXml.DocumentElement;
        }
        #endregion

		#region Worksheet Public Properties
		/// <summary>
		/// Read-only: the Uri to the worksheet within the package
		/// </summary>
		protected internal Uri WorksheetUri { get { return (_worksheetUri); } }
		/// <summary>
		/// Read-only: a reference to the PackagePart for the worksheet within the package
		/// </summary>
		protected internal PackagePart Part { get { return (xlPackage.Package.GetPart(WorksheetUri)); } }
		/// <summary>
		/// Read-only: the ID for the worksheet's relationship with the workbook in the package
		/// </summary>
		protected internal string RelationshipID { get { return (_relationshipID); } }
		/// <summary>
		/// The unique identifier for the worksheet.  Note that these can be random, so not
		/// too useful in code!
		/// </summary>
		protected internal int SheetID { get { return (_sheetID); } }
        /// <summary>
        /// The position of the worksheet.
        /// </summary>
        protected internal int PositionID { get { return (_positionID); } }
        /// <summary>
        /// Address for autofilter
        /// <seealso cref="ExcelRangeBase.AutoFilter" />        
        /// </summary>
        public ExcelAddressBase AutoFilterAddress
        {
            get
            {
                return new ExcelAddressBase(GetXmlNode("d:autoFilter/@ref"));
            }
            internal set
            {
                SetXmlNode("d:autoFilter/@ref", value.Address);
            }
        }

		/// <summary>
		/// Returns a ExcelWorksheetView object that allows you to
		/// set the view state properties of the worksheet
		/// </summary>
		public ExcelWorksheetView View
		{
			get
			{
				if (_sheetView == null)
				{
                    XmlNode node = _worksheetXml.SelectSingleNode("//d:sheetView", NameSpaceManager);
                    if (node == null)
                    {
                        CreateNode("d:sheetViews/d:sheetView"); //this one shouls always exist. but check anyway
                    }
                    _sheetView = new ExcelWorksheetView(NameSpaceManager, node, this);
				}
				return (_sheetView);
			}
		}

		#region Name // Worksheet Name
		/// <summary>
		/// The worksheet's name as it appears on the tab
		/// </summary>
		public string Name
		{
			get { return (_name); }
			set
			{
				XmlNode sheetNode = xlPackage.Workbook.WorkbookXml.SelectSingleNode(string.Format("//d:sheet[@sheetId={0}]", _sheetID), NameSpaceManager);
				if (sheetNode != null)
				{
					XmlAttribute nameAttr = (XmlAttribute)sheetNode.Attributes.GetNamedItem("name");
					if (nameAttr != null)
					{
						nameAttr.Value = value;
					}
				}
				_name = value;
			}
		}
		#endregion // END Worksheet Name
        private ExcelNamedRangeCollection _names;
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
		#region Hidden
		/// <summary>
		/// Indicates if the worksheet is hidden in the workbook
		/// </summary>
		public eWorkSheetHidden Hidden
		{
			get { return (_hidden); }
			set
			{
				XmlElement sheetNode = xlPackage.Workbook.WorkbookXml.SelectSingleNode(string.Format("//d:sheet[@sheetId={0}]", _sheetID), NameSpaceManager) as XmlElement;
				if (sheetNode != null)
				{
                    if (value==eWorkSheetHidden.Hidden)
                    {
                        sheetNode.SetAttribute("state", "hidden");
                    }
                    else if (value == eWorkSheetHidden.VeryHidden)
                    {
                        sheetNode.SetAttribute("state", "veryHidden");
                    }
                    else
                    {
                        sheetNode.RemoveAttribute("state");
                    }
				}
				_hidden = value;
			}
		}
		#endregion

		#region defaultRowHeight
        double _defaultRowHeight = double.NaN;
        /// <summary>
		/// Allows you to get/set the default height of all rows in the worksheet
		/// </summary>
        public double defaultRowHeight
		{
			get 
			{
				if(double.IsNaN(_defaultRowHeight))
                {
                    XmlElement sheetFormat = (XmlElement)WorksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
                    if (sheetFormat == null)
                    {
                        _defaultRowHeight = 15; // Excel's default height
                    }
                    else
                    {
                        string ret = sheetFormat.GetAttribute("defaultRowHeight");
                        _defaultRowHeight = double.Parse(ret, _ci);
                    }
                }
				return _defaultRowHeight;
			}
			set
			{
                _defaultRowHeight = value;
                XmlElement sheetFormat = (XmlElement)WorksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
				if (sheetFormat == null)
				{
					// create the node as it does not exist
					sheetFormat = WorksheetXml.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
					// find location to insert new element
					XmlNode sheetViews = WorksheetXml.SelectSingleNode("//d:sheetViews", NameSpaceManager);
					// insert the new node
					WorksheetXml.DocumentElement.InsertAfter(sheetFormat, sheetViews);
				}
				sheetFormat.SetAttribute("defaultRowHeight", value.ToString());
			}
		}
		#endregion
        #region defaultColWidth
        /// <summary>
        /// Allows you to get/set the default width of all rows in the worksheet
        /// </summary>
        public double defaultColWidth
        {
            get
            {
                double retValue = 9.140625 ; // Excel's default height
                XmlElement sheetFormat = (XmlElement)WorksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
                if (sheetFormat != null)
                {
                    string ret = sheetFormat.GetAttribute("defaultColWidth");
                    if (ret != "")
                        retValue = double.Parse(ret, _ci);
                }
                return retValue;
            }
            set
            {
                XmlElement sheetFormat = (XmlElement)WorksheetXml.SelectSingleNode("//d:sheetFormatPr", NameSpaceManager);
                if (sheetFormat == null)
                {
                    // create the node as it does not exist
                    sheetFormat = WorksheetXml.CreateElement("sheetFormatPr", ExcelPackage.schemaMain);
                    // find location to insert new element
                    XmlNode sheetViews = WorksheetXml.SelectSingleNode("//d:sheetViews", NameSpaceManager);
                    // insert the new node
                    WorksheetXml.DocumentElement.InsertAfter(sheetFormat, sheetViews);
                }
                sheetFormat.SetAttribute("defaultColWidth", value.ToString());
            }
        }
        #endregion
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
                SetXmlNode(outLineSummaryBelowPath, value ? "1" : "0");
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
                SetXmlNode(outLineSummaryRightPath, value ? "1" : "0");
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
                SetXmlNode(outLineApplyStylePath, value ? "1" : "0");
            }
        }
        #region WorksheetXml
		/// <summary>
		/// The XML document holding all the worksheet data.
		/// </summary>
		public XmlDocument WorksheetXml
		{
			get
			{
				return (_worksheetXml);
			}
		}
        private ExcelVmlDrawings _vmlDrawings = null;
        internal ExcelVmlDrawings VmlDrawings
        {
            get
            {
                if (_vmlDrawings == null)
                {
                    var vmlNode = _worksheetXml.DocumentElement.SelectSingleNode("d:legacyDrawing/@r:id", NameSpaceManager);
                    if (vmlNode == null)
                    {
                        _vmlDrawings = new ExcelVmlDrawings(xlPackage, this, null);

                    }
                    else
                    {
                        var rel = Part.GetRelationship(vmlNode.Value);
                        var vmlUri = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

                        _vmlDrawings = new ExcelVmlDrawings(xlPackage, this, vmlUri);
                        _vmlDrawings.RelId = rel.Id;
                    }
                }
                return _vmlDrawings;
            }
        }
        internal ExcelCommentCollection _comments = null;
        public ExcelCommentCollection Comments
        {
            get
            {
                if (_comments == null)
                {
                    _comments = new ExcelCommentCollection(xlPackage, this, NameSpaceManager);
                }
                return _comments;
            }
        }
        private void CreateXml()
        {
            _worksheetXml = new XmlDocument();
            _worksheetXml.PreserveWhitespace = ExcelPackage.preserveWhitespace;
            PackagePart packPart = xlPackage.Package.GetPart(WorksheetUri);
            string xml = "";

            // First Columns, rows, cells, mergecells, hyperlinks and pagebreakes are loaded from a xmlstream to optimize speed...

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

            ClearNodes();
        }

        private void LoadRowPageBreakes(XmlTextReader xr)
        {
            if(!ReadUntil(xr, "rowBreaks","colBreaks")) return;
            while (xr.Read())
            {
                if (xr.Name == "brk")
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
                if (xr.Name == "brk")
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
                            stream.Seek(end - BLOCKSIZE, SeekOrigin.Begin);
                            int size = stream.Length - stream.Position < BLOCKSIZE ? (int)(stream.Length - stream.Position) : BLOCKSIZE;
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
            while (!Array.Exists(tagName, tag => xr.Name.EndsWith(tag)))
            {
                xr.Read();
                if (xr.EOF) return false;
            }
            return (xr.Name.EndsWith(tagName[0]));
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
                    if(xr.Name!="col") break;
                    int min = int.Parse(xr.GetAttribute("min"));

                    int style;
                    if (xr.GetAttribute("style") == null || !int.TryParse(xr.GetAttribute("style"), out style))
                    {
                        style = 0;
                    }
                    ExcelColumn col = new ExcelColumn(this, min);
                   
                    col._columnMax = int.Parse(xr.GetAttribute("max")); 
                    col.StyleID = style;
                    col.Width = xr.GetAttribute("width") == null ? 0 : double.Parse(xr.GetAttribute("width"), _ci); 
                    col.BestFit = xr.GetAttribute("bestFit") != null && xr.GetAttribute("bestFit") == "1" ? true : false;
                    col.Collapsed = xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed") == "1" ? true : false;
                    col.Phonetic = xr.GetAttribute("phonetic") != null && xr.GetAttribute("phonetic") == "1" ? true : false;
                    col.OutlineLevel = xr.GetAttribute("outlineLevel") == null ? 0 : int.Parse(xr.GetAttribute("outlineLevel"), _ci);
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
                if (xr.Name == nodeText || xr.Name == altNode) return true;
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
                if (xr.Name == "hyperlink")
                {
                    int fromRow, fromCol, toRow, toCol;
                    ExcelCell.GetRowColFromAddress(xr.GetAttribute("ref"), out fromRow, out fromCol, out toRow, out toCol);
                    ulong id = ExcelCell.GetCellID(_sheetID, fromRow, fromCol);
                    ExcelCell cell = _cells[id] as ExcelCell;
                    if (xr.GetAttribute("id", ExcelPackage.schemaRelationships) != null)
                    {
                        cell.HyperLinkRId = xr.GetAttribute("id", ExcelPackage.schemaRelationships);
                        cell.Hyperlink = Part.GetRelationship(cell.HyperLinkRId).TargetUri;
                        Part.DeleteRelationship(cell.HyperLinkRId); //Delete the relationship, it is recreated when we save the package.
                    }
                    else if (xr.GetAttribute("location") != null)
                    {
                        ExcelHyperLink hl = new ExcelHyperLink(xr.GetAttribute("location"), xr.GetAttribute("display"));
                        hl.RowSpann = toRow - fromRow;
                        hl.ColSpann = toCol - fromCol;
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

            ReadUntil(xr, "sheetData", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks");
            ExcelCell cell = null;
            xr.Read();
            
            while (!xr.EOF)
            {
                while (xr.NodeType == XmlNodeType.EndElement)
                {
                    xr.Read();
                }
                if (xr.Name == "row")
                {
                    int row = Convert.ToInt32(xr.GetAttribute("r"));

                    if (xr.AttributeCount > 2 || (xr.AttributeCount == 2 && xr.GetAttribute("spans") != null))
                    {
                        rowList.Add(AddRow(xr, row));
                    }
                    xr.Read();
                }
                else if (xr.Name == "c")
                {
                    if (cell != null) cellList.Add(cell);
                    cell = new ExcelCell(this, xr.GetAttribute("r"));
                    if (xr.GetAttribute("t") != null) cell.DataType = xr.GetAttribute("t");
                    cell.StyleID = xr.GetAttribute("s") == null ? 0 : int.Parse(xr.GetAttribute("s"));
                    xr.Read();
                }
                else if (xr.Name == "v")
                {
                    cell._value = GetValueFromXml(cell, xr);
                    xr.Read();
                }
                else if (xr.Name == "f")
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
                    if (xr.Name != "mergeCell") break;

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
            r.Height = xr.GetAttribute("ht") == null ? defaultRowHeight : double.Parse(xr.GetAttribute("ht"), _ci);
            r.Hidden = xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1" ? true : false; ;
            r.OutlineLevel = xr.GetAttribute("outlineLevel") == null ? 0 : int.Parse(xr.GetAttribute("outlineLevel"), _ci); ;
            r.Phonetic = xr.GetAttribute("ph") != null && xr.GetAttribute("ph") == "1" ? true : false; ;
            r.StyleID = xr.GetAttribute("s") == null ? 0 : int.Parse(xr.GetAttribute("s"), _ci);
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
                value = xlPackage.Workbook._sharedStringsList[ix].Text;
                cell.IsRichText = xlPackage.Workbook._sharedStringsList[ix].isRichText;
            }
            else if (cell.DataType == "str")
            {
                value = xr.ReadElementContentAsString();
            }
            else
            {
                int n = cell.Style.Numberformat.NumFmtID;
                string v = xr.ReadElementContentAsString();

                if ((n >= 14 && n <= 22) || (n >= 45 && n <= 47))
                {
                    double res;
                    if (double.TryParse(v, out res))
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
                    if (double.TryParse(v, NumberStyles.Any, _ci, out d))
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
        private string GetSharedString(int stringID)
        {
            string retValue = null;
            XmlNodeList stringNodes = xlPackage.Workbook.SharedStringsXml.SelectNodes(string.Format("//d:si", stringID), NameSpaceManager);
            XmlNode stringNode = stringNodes[stringID];
            if (stringNode != null)
                retValue = stringNode.InnerText;
            return (retValue);
        }
        #endregion
		#region HeaderFooter
		/// <summary>
		/// A reference to the header and footer class which allows you to 
		/// set the header and footer for all odd, even and first pages of the worksheet
		/// </summary>
		public ExcelHeaderFooter HeaderFooter
		{
			get
			{
				if (_headerFooter == null)
				{
                    XmlNode headerFooterNode = TopNode.SelectSingleNode("d:headerFooter", NameSpaceManager);
                    if (headerFooterNode == null)
                        headerFooterNode= CreateNode("d:headerFooter");
                    _headerFooter = new ExcelHeaderFooter(NameSpaceManager, headerFooterNode);
				}
				return (_headerFooter);
			}
		}
		#endregion

        #region "PrinterSettings"
        public ExcelPrinterSettings PrinterSettings
        {
            get
            {
                var ps = new ExcelPrinterSettings(NameSpaceManager, TopNode);
                ps.SchemaNodeOrder = SchemaNodeOrder;
                return ps;
            }
        }
        #endregion
        // TODO: implement freeze pane. 

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
        /// Provide access to a range of cells
        /// </summary>  
        public ExcelRange Cells
        {
            get
            {
                return new ExcelRange(this/*,Address*/);
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
/*        public ExcelRange Cells(int FromCol, int FromRow, int ToCol, int ToRow)
        {
            return new ExcelRange(this, FromCol, FromRow, ToCol, ToRow);
        }*/
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
                    ExcelColumn copy = Copy(column, col+1);
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
                            ExcelColumn newC = Copy(checkColumn, col + 1);
                            newC.ColumnMax = maxCol;
                        }
                        return Copy(checkColumn, col);                        
                    }
                }
                column = new ExcelColumn(this, col);
                _columns.Add(column);
             }
            return column;
		}

        private ExcelColumn Copy(ExcelColumn c, int col)
        {
            ExcelColumn newC = new ExcelColumn(this, col);
            if (c.StyleName != "")
                newC.StyleName = c.StyleName;
            else
                newC.StyleID = c.StyleID;        
            newC.Width = c.Width;
            newC.Hidden = c.Hidden;
            _columns.Add(newC);
            return newC;
       }

        /// <summary>
        /// Selects a range in the worksheet. The actice cell is the topmost cell.
        /// Make the current worksheet active.
        /// </summary>
        /// <param name="Address">A address range</param>
        public void Select(string Address)
        {
            Select(Address, true);
        }
        /// <summary>
        /// Selects a range in the worksheet. The actice cell is the topmost cell.
        /// </summary>
        /// <param name="Address">A address range</param>
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
        ///// <summary>
        ///// Inserts conditional formatting for the cell range.
        ///// Currently only supports the dataBar style.
        ///// </summary>
        ///// <param name="startCell"></param>
        ///// <param name="endCell"></param>
        ///// <param name="color"></param>
        //internal void CreateConditionalFormatting(ExcelCell startCell, ExcelCell endCell, string color)
        //{
        //    throw(new NotImplementedException("Conditional formatting has been removed for now."));
        //}

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
            
            AddMergedCells(rowFrom, rows);

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
        private void AddMergedCells(int position, int rows)
        {
            for(int i=0;i<_mergedCells.Count;i++)
            {
                int fromRow, toRow, fromCol, toCol;
                ExcelCellBase.GetRowColFromAddress(_mergedCells[i], out fromRow, out  fromCol, out  toRow, out toCol);

                if (fromRow >= position) 
                {
                    fromRow += rows;
                }
                if (toRow >= position)
                {
                    toRow += rows;
                }

                //Set merged prop for cells
                for (int row = fromRow; row <= toRow; row++)
                {
                    for (int col = fromCol; col <= toCol; col++)
                    {
                        Cell(row, col).Merge = true;
                    }
                }

                _mergedCells.List[i] = ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol);
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
                f.Index = GetMaxShareFunctionIndex();
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
            AddMergedCells(rowFrom, -rows);
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

		#endregion // END Worksheet Public Methods

		#region Worksheet Private Methods

		#region Worksheet Save
		/// <summary>
		/// Saves the worksheet to the package.  For internal use only.
		/// </summary>
		protected internal void Save()  // Worksheet Save
		{
			#region Delete the printer settings component (if it exists)
			// we also need to delete the relationship from the pageSetup tag
			XmlNode pageSetup = WorksheetXml.SelectSingleNode("//d:pageSetup", NameSpaceManager);
			if (pageSetup != null)
			{
				XmlAttribute attr = (XmlAttribute)pageSetup.Attributes.GetNamedItem("id", ExcelPackage.schemaRelationships);
				if (attr != null)
				{
					string relID = attr.Value;
					// first delete the attribute from the XML
					pageSetup.Attributes.Remove(attr);

					// get the URI
					PackageRelationship relPrinterSettings = Part.GetRelationship(relID);
					Uri printerSettingsUri = new Uri("/xl" + relPrinterSettings.TargetUri.ToString().Replace("..", ""), UriKind.Relative);

					// now delete the relationship
					Part.DeleteRelationship(relPrinterSettings.Id);

					// now delete the part from the package
					xlPackage.Package.DeletePart(printerSettingsUri);
				}
			}
			#endregion

			if (_worksheetXml != null)
			{
                
				// save the header & footer (if defined)
				if (_headerFooter != null)
					HeaderFooter.Save();
                // replace the numeric Cell IDs we inserted with AddNumericCellIDs()
				//ReplaceNumericCellIDs();

                if (_cells.Count > 0)
                {
                    this.SetXmlNode("d:dimension/@ref", Dimension.Address);
                }

                SaveComments();
                SaveXml();
				// save worksheet to package
                //PackagePart partPack = xlPackage.Package.GetPart(WorksheetUri);
                //WorksheetXml.Save(Part.GetStream(FileMode.Create, FileAccess.Write));

                xlPackage.WriteDebugFile(WorksheetXml, @"xl\worksheets", "sheet" + SheetID + ".xml");
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
                //xlPackage.WriteDebugFile(WorksheetXml, @"xl\drawings", "drawing" + SheetID + ".xml");                
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
                        xlPackage.Package.DeletePart(_comments.Uri);                        
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
                        _comments.Part = xlPackage.Package.CreatePart(_comments.Uri, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", xlPackage.Compression);
                        var rel = Part.CreateRelationship(_comments.Uri, TargetMode.Internal, ExcelPackage.schemaRelationships+"/comments");
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
                        xlPackage.Package.DeletePart(_vmlDrawings.Uri);
                    }
                }
                else
                {
                    if (_vmlDrawings.Uri == null)
                    {
                        _vmlDrawings.Uri = new Uri(string.Format(@"/xl/drawings/vmlDrawing{0}.vml", SheetID), UriKind.Relative);
                    }
                    if (_vmlDrawings.Part == null)
                    {
                        _vmlDrawings.Part = xlPackage.Package.CreatePart(_vmlDrawings.Uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", xlPackage.Compression);
                        var rel=Part.CreateRelationship(_vmlDrawings.Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
                        SetXmlNode("d:legacyDrawing/@r:id", rel.Id);
                        _vmlDrawings.RelId = rel.Id;
                    }
                    _vmlDrawings.VmlDrawingXml.Save(_vmlDrawings.Part.GetStream());
                }
            }
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
            PackagePart partPack = xlPackage.Package.GetPart(WorksheetUri);
            StreamWriter sw=new StreamWriter(Part.GetStream(FileMode.Create, FileAccess.Write));

            int colStart=0, colEnd=0;
            GetBlockPos(xml, "cols", ref colStart, ref colEnd);

            sw.Write(xml.Substring(0, colStart));
            var colBreaks = new List<int>();
            if (_columns.Count > 0)
            {
                UpdateColumnData(sw, ref colBreaks);
            }

            int cellStart = colEnd, cellEnd = colEnd;
            GetBlockPos(xml, "sheetData", ref cellStart, ref cellEnd);
            sw.Write(xml.Substring(colEnd, cellStart - colEnd));
            var rowBreaks=new List<int>();
            UpdateRowCellData(sw, ref rowBreaks);

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
            if (rowBreaks.Count > 0)
            {
                UpdateRowBreaks(sw, ref rowBreaks);
            }

            int colBreakStart = rowBreakEnd, colBreakEnd = rowBreakEnd;
            GetBlockPos(xml, "colBreaks", ref colBreakStart, ref colBreakEnd);
            sw.Write(xml.Substring(rowBreakEnd, colBreakStart - rowBreakEnd));
            if (colBreaks.Count > 0)
            {
                UpdateColBreaks(sw, ref colBreaks);
            }

            sw.Write(xml.Substring(colBreakEnd, xml.Length - colBreakEnd));
            sw.Flush();
        }

        private void UpdateColBreaks(StreamWriter sw, ref List<int> colBreaks)
        {
            sw.Write(string.Format("<colBreaks count=\"{0}\" manualBreakCount=\"{0}\">", colBreaks.Count));
            foreach (int col in colBreaks)
            {
                sw.Write("<brk id=\"{0}\" max=\"16383\" man=\"1\" />", col);
            }
            sw.Write("</colBreaks>");
        }

        private void UpdateRowBreaks(StreamWriter sw, ref List<int> rowBreaks)
        {
            sw.Write(string.Format("<rowBreaks count=\"{0}\" manualBreakCount=\"{0}\">", rowBreaks.Count));
            foreach (int row in rowBreaks)
            {
                sw.Write("<brk id=\"{0}\" max=\"1048575\" man=\"1\" />", row);
            }
            sw.Write("</rowBreaks>");

        }
        /// <summary>
        /// Inserts the cols collection into the XML document
        /// </summary>
        private void UpdateColumnData(StreamWriter sw, ref List<int> colBreaks)
        {
            ExcelColumn prevCol = null;
            foreach (ExcelColumn col in _columns)
            {                
                if (prevCol != null)
                {
                    if(prevCol.ColumnMax != col.ColumnMin-1)
                    {
                        prevCol._columnMax=col.ColumnMin-1;
                    }
                }
                prevCol = col;
            }
            sw.Write("<cols>");
            foreach (ExcelColumn col in _columns)
            {
                ExcelStyleCollection<ExcelXfs> cellXfs = xlPackage.Workbook.Styles.CellXfs;

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
                sw.Write(string.Format(_ci, " width=\"{0}\" customWidth=\"1\"", col.Width));
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

                if (col.PageBreak)
                {
                    colBreaks.Add(col.ColumnMin);
                }
            }
            sw.Write("</cols>");
        }
        /// <summary>
        /// Insert row and cells into the XML document
        /// </summary>
        private void UpdateRowCellData(StreamWriter sw,ref List<int> rowBreaks)
        {
            ExcelStyleCollection<ExcelXfs> cellXfs = xlPackage.Workbook.Styles.CellXfs;
            
            _hyperLinkCells = new List<ulong>();
            int row = -1;

            StringBuilder sbXml = new StringBuilder();
            var ss = xlPackage.Workbook._sharedStrings;
            sw.Write("<sheetData>");
            foreach (ExcelCell cell in _cells)
            {
                //ExcelCell cell = _cells[cellID];
                long styleID = cell.StyleID >= 0 ? cellXfs[cell.StyleID].newID : cell.StyleID;
                
                //Add the row element if it's a new row
                if (row != cell.Row)
                {
                    if (row != -1) sw.Write("</row>");

                    ulong rowID = ExcelRow.GetRowID(SheetID, cell.Row);
                    sw.Write("<row r=\"{0}\" ", cell.Row);
                    if (_rows.ContainsKey(rowID))
                    {
                        ExcelRow currRow = _rows[rowID] as ExcelRow;
                        if (currRow.Hidden == true)
                        {
                            sw.Write("ht=\"0\" hidden=\"1\" ");
                        }
                        else if (currRow.Height != defaultRowHeight )
                        {
                            sw.Write(string.Format(_ci, "ht=\"{0}\" customHeight=\"1\" ", currRow.Height));
                        }   

                        if(currRow.StyleID > 0)
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
                                    //ulong prevRowID = ExcelRow.GetRowID(SheetID, currRow.Row - 1), nextRowID = ExcelRow.GetRowID(SheetID, currRow.Row+1);
                                    //ExcelRow prevRow;//, nextRow;
                                    //if (_rows.ContainsKey(prevRowID))
                                    //{
                                    //    prevRow = _rows[prevRowID] as ExcelRow;
                                    //    //nextRow = _rows[nextRowID] as ExcelRow;
                                    //    if (prevRow.Collapsed)
                                    //    {
                                            sw.Write(" collapsed=\"1\" hidden=\"1\""); //Always hidden                                        
                                    //    }
                                    //    else
                                    //    {
                                    //        sw.Write(" collapsed=\"1\""); //not hidden, only collapsed
                                    //    }
                                    //}
                                    //else
                                    //{
                                    //    sw.Write(" collapsed=\"1\""); //Always hidden
                                    //}
                                }
                            }
                        }
                        if (currRow.Phonetic)
                        {
                            sw.Write("ph=\"1\" ");
                        }
                        if (currRow.PageBreak)
                        {
                            rowBreaks.Add(currRow.Row);
                        }
                    }
                    sw.Write(">");
                    row = cell.Row;
                }
                if (cell.SharedFormulaID >= 0)
                {                    
                    var f = _sharedFormulas[cell.SharedFormulaID];
                    if (f.Address.IndexOf(':') > 0)
                    {
                        if (f.StartCol == cell.Column && f.StartRow == cell.Row)
                        {
                            sw.Write("<c r=\"{0}\" s=\"{1}\"><f ref=\"{2}\" t=\"shared\"  si=\"{3}\">{4}</f></c>", cell.CellAddress, styleID < 0 ? 0 : styleID, f.Address, cell.SharedFormulaID, SecurityElement.Escape(f.Formula));
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
                    if (cell.Value == null)
                    {
                        sw.Write("<c r=\"{0}\" s=\"{1}\" />", cell.CellAddress, styleID < 0 ? 0 : styleID);
                    }
                    else
                    {
                        if ((cell.Value.GetType().IsPrimitive || cell.Value is double || cell.Value is decimal || cell.Value is DateTime) && cell.DataType != "s")
                        {
                            string s;
                            try
                            {
                                if (cell.Value is DateTime)
                                {
                                    s = ((DateTime)cell.Value).ToOADate().ToString(_ci);
                                }
                                else
                                {
                                    if (cell.Value is double && double.IsNaN((double)cell.Value))
                                    {
                                        s = "0";
                                    }
                                    else
                                    {
                                        s = Convert.ToDecimal(cell.Value, _ci).ToString(_ci);
                                    }
                                }
                            }

                            catch
                            {
                                s = "0";
                            }
                            sw.Write("<c r=\"{0}\" s=\"{1}\">", cell.CellAddress, styleID < 0 ? 0 : styleID);
                            sw.Write("<v>{0}</v></c>", s);
                        }
                        else
                        {
                            int ix;
                            if (!ss.ContainsKey(cell.Value.ToString()))
                            {
                                ix = ss.Count;
                                ss.Add(cell.Value.ToString(), new ExcelWorkbook.SharedStringItem() { isRichText = cell.IsRichText, pos = ix });
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
            if (row != -1) sw.Write("</row>");
            sw.Write("</sheetData>");
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
                    if (cell.Hyperlink is ExcelHyperLink && (cell.Hyperlink as ExcelHyperLink).ReferenceAddress != "")
                    {
                        ExcelHyperLink hl = cell.Hyperlink as ExcelHyperLink;
                        sw.Write("<hyperlink ref=\"{0}\" location=\"{1}\" display=\"{2}\" />", 
                                Cells[cell.Row, cell.Column, cell.Row+hl.RowSpann, cell.Column+hl.ColSpann].Address, 
                                ExcelCell.GetFullAddress(Name, hl.ReferenceAddress),
                                SecurityElement.Escape(hl.Display));

                    }
                    else
                    {
                        string id;
                        if (hyps.ContainsKey(cell.Hyperlink.AbsolutePath))
                        {
                            id = hyps[cell.Hyperlink.AbsolutePath];
                        }
                        else
                        {
                            PackageRelationship relationship = Part.CreateRelationship(cell.Hyperlink, TargetMode.External, ExcelPackage.schemaHyperlink);
                            sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\" />",cell.CellAddress, relationship.Id);

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
        /// Dimension address if the worksheet. 
        /// Top left cell to Bottom right.
        /// If the worksheet has no cells, null is returned
        /// </summary>
        public ExcelAddressBase Dimension
        {
            get
            {
                if (_cells.Count > 0)
                {
                    return new ExcelAddressBase((_cells[0] as ExcelCell).Row, _minCol, (_cells[_cells.Count - 1] as ExcelCell).Row, _maxCol);
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
                    _protection = new ExcelSheetProtection(NameSpaceManager, TopNode);
                }
                return _protection;
            }
        }
        public void SetPrintArea(ExcelAddress address)
        {
            if(Names.ContainsKey("_xlnm.Print_Area"))
            {
                Names["_xlnm.Print_Area"].Address = ExcelAddress.GetFullAddress(Name, address.Address);
            }
            else
            {
                Names.Add("_xlnm.Print_Area", Cells[address.Address]);
            }
        }
        public void ClearPrintArea()
        {
            if(Names.ContainsKey("_xlnm.Print_Area"))
            {
                Names.Remove("_xlnm.Print_Area");
            }
        }
        #region Drawing
        ExcelDrawings _drawings = null;
        public ExcelDrawings Drawings
        {
            get
            {
                if (_drawings == null)
                {
                    _drawings = new ExcelDrawings(xlPackage, this);
                }
                return _drawings;
            }
        }
        #endregion

		/// <summary>
		/// Returns the style ID given a style name.  
		/// The style ID will be created if not found, but only if the style name exists!
		/// </summary>
		/// <param name="StyleName"></param>
		/// <returns></returns>
		protected internal int GetStyleID(string StyleName)
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
                return xlPackage.Workbook;
            }
        }
		#endregion
        #endregion  // END Worksheet Private Methods

        internal int GetMaxShareFunctionIndex()
        {
            int i = _sharedFormulas.Count + 1;
            while(_sharedFormulas.ContainsKey(i))
            {
                i++;
            }
            return i;
        }
    }  // END class Worksheet
}
