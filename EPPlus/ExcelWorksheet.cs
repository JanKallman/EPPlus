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
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;

using System.Linq;
using EPPlus.Core.Compatibility;

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
        //Merged = 0x1,
        RichText = 0x2,
        SharedFormula = 0x4,
        ArrayFormula = 0x8
    }
    /// <summary>
    /// For Cell value structure (for memory optimization of huge sheet)
    /// </summary>
    public struct ExcelCoreValue
    {
        internal object _value;
        internal int _styleId;

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
    public class ExcelWorksheet : XmlHelper, IEqualityComparer<ExcelWorksheet>, IDisposable
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

            private IEnumerable<Token> Tokens { get; set; }

            internal string GetFormula(int row, int column, string worksheet)
            {
                if ((StartRow == row && StartCol == column))
                {
                    return Formula;
                }

                if (Tokens == null)
                {
                    Tokens = _tokenizer.Tokenize(Formula, worksheet);
                }

                string f = "";
                foreach (var token in Tokens)
                {
                    if (token.TokenType == TokenType.ExcelAddress)
                    {
                        var a = new ExcelFormulaAddress(token.Value);
                        f += !string.IsNullOrEmpty(a._wb) || !string.IsNullOrEmpty(a._ws)
                            ? token.Value
                            : a.GetOffset(row - StartRow, column - StartCol);
                    }
                    else
                    {
                        f += token.Value;
                    }
                }
                return f;
            }

            public Formulas Clone()
            {
                var formulas = new Formulas(this._tokenizer);
                formulas.Index = this.Index;
                formulas.Address = this.Address;
                formulas.IsArray = this.IsArray;
                formulas.Formula = this.Formula;
                formulas.StartRow = this.StartRow;
                formulas.StartCol = this.StartCol;
                return formulas;
            }
        }
        /// <summary>
        /// Collection containing merged cell addresses
        /// </summary>
        public class MergeCellsCollection : IEnumerable<string>
        {
            internal MergeCellsCollection()
            {

            }
            internal CellStore<int> _cells = new CellStore<int>();
            List<string> _list = new List<string>();
            internal List<string> List { get { return _list; } }
            public string this[int row, int column]
            {
                get
                {
                    int ix = -1;
                    if (_cells.Exists(row, column, ref ix) && ix >= 0 && ix < List.Count)  //Fixes issue 15075
                    {
                        return List[ix];
                    }
                    else
                    {
                        return null;
                    }
                }
            }
            public string this[int index]
            {
                get
                {
                    return _list[index];
                }
            }
            internal void Add(ExcelAddressBase address, bool doValidate)
            {
                int ix = 0;

                //Validate
                if (doValidate && Validate(address) == false)
                {
                    throw (new ArgumentException("Can't merge and already merged range"));
                }
                lock (this)
                {
                    ix = _list.Count;
                    _list.Add(address.Address);
                    SetIndex(address, ix);
                }
            }

            private bool Validate(ExcelAddressBase address)
            {
                int ix = 0;
                if (_cells.Exists(address._fromRow, address._fromCol, ref ix))
                {
                    if (ix >= 0 && ix < _list.Count && _list[ix] != null && address.Address == _list[ix])
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                var cse = new CellsStoreEnumerator<int>(_cells, address._fromRow, address._fromCol, address._toRow, address._toCol);
                //cells
                while (cse.Next())
                {
                    return false;
                }
                //Entire column
                cse = new CellsStoreEnumerator<int>(_cells, 0, address._fromCol, 0, address._toCol);
                while (cse.Next())
                {
                    return false;
                }
                //Entire row
                cse = new CellsStoreEnumerator<int>(_cells, address._fromRow, 0, address._toRow, 0);
                while (cse.Next())
                {
                    return false;
                }
                return true;
            }

            internal void SetIndex(ExcelAddressBase address, int ix)
            {
                if (address._fromRow == 1 && address._toRow == ExcelPackage.MaxRows) //Entire row
                {
                    for (int col = address._fromCol; col <= address._toCol; col++)
                    {
                        _cells.SetValue(0, col, ix);
                    }
                }
                else if (address._fromCol == 1 && address._toCol == ExcelPackage.MaxColumns) //Entire row
                {
                    for (int row = address._fromRow; row <= address._toRow; row++)
                    {
                        _cells.SetValue(row, 0, ix);
                    }
                }
                else
                {
                    for (int col = address._fromCol; col <= address._toCol; col++)
                    {
                        for (int row = address._fromRow; row <= address._toRow; row++)
                        {
                            _cells.SetValue(row, col, ix);
                        }
                    }
                }
            }
            public int Count
            {
                get
                {
                    return _list.Count;
                }
            }
            internal void Remove(string Item)
            {
                _list.Remove(Item);
            }
            #region IEnumerable<string> Members

            public IEnumerator<string> GetEnumerator()
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
            internal void Clear(ExcelAddressBase Destination)
            {
                var cse = new CellsStoreEnumerator<int>(_cells, Destination._fromRow, Destination._fromCol, Destination._toRow, Destination._toCol);
                var used = new HashSet<int>();
                while (cse.Next())
                {
                    var v = cse.Value;
                    if (!used.Contains(v) && _list[v] != null)
                    {
                        var adr = new ExcelAddressBase(_list[v]);
                        if (!(Destination.Collide(adr) == ExcelAddressBase.eAddressCollition.Inside || Destination.Collide(adr) == ExcelAddressBase.eAddressCollition.Equal))
                        {
                            throw (new InvalidOperationException(string.Format("Can't delete/overwrite merged cells. A range is partly merged with the another merged range. {0}", adr._address)));
                        }
                        used.Add(v);
                    }
                }

                _cells.Clear(Destination._fromRow, Destination._fromCol, Destination._toRow - Destination._fromRow + 1, Destination._toCol - Destination._fromCol + 1);
                foreach (var i in used)
                {
                    _list[i] = null;
                }
            }
        }
        //internal CellStore<object> _values;
        //internal CellStore<string> _types;
        //internal CellStore<int> _styles;
        internal CellStore<ExcelCoreValue> _values;
        internal CellStore<object> _formulas;
        internal FlagCellStore _flags;
        internal CellStore<List<Token>> _formulaTokens;

        internal CellStore<Uri> _hyperLinks;
        internal CellStore<int> _commentsStore;

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
            SchemaNodeOrder = new string[] { "sheetPr", "tabColor", "outlinePr", "pageSetUpPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews", "customSheetViews", "mergeCells", "phoneticPr", "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter", "linePrint", "rowBreaks", "colBreaks", "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "activeXControls", "webPublishItems", "tableParts", "extLst" };
            _package = excelPackage;
            _relationshipID = relID;
            _worksheetUri = uriWorksheet;
            _name = sheetName;
            _sheetID = sheetID;
            _positionID = positionID;
            Hidden = hide;

            /**** Cellstore ****/
            _values = new CellStore<ExcelCoreValue>();
            //_types = new CellStore<string>();
            //_styles = new CellStore<int>();
            _formulas = new CellStore<object>();
            _flags = new FlagCellStore();
            _commentsStore = new CellStore<int>();
            _hyperLinks = new CellStore<Uri>();

            _names = new ExcelNamedRangeCollection(Workbook, this);

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
                if (value == null)
                {
                    DeleteAllNode("d:autoFilter/@ref");
                }
                else
                {
                    SetXmlNodeString("d:autoFilter/@ref", value.Address);
                }
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
                        CreateNode("d:sheetViews/d:sheetView");     //this one shouls always exist. but check anyway
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
                value = _package.Workbook.Worksheets.ValidateFixSheetName(value);
                foreach (var ws in Workbook.Worksheets)
                {
                    if (ws.PositionID != PositionID && ws.Name.Equals(value, StringComparison.OrdinalIgnoreCase))
                    {
                        throw (new ArgumentException("Worksheet name must be unique"));
                    }
                }
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
                if (string.IsNullOrEmpty(n.NameFormula) && n.NameValue == null)
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
                    ws.UpdateCrossSheetReferenceNames(_name, value);
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
                string state = _package.Workbook.GetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID));
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
                    v = value.ToString();
                    v = v.Substring(0, 1).ToLowerInvariant() + v.Substring(1);
                    _package.Workbook.SetXmlNodeString(string.Format("d:sheets/d:sheet[@sheetId={0}]/@state", _sheetID), v);
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
                _defaultRowHeight = GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight");
                if (double.IsNaN(_defaultRowHeight) || CustomHeight == false)
                {
                    _defaultRowHeight = GetRowHeightFromNormalStyle();
                }
                return _defaultRowHeight;
            }
            set
            {
                CheckSheetType();
                _defaultRowHeight = value;
                if (double.IsNaN(value))
                {
                    DeleteNode("d:sheetFormatPr/@defaultRowHeight");
                }
                else
                {
                    SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", value.ToString(CultureInfo.InvariantCulture));
                    //Check if this is the default width for the normal style
                    double defHeight = GetRowHeightFromNormalStyle();
                    CustomHeight = true;
                }
            }
        }

        private double GetRowHeightFromNormalStyle()
        {
            var ix = Workbook.Styles.NamedStyles.FindIndexByID("Normal");
            if (ix >= 0)
            {
                var f = Workbook.Styles.NamedStyles[ix].Style.Font;
                return ExcelFontXml.GetFontHeight(f.Name, f.Size) * 0.75;
            }
            else
            {
                return 15;   //Default Calibri 11
            }
        }

        /// <summary>
        /// 'True' if defaultRowHeight value has been manually set, or is different from the default value.
        /// Is automaticlly set to 'True' when assigning the DefaultRowHeight property
        /// </summary>
        public bool CustomHeight
        {
            get
            {
                return GetXmlNodeBool("d:sheetFormatPr/@customHeight");
            }
            set
            {
                SetXmlNodeBool("d:sheetFormatPr/@customHeight", value);
            }
        }
        /// <summary>
        /// Get/set the default width of all columns in the worksheet
        /// </summary>
        public double DefaultColWidth
        {
            get
            {
                CheckSheetType();
                double ret = GetXmlNodeDouble("d:sheetFormatPr/@defaultColWidth");
                if (double.IsNaN(ret))
                {
                    var mfw = Convert.ToDouble(Workbook.MaxFontWidth);
                    var widthPx = mfw * 7;
                    var margin = Math.Truncate(mfw / 4 + 0.999) * 2 + 1;
                    if (margin < 5) margin = 5;
                    while (Math.Truncate((widthPx - margin) / mfw * 100 + 0.5) / 100 < 8)
                    {
                        widthPx++;
                    }
                    widthPx = widthPx % 8 == 0 ? widthPx : 8 - widthPx % 8 + widthPx;
                    var width = Math.Truncate((widthPx - margin) / mfw * 100 + 0.5) / 100;
                    return Math.Truncate((width * mfw + margin) / mfw * 256) / 256;
                }
                return ret;
            }
            set
            {
                CheckSheetType();
                SetXmlNodeString("d:sheetFormatPr/@defaultColWidth", value.ToString(CultureInfo.InvariantCulture));

                if (double.IsNaN(GetXmlNodeDouble("d:sheetFormatPr/@defaultRowHeight")))
                {
                    SetXmlNodeString("d:sheetFormatPr/@defaultRowHeight", GetRowHeightFromNormalStyle().ToString(CultureInfo.InvariantCulture));
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

#if Core
            var xr = XmlReader.Create(stream,new XmlReaderSettings() { DtdProcessing = DtdProcessing.Prohibit, IgnoreWhitespace = true });
#else
            var xr = new XmlTextReader(stream);
            xr.ProhibitDtd = true;
            xr.WhitespaceHandling = WhitespaceHandling.None;
#endif
            LoadColumns(xr);    //columnXml
            long start = stream.Position;
            LoadCells(xr);
            var nextElementLength = GetAttributeLength(xr);
            long end = stream.Position - nextElementLength;
            LoadMergeCells(xr);
            LoadHyperLinks(xr);
            LoadRowPageBreakes(xr);
            LoadColPageBreakes(xr);
            //...then the rest of the Xml is extracted and loaded into the WorksheetXml document.
            stream.Seek(0, SeekOrigin.Begin);
            Encoding encoding;
            xml = GetWorkSheetXml(stream, start, end, out encoding);

            // now release stream buffer (already converted whole Xml into XmlDocument Object and String)
            stream.Dispose();
            packPart.Stream = new MemoryStream();

            //first char is invalid sometimes?? 
            if (xml[0] != '<')
                LoadXmlSafe(_worksheetXml, xml.Substring(1, xml.Length - 1), encoding);
            else
                LoadXmlSafe(_worksheetXml, xml, encoding);

            _package.DoAdjustDrawings = doAdjust;
            ClearNodes();
        }
        /// <summary>
        /// Get the lenth of the attributes
        /// Conditional formatting attributes can be extremly long som get length of the attributes to finetune position.
        /// </summary>
        /// <param name="xr"></param>
        /// <returns></returns>
        private int GetAttributeLength(XmlReader xr)
        {
            if (xr.NodeType != XmlNodeType.Element) return 0;
            var length = 0;

            for (int i = 0; i < xr.AttributeCount; i++)
            {
                var a = xr.GetAttribute(i);
                length += string.IsNullOrEmpty(a) ? 0 : a.Length;
            }
            return length;
        }
        private void LoadRowPageBreakes(XmlReader xr)
        {
            if (!ReadUntil(xr, "rowBreaks", "colBreaks")) return;
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
        private void LoadColPageBreakes(XmlReader xr)
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
            if (_worksheetXml.SelectSingleNode("//d:cols", NameSpaceManager) != null)
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
        const int BLOCKSIZE = 8192;
        /// <summary>
        /// Extracts the workbook XML without the sheetData-element (containing all cell data).
        /// Xml-Cell data can be extreemly large (GB), so we find the sheetdata element in the streem (position start) and 
        /// then tries to find the &lt;/sheetData&gt; element from the end-parameter.
        /// This approach is to avoid out of memory exceptions reading large packages
        /// </summary>
        /// <param name="stream">the worksheet stream</param>
        /// <param name="start">Position from previous reading where we found the sheetData element</param>
        /// <param name="end">End position, where &lt;/sheetData&gt; or &lt;sheetData/&gt; is found</param>
        /// <param name="encoding">Encoding</param>
        /// <returns>The worksheet xml, with an empty sheetdata. (Sheetdata is in memory in the worksheet)</returns>
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
                sb.Append(block, 0, pos);
                length += size;
            }
            while (length < start + 20 && length < end);    //the  start-pos contains the stream position of the sheetData element. Add 20 (with some safty for whitespace, streampointer diff etc, just so be sure). 
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
                if (Utils.ConvertUtil._invariantCompareInfo.IsSuffix(startmMatch.Value, "/>"))        //Empty sheetdata
                {
                    xml += s.Substring(startmMatch.Index, s.Length - startmMatch.Index);
                }
                else
                {
                    if (sr.Peek() != -1)        //Now find the end tag </sheetdata> so we can add the end of the xml document
                    {
                        /**** Fixes issue 14788. Fix by Philip Garrett ****/
                        long endSeekStart = end;

                        while (endSeekStart >= 0)
                        {
                            endSeekStart = Math.Max(endSeekStart - BLOCKSIZE, 0);
                            int size = (int)(end - endSeekStart);
                            stream.Seek(endSeekStart, SeekOrigin.Begin);
                            block = new char[size];
                            sr = new StreamReader(stream);
                            pos = sr.ReadBlock(block, 0, size);
                            sb = new StringBuilder();
                            sb.Append(block, 0, pos);
                            s = sb.ToString();
                            endMatch = Regex.Match(s, string.Format("(</[^>]*{0}[^>]*>)", "sheetData"));
                            if (endMatch.Success)
                            {
                                break;
                            }
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
            var startPos = startmMatch.Index + start;
            if (startmMatch.Value.Substring(startmMatch.Value.Length - 2, 1) == "/")
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
        private bool ReadUntil(XmlReader xr, params string[] tagName)
        {
            if (xr.EOF) return false;
            while (!Array.Exists(tagName, tag => Utils.ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tag)))
            {
                xr.Read();
                if (xr.EOF) return false;
            }
            return (Utils.ConvertUtil._invariantCompareInfo.IsSuffix(xr.LocalName, tagName[0]));
        }
        private void LoadColumns(XmlReader xr)//(string xml)
        {
            var colList = new List<IRangeID>();
            if (ReadUntil(xr, "cols", "sheetData"))
            {
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Whitespace) continue;
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
                        SetValueInner(0, min, col);

                        int style;
                        if (!(xr.GetAttribute("style") == null || !int.TryParse(xr.GetAttribute("style"), out style)))
                        {
                            SetStyleInner(0, min, style);
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
        private static bool ReadXmlReaderUntil(XmlReader xr, string nodeText, string altNode)
        {
            do
            {
                if (xr.LocalName == nodeText || xr.LocalName == altNode) return true;
            }
            while (xr.Read());
#if !Core
            xr.Close();
#endif
            return false;
        }
        /// <summary>
        /// Load Hyperlinks
        /// </summary>
        /// <param name="xr">The reader</param>
        private void LoadHyperLinks(XmlReader xr)
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
                            try
                            {
                                hl = new ExcelHyperLink(uri.AbsoluteUri);
                            }
                            catch
                            {
                                hl = new ExcelHyperLink(uri.OriginalString, UriKind.Absolute);
                            }
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
                        hl = GetHyperlinkFromRef(xr, "location", fromRow, toRow, fromCol, toCol);
                    }
                    else if (xr.GetAttribute("ref") != null)
                    {
                        hl = GetHyperlinkFromRef(xr, "ref", fromRow, toRow, fromCol, toCol);
                    }
                    else
                    {
                        // not enough info to create a hyperlink
                        break;
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

        private ExcelHyperLink GetHyperlinkFromRef(XmlReader xr, string refTag, int fromRow = 0, int toRow = 0, int fromCol = 0, int toCol = 0)
        {
            var hl = new ExcelHyperLink(xr.GetAttribute(refTag), xr.GetAttribute("display"));
            hl.RowSpann = toRow - fromRow;
            hl.ColSpann = toCol - fromCol;
            return hl;
        }

        /// <summary>
        /// Load cells
        /// </summary>
        /// <param name="xr">The reader</param>
        private void LoadCells(XmlReader xr)
        {
            ReadUntil(xr, "sheetData", "mergeCells", "hyperlinks", "rowBreaks", "colBreaks");
            ExcelAddressBase address = null;
            string type = "";
            int style = 0;
            int row = 0;
            int col = 0;
            xr.Read();

            while (!xr.EOF)
            {
                while (xr.NodeType == XmlNodeType.EndElement)
                {
                    xr.Read();
                    continue;
                }
                if (xr.LocalName == "row")
                {
                    col = 0;
                    var r = xr.GetAttribute("r");
                    if (r == null)
                    {
                        row++;
                    }
                    else
                    {
                        row = Convert.ToInt32(r);
                    }

                    if (DoAddRow(xr))
                    {
                        SetValueInner(row, 0, AddRow(xr, row));
                        if (xr.GetAttribute("s") != null)
                        {
                            SetStyleInner(row, 0, int.Parse(xr.GetAttribute("s"), CultureInfo.InvariantCulture));
                        }
                    }
                    xr.Read();
                }
                else if (xr.LocalName == "c")
                {
                    //if (cell != null) cellList.Add(cell);
                    //cell = new ExcelCell(this, xr.GetAttribute("r"));
                    var r = xr.GetAttribute("r");
                    if (r == null)
                    {
                        //Handle cells with no reference
                        col++;
                        address = new ExcelAddressBase(row, col, row, col);
                    }
                    else
                    {
                        address = new ExcelAddressBase(r);
                        col = address._fromCol;
                    }


                    //Datetype
                    if (xr.GetAttribute("t") != null)
                    {
                        type = xr.GetAttribute("t");
                        //_types.SetValue(address._fromRow, address._fromCol, type); 
                    }
                    else
                    {
                        type = "";
                    }
                    //Style
                    if (xr.GetAttribute("s") != null)
                    {
                        style = int.Parse(xr.GetAttribute("s"));
                        SetStyleInner(address._fromRow, address._fromCol, style);
                        //SetValueInner(address._fromRow, address._fromCol, null); //TODO:Better Performance ??
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
                        SetValueInner(address._fromRow, address._fromCol, null);
                        //formulaList.Add(cell);
                    }
                    else if (t == "shared")
                    {

                        string si = xr.GetAttribute("si");
                        if (si != null)
                        {
                            var sfIndex = int.Parse(si);
                            _formulas.SetValue(address._fromRow, address._fromCol, sfIndex);
                            SetValueInner(address._fromRow, address._fromCol, null);
                            string fAddress = xr.GetAttribute("ref");
                            string formula = ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString());
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
                        _formulas.SetValue(address._fromRow, address._fromCol, afIndex);
                        SetValueInner(address._fromRow, address._fromCol, null);
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
                        SetValueInner(address._fromRow, address._fromCol, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
                        //cell._value = xr.ReadInnerXml();
                    }
                    else
                    {
                        if (xr.LocalName == "r")
                        {
                            var rXml = xr.ReadOuterXml();
                            while (xr.LocalName == "r")
                            {
                                rXml += xr.ReadOuterXml();
                            }
                            SetValueInner(address._fromRow, address._fromCol, rXml);
                        }
                        else
                        {
                            SetValueInner(address._fromRow, address._fromCol, xr.ReadOuterXml());
                        }
                        //_types.SetValue(address._fromRow, address._fromCol, "rt");
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

        private bool DoAddRow(XmlReader xr)
        {
            var c = xr.GetAttribute("r") == null ? 0 : 1;
            if (xr.GetAttribute("spans") != null)
            {
                c++;
            }
            return xr.AttributeCount > c;
        }
        /// <summary>
        /// Load merged cells
        /// </summary>
        /// <param name="xr"></param>
        private void LoadMergeCells(XmlReader xr)
        {
            if (ReadUntil(xr, "mergeCells", "hyperlinks", "rowBreaks", "colBreaks") && !xr.EOF)
            {
                while (xr.Read())
                {
                    if (xr.LocalName != "mergeCell") break;
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        string address = xr.GetAttribute("ref");
                        //int fromRow, fromCol, toRow, toCol;
                        //ExcelCellBase.GetRowColFromAddress(address, out fromRow, out fromCol, out toRow, out toCol);
                        //for (int row = fromRow; row <= toRow; row++)
                        //{
                        //    for (int col = fromCol; col <= toCol; col++)
                        //    {
                        //        _flags.SetFlagValue(row, col, true,CellFlags.Merged);
                        //    }
                        //}
                        //_mergedCells.List.Add(address);
                        _mergedCells.Add(new ExcelAddress(address), false);
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
        private RowInternal AddRow(XmlReader xr, int row)
        {
            return new RowInternal()
            {
                Collapsed = (xr.GetAttribute("collapsed") != null && xr.GetAttribute("collapsed") == "1" ? true : false),
                OutlineLevel = (xr.GetAttribute("outlineLevel") == null ? (short)0 : short.Parse(xr.GetAttribute("outlineLevel"), CultureInfo.InvariantCulture)),
                Height = (xr.GetAttribute("ht") == null ? -1 : double.Parse(xr.GetAttribute("ht"), CultureInfo.InvariantCulture)),
                Hidden = (xr.GetAttribute("hidden") != null && xr.GetAttribute("hidden") == "1" ? true : false),
                Phonetic = xr.GetAttribute("ph") != null && xr.GetAttribute("ph") == "1" ? true : false,
                CustomHeight = xr.GetAttribute("customHeight") == null ? false : xr.GetAttribute("customHeight") == "1"
            };
        }

        private void SetValueFromXml(XmlReader xr, string type, int styleID, int row, int col)
        {
            //XmlNode vnode = colNode.SelectSingleNode("d:v", NameSpaceManager);
            //if (vnode == null) return null;
            if (type == "s")
            {
                int ix = xr.ReadElementContentAsInt();
                SetValueInner(row, col, _package.Workbook._sharedStringsList[ix].Text);
                if (_package.Workbook._sharedStringsList[ix].isRichText)
                {
                    _flags.SetFlagValue(row, col, true, CellFlags.RichText);
                }
            }
            else if (type == "str")
            {
                SetValueInner(row, col, ConvertUtil.ExcelDecodeString(xr.ReadElementContentAsString()));
            }
            else if (type == "b")
            {
                SetValueInner(row, col, (xr.ReadElementContentAsString() != "0"));
            }
            else if (type == "e")
            {
                SetValueInner(row, col, GetErrorType(xr.ReadElementContentAsString()));
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
                        if (Workbook.Date1904)
                        {
                            res += ExcelWorkbook.date1904Offset;
                        }
                        if (res >= -657435.0 && res < 2958465.9999999)
                        {
                            SetValueInner(row, col, DateTime.FromOADate(res));
                        }
                        else
                        {
                            SetValueInner(row, col, res);
                        }
                    }
                    else
                    {
                        SetValueInner(row, col, v);
                    }
                }
                else
                {
                    double d;
                    if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                    {
                        SetValueInner(row, col, d);
                    }
                    else
                    {
                        SetValueInner(row, col, double.NaN);
                    }
                }
            }
        }

        private object GetErrorType(string v)
        {
            return ExcelErrorValue.Parse(ConvertUtil._invariantTextInfo.ToUpper(v));
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
                        headerFooterNode = CreateNode("d:headerFooter");
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
        MergeCellsCollection _mergedCells = new MergeCellsCollection();
        /// <summary>
        /// Addresses to merged ranges
        /// </summary>
        public MergeCellsCollection MergedCells
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
            //var v = GetValueInner(row, 0);
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
            if (row < 1 || row > ExcelPackage.MaxRows)
            {
                throw (new ArgumentException("Row number out of bounds"));
            }
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
            if (col < 1 || col > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentException("Column number out of bounds"));
            }
            var column = GetValueInner(0, col) as ExcelColumn;
            if (column != null)
            {

                if (column.ColumnMin != column.ColumnMax)
                {
                    int maxCol = column.ColumnMax;
                    column.ColumnMax = col;
                    ExcelColumn copy = CopyColumn(column, col + 1, maxCol);
                }
            }
            else
            {
                int r = 0, c = col;
                if (_values.PrevCell(ref r, ref c))
                {
                    column = GetValueInner(0, c) as ExcelColumn;
                    int maxCol = column.ColumnMax;
                    if (maxCol >= col)
                    {
                        column.ColumnMax = col - 1;
                        if (maxCol > col)
                        {
                            ExcelColumn newC = CopyColumn(column, col + 1, maxCol);
                        }
                        return CopyColumn(column, col, col);
                    }
                }

                column = new ExcelColumn(this, col);
                SetValueInner(0, col, column);
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
            newC.ColumnMax = maxCol < ExcelPackage.MaxColumns ? maxCol : ExcelPackage.MaxColumns;
            if (c.StyleName != "")
                newC.StyleName = c.StyleName;
            else
                newC.StyleID = c.StyleID;

            newC.OutlineLevel = c.OutlineLevel;
            newC.Phonetic = c.Phonetic;
            newC.BestFit = c.BestFit;
            //_columns.Add(newC);
            SetValueInner(0, col, newC);
            newC._width = c._width;
            newC._hidden = c._hidden;
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
		public void InsertRow(int rowFrom, int rows, int copyStylesFromRow)
        {
            CheckSheetType();
            var d = Dimension;

            if (rowFrom < 1)
            {
                throw (new ArgumentOutOfRangeException("rowFrom can't be lesser that 1"));
            }

            //Check that cells aren't shifted outside the boundries
            if (d != null && d.End.Row > rowFrom && d.End.Row + rows > ExcelPackage.MaxRows)
            {
                throw (new ArgumentOutOfRangeException("Can't insert. Rows will be shifted outside the boundries of the worksheet."));
            }

            lock (this)
            {
                _values.Insert(rowFrom, 0, rows, 0);
                _formulas.Insert(rowFrom, 0, rows, 0);
                _commentsStore.Insert(rowFrom, 0, rows, 0);
                _hyperLinks.Insert(rowFrom, 0, rows, 0);
                _flags.Insert(rowFrom, 0, rows, 0);
                Comments.Insert(rowFrom, 0, rows, 0);
                _names.Insert(rowFrom, 0, rows, 0);
                Workbook.Names.Insert(rowFrom, 0, rows, 0, n => n.Worksheet == this);

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
                    f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, 0, rowFrom, 0, this.Name, this.Name);
                }
                var cse = new CellsStoreEnumerator<object>(_formulas);
                while (cse.Next())
                {
                    if (cse.Value is string)
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), rows, 0, rowFrom, 0, this.Name, this.Name);
                    }
                }
                FixMergedCellsRow(rowFrom, rows, false);
                if (copyStylesFromRow > 0)
                {
                    var cseS = new CellsStoreEnumerator<ExcelCoreValue>(_values, copyStylesFromRow, 0, copyStylesFromRow, ExcelPackage.MaxColumns); //Fixes issue 15068 , 15090
                    while (cseS.Next())
                    {
                        if (cseS.Value._styleId == 0) continue;
                        for (var r = 0; r < rows; r++)
                        {
                            SetStyleInner(rowFrom + r, cseS.Column, cseS.Value._styleId);
                        }
                    }
                    var newOutlineLevel = this.Row(copyStylesFromRow + rows).OutlineLevel;
                    for (var r = 0; r < rows; r++)
                    {
                        this.Row(rowFrom + r).OutlineLevel = newOutlineLevel;
                    }
                }
                foreach (var tbl in Tables)
                {
                    tbl.Address = tbl.Address.AddRow(rowFrom, rows);
                }
                foreach (var ptbl in PivotTables)
                {
                    ptbl.Address = ptbl.Address.AddRow(rowFrom, rows);
                    ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.AddRow(rowFrom, rows).Address;
                }
                
                //Issue 15573
                foreach (ExcelDataValidation dv in DataValidations)
                {
                    var addr=dv.Address;
                    var newAddr =addr.AddRow(rowFrom, rows).Address;
                    if (addr.Address != newAddr)
                    {
                        dv.SetAddress(newAddr);
                    }
                }
            }

            // Update cross-sheet references.
            foreach (var sheet in Workbook.Worksheets.Where(sheet => sheet != this))
            {
                sheet.UpdateCrossSheetReferences(this.Name, rowFrom, rows, 0, 0);
            }
        }
        /// <summary>
        /// Inserts a new column into the spreadsheet.  Existing columns below the position are 
        /// shifted down.  All formula are updated to take account of the new column.
        /// </summary>
        /// <param name="columnFrom">The position of the new column</param>
        /// <param name="columns">Number of columns to insert</param>        
        public void InsertColumn(int columnFrom, int columns)
        {
            InsertColumn(columnFrom, columns, 0);
        }
        ///<summary>
        /// Inserts a new column into the spreadsheet.  Existing column to the left are 
        /// shifted.  All formula are updated to take account of the new column.
        /// </summary>
        /// <param name="columnFrom">The position of the new column</param>
        /// <param name="columns">Number of columns to insert.</param>
        /// <param name="copyStylesFromColumn">Copy Styles from this column. Applied to all inserted columns</param>
        public void InsertColumn(int columnFrom, int columns, int copyStylesFromColumn)
        {
            CheckSheetType();
            var d = Dimension;

            if (columnFrom < 1)
            {
                throw (new ArgumentOutOfRangeException("columnFrom can't be lesser that 1"));
            }
            //Check that cells aren't shifted outside the boundries
            if (d != null && d.End.Column > columnFrom && d.End.Column + columns > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentOutOfRangeException("Can't insert. Columns will be shifted outside the boundries of the worksheet."));
            }

            lock (this)
            {
                _values.Insert(0, columnFrom, 0, columns);
                _formulas.Insert(0, columnFrom, 0, columns);
                //_styles.Insert(0, columnFrom, 0, columns);
                //_types.Insert(0, columnFrom, 0, columns);
                _commentsStore.Insert(0, columnFrom, 0, columns);
                _hyperLinks.Insert(0, columnFrom, 0, columns);
                _flags.Insert(0, columnFrom, 0, columns);
                _names.Insert(0, columnFrom, 0, columns);
                Comments.Insert(0, columnFrom, 0, columns);
                Workbook.Names.Insert(0, columnFrom, 0, columns, n => n.Worksheet == this);

                foreach (var f in _sharedFormulas.Values)
                {
                    if (f.StartCol >= columnFrom) f.StartCol += columns;
                    var a = new ExcelAddressBase(f.Address);
                    if (a._fromCol >= columnFrom)
                    {
                        a._fromCol += columns;
                        a._toCol += columns;
                    }
                    else if (a._toCol >= columnFrom)
                    {
                        a._toCol += columns;
                    }
                    f.Address = ExcelAddressBase.GetAddress(a._fromRow, a._fromCol, a._toRow, a._toCol);
                    f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, 0, columns, 0, columnFrom, this.Name, this.Name);
                }

                var cse = new CellsStoreEnumerator<object>(_formulas);
                while (cse.Next())
                {
                    if (cse.Value is string)
                    {
                        cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), 0, columns, 0, columnFrom, this.Name, this.Name);
                    }
                }

                FixMergedCellsColumn(columnFrom, columns, false);

                var csec = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 1, 0, ExcelPackage.MaxColumns);
                var lst = new List<ExcelColumn>();
                foreach (var val in csec)
                {
                    var col = val._value;
                    if (col is ExcelColumn)
                    {
                        lst.Add((ExcelColumn)col);
                    }
                }

                for (int i = lst.Count - 1; i >= 0; i--)
                {
                    var c = lst[i];
                    if (c._columnMin >= columnFrom)
                    {
                        if (c._columnMin + columns <= ExcelPackage.MaxColumns)
                        {
                            c._columnMin += columns;
                        }
                        else
                        {
                            c._columnMin = ExcelPackage.MaxColumns;
                        }

                        if (c._columnMax + columns <= ExcelPackage.MaxColumns)
                        {
                            c._columnMax += columns;
                        }
                        else
                        {
                            c._columnMax = ExcelPackage.MaxColumns;
                        }
                    }
                    else if (c._columnMax >= columnFrom)
                    {
                        var cc = c._columnMax - columnFrom;
                        c._columnMax = columnFrom - 1;
                        CopyColumn(c, columnFrom + columns, columnFrom + columns + cc);
                    }
                }


                //Copy style from another column?
                if (copyStylesFromColumn > 0)
                {
                    if (copyStylesFromColumn >= columnFrom)
                    {
                        copyStylesFromColumn += columns;
                    }

                    //Get styles to a cached list, 
                    var l = new List<int[]>();
                    var sce = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, copyStylesFromColumn, ExcelPackage.MaxRows, copyStylesFromColumn);
                    lock (sce)
                    {
                        while (sce.Next())
                        {
                            if (sce.Value._styleId == 0) continue;
                            l.Add(new int[] { sce.Row, sce.Value._styleId });
                        }
                    }

                    //Set the style id's from the list.
                    foreach (var sc in l)
                    {
                        for (var c = 0; c < columns; c++)
                        {
                            if (sc[0] == 0)
                            {
                                var col = Column(columnFrom + c);   //Create the column
                                col.StyleID = sc[1];
                            }
                            else
                            {
                                SetStyleInner(sc[0], columnFrom + c, sc[1]);
                            }
                        }
                    }
                    var newOutlineLevel = this.Column(copyStylesFromColumn).OutlineLevel;
                    for (var c = 0; c < columns; c++)
                    {
                        this.Column(columnFrom + c).OutlineLevel = newOutlineLevel;
                    }
                }
                //Adjust tables
                foreach (var tbl in Tables)
                {
                    if (columnFrom > tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
                    {
                        InsertTableColumns(columnFrom, columns, tbl);
                    }

                    tbl.Address = tbl.Address.AddColumn(columnFrom, columns);
                }
                foreach (var ptbl in PivotTables)
                {
                    if (columnFrom <= ptbl.Address.End.Column)
                    {
                        ptbl.Address = ptbl.Address.AddColumn(columnFrom, columns);
                    }
                    if (columnFrom <= ptbl.CacheDefinition.SourceRange.End.Column)
                    {
                        if (ptbl.CacheDefinition.CacheSource == eSourceType.Worksheet)
                        {
                            ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.AddColumn(columnFrom, columns).Address;
                        }
                    }
                }

                //Issue 15573
                foreach (ExcelDataValidation dv in DataValidations)
                {
                    var addr = dv.Address;
                    var newAddr = addr.AddColumn(columnFrom, columns).Address;
                    if (addr.Address != newAddr)
                    {
                        dv.SetAddress(newAddr);
                    }
                }

                // Update cross-sheet references.
                foreach (var sheet in Workbook.Worksheets.Where(sheet => sheet != this))
                {
                    sheet.UpdateCrossSheetReferences(this.Name, 0, 0, columnFrom, columns);
                }
            }
        } 
        private static void InsertTableColumns(int columnFrom, int columns, ExcelTable tbl)
        {
            var node = tbl.Columns[0].TopNode.ParentNode;
            var ix = columnFrom - tbl.Address.Start.Column - 1;
            var insPos = node.ChildNodes[ix];
            ix += 2;
            for (int i = 0; i < columns; i++)
            {
                var name =
                    tbl.Columns.GetUniqueName(string.Format("Column{0}",
                        (ix++).ToString(CultureInfo.InvariantCulture)));
                XmlElement tableColumn =
                    (XmlElement) tbl.TableXml.CreateNode(XmlNodeType.Element, "tableColumn", ExcelPackage.schemaMain);
                tableColumn.SetAttribute("id", (tbl.Columns.Count + i + 1).ToString(CultureInfo.InvariantCulture));
                tableColumn.SetAttribute("name", name);
                insPos = node.InsertAfter(tableColumn, insPos);
            } //Create tbl Column
            tbl._cols = new ExcelTableColumnCollection(tbl);
        }

        /// <summary>
        /// Adds a value to the row of merged cells to fix for inserts or deletes
        /// </summary>
        /// <param name="row"></param>
        /// <param name="rows"></param> 
        /// <param name="delete"></param>
        private void FixMergedCellsRow(int row, int rows, bool delete)
        {
            if (delete)
            {
                _mergedCells._cells.Delete(row, 0, rows, 0);
            }
            else
            {
                _mergedCells._cells.Insert(row, 0, rows, 0);
            }

            List<int> removeIndex = new List<int>();
            for (int i = 0; i < _mergedCells.Count; i++)
            {
                if (!string.IsNullOrEmpty( _mergedCells[i]))
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
                        if (newAddr.Address != addr.Address)
                        {
                        //    _mergedCells._cells.Insert(row, 0, rows, 0);
                            _mergedCells.SetIndex(newAddr, i);
                        }
                    }

                    if (newAddr.Address != addr.Address)
                    {
                        _mergedCells.List[i] = newAddr._address;
                    }
                }
            }
            for (int i = removeIndex.Count - 1; i >= 0; i--)
            {
                _mergedCells.List.RemoveAt(removeIndex[i]);
            }
        }
        /// <summary>
        /// Adds a value to the row of merged cells to fix for inserts or deletes
        /// </summary>
        /// <param name="column"></param>
        /// <param name="columns"></param>
        /// <param name="delete"></param>
        private void FixMergedCellsColumn(int column, int columns, bool delete)
        {
            if (delete)
            {
                _mergedCells._cells.Delete(0, column, 0, columns);
            }
            else
            {
                _mergedCells._cells.Insert(0, column, 0, columns);                
            }
            List<int> removeIndex = new List<int>();
            for (int i = 0; i < _mergedCells.Count; i++)
            {
                if (!string.IsNullOrEmpty(_mergedCells[i]))
                {
                    ExcelAddressBase addr = new ExcelAddressBase(_mergedCells[i]), newAddr;
                    if (delete)
                    {
                        newAddr = addr.DeleteColumn(column, columns);
                        if (newAddr == null)
                        {
                            removeIndex.Add(i);
                            continue;
                        }
                    }
                    else
                    {
                        newAddr = addr.AddColumn(column, columns);
                        if (newAddr.Address != addr.Address)
                        {
                            _mergedCells.SetIndex(newAddr, i);
                        }
                    }

                    if (newAddr.Address != addr.Address)
                    {
                        _mergedCells.List[i] = newAddr._address;
                    }
                }
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
                        newFormula = ExcelCellBase.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(formualR1C1, row, col), rows, 0, startRow, 0, this.Name, this.Name);
                        currentFormulaR1C1 = ExcelRangeBase.TranslateToR1C1(newFormula, row, col);
                    }
                    else
                    {
                        newFormula = ExcelCellBase.UpdateFormulaReferences(ExcelCellBase.TranslateFromR1C1(formualR1C1, row-rows, col), rows, 0, startRow, 0, this.Name, this.Name);
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
        /// Delete the specified row from the worksheet.
        /// </summary>
        /// <param name="row">A row to be deleted</param>
        public void DeleteRow(int row)
        {
            DeleteRow(row, 1);
        }
        /// <summary>
        /// Delete the specified row from the worksheet.
        /// </summary>
        /// <param name="rowFrom">The start row</param>
        /// <param name="rows">Number of rows to delete</param>
        public void DeleteRow(int rowFrom, int rows)
        {
            CheckSheetType();
            if (rowFrom < 1 || rowFrom + rows > ExcelPackage.MaxRows)
            {
                throw(new ArgumentException("Row out of range. Spans from 1 to " + ExcelPackage.MaxRows.ToString(CultureInfo.InvariantCulture)));
            }
            lock (this)
            {
                _values.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
                _formulas.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
                _flags.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
                _commentsStore.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
                _hyperLinks.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
                _names.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);

                Comments.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns);
                Workbook.Names.Delete(rowFrom, 0, rows, ExcelPackage.MaxColumns, n => n.Worksheet == this);

                AdjustFormulasRow(rowFrom, rows);
                FixMergedCellsRow(rowFrom, rows, true);

                foreach (var tbl in Tables)
                {
                    tbl.Address = tbl.Address.DeleteRow(rowFrom, rows);
                }
                foreach (var ptbl in PivotTables)
                {
                    if (ptbl.Address.Start.Row > rowFrom + rows)
                    {
                        ptbl.Address = ptbl.Address.DeleteRow(rowFrom, rows);
                    }
                }
                //Issue 15573
                foreach (ExcelDataValidation dv in DataValidations)
                {
                    var addr = dv.Address;
                    if (addr.Start.Row > rowFrom + rows)
                    {
                        var newAddr = addr.DeleteRow(rowFrom, rows).Address;
                        if (addr.Address != newAddr)
                        {
                            dv.SetAddress(newAddr);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Delete the specified column from the worksheet.
        /// </summary>
        /// <param name="column">The column to be deleted</param>
        public void DeleteColumn(int column)
        {
            DeleteColumn(column,1);
        }
        /// <summary>
        /// Delete the specified column from the worksheet.
        /// </summary>
        /// <param name="columnFrom">The start column</param>
        /// <param name="columns">Number of columns to delete</param>
        public void DeleteColumn(int columnFrom, int columns)
        {
            if (columnFrom < 1 || columnFrom + columns > ExcelPackage.MaxColumns)
            {
                throw (new ArgumentException("Column out of range. Spans from 1 to " + ExcelPackage.MaxColumns.ToString(CultureInfo.InvariantCulture)));
            }
            lock (this)
            {
                var col = GetValueInner(0, columnFrom) as ExcelColumn;
                if (col == null)
                {
                    var r = 0; 
                    var c = columnFrom;
                    if(_values.PrevCell(ref r,ref c))
                    {
                        col = GetValueInner(0, c) as ExcelColumn;
                        if(col._columnMax >= columnFrom)
                        {
                            col.ColumnMax=columnFrom-1;
                        }
                    }
                }

                _values.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
                //_types.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
                _formulas.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
                //_styles.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
                _flags.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
                _commentsStore.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
                _hyperLinks.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);
                _names.Delete(0, columnFrom, ExcelPackage.MaxRows, columns);

                Comments.Delete(0, columnFrom, 0, columns);
                Workbook.Names.Delete(0, columnFrom, ExcelPackage.MaxRows, columns, n => n.Worksheet == this);

                AdjustFormulasColumn(columnFrom, columns);
                FixMergedCellsColumn(columnFrom, columns, true);

                var csec = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, columnFrom, 0, ExcelPackage.MaxColumns);
                foreach (var val in csec)    
                {
                    var column = val._value;
                    if (column is ExcelColumn)
                    {
                        var c = (ExcelColumn)column;
                        if (c._columnMin >= columnFrom)
                        {
                            c._columnMin -= columns;
                            c._columnMax -= columns;
                        }
                    }
                }

                foreach (var tbl in Tables)
                {
                    if (columnFrom >= tbl.Address.Start.Column && columnFrom <= tbl.Address.End.Column)
                    {
                        var node = tbl.Columns[0].TopNode.ParentNode;
                        var ix = columnFrom - tbl.Address.Start.Column;
                        for (int i = 0; i < columns; i++)
                        {
                            if (node.ChildNodes.Count > ix)
                            {
                                node.RemoveChild(node.ChildNodes[ix]);
                            }
                        } 
                        tbl._cols = new ExcelTableColumnCollection(tbl);
                    }

                    tbl.Address = tbl.Address.DeleteColumn(columnFrom, columns);

                    foreach (var ptbl in PivotTables)
                    {
                        if (ptbl.Address.Start.Column > columnFrom + columns)
                        {
                            ptbl.Address = ptbl.Address.DeleteColumn(columnFrom, columns);
                        }
                        if (ptbl.CacheDefinition.SourceRange.Start.Column > columnFrom + columns)
                        {
                            ptbl.CacheDefinition.SourceRange.Address = ptbl.CacheDefinition.SourceRange.DeleteColumn(columnFrom, columns).Address;
                        }
                    }
                }

                //Issue 15573
                foreach (ExcelDataValidation dv in DataValidations)
                {
                    var addr = dv.Address;
                    if (addr.Start.Column > columnFrom + columns)
                    {
                        var newAddr = addr.DeleteColumn(columnFrom, columns).Address;
                        if (addr.Address != newAddr)
                        {
                            dv.SetAddress(newAddr);
                        }
                    }
                }
            }
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
                    if (sf.StartRow > rowFrom)
                    {
                        var r = Math.Min(sf.StartRow - rowFrom, rows);
                        sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, -r, 0, rowFrom, 0, this.Name, this.Name);
                        sf.StartRow -= r;
                    }
                }
            }
            foreach (var ix in delSF)
            {
                _sharedFormulas.Remove(ix);
            }
            delSF = null;
            var cse = new CellsStoreEnumerator<object>(_formulas, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Value is string)
                {
                    cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), -rows, 0, rowFrom, 0, this.Name, this.Name);
                }
            }
        }
        internal void AdjustFormulasColumn(int columnFrom, int columns)
        {
            var delSF = new List<int>();
            foreach (var sf in _sharedFormulas.Values)
            {
                var a = new ExcelAddress(sf.Address).DeleteColumn(columnFrom, columns);
                if (a == null)
                {
                    delSF.Add(sf.Index);
                }
                else
                {
                    sf.Address = a.Address;
                    //sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, 0, -columns, 0, columnFrom);
                    if (sf.StartCol > columnFrom)
                    {
                        var c = Math.Min(sf.StartCol - columnFrom, columns);
                        sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, 0, -c, 0, 1, this.Name, this.Name);
                        sf.StartCol-= c;
                    }

                    //sf.Address = a.Address;
                    //sf.Formula = ExcelCellBase.UpdateFormulaReferences(sf.Formula, 0,-columns,0, columnFrom);
                    //if (sf.StartCol >= columnFrom)
                    //{
                    //    sf.StartCol -= sf.StartCol;
                    //}
                }
            }
            foreach (var ix in delSF)
            {
                _sharedFormulas.Remove(ix);
            }
            delSF = null;
            var cse = new CellsStoreEnumerator<object>(_formulas, 1, 1,  ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            while (cse.Next())
            {
                if (cse.Value is string)
                {
                    cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), 0, -columns, 0, columnFrom, this.Name, this.Name);
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
            var v = GetValueInner(Row, Column);
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
            var v = GetValueInner(Row, Column);           
            if (v==null)
            {
                return default(T);
            }

            //var cell=((ExcelCell)_cells[cellID]);
            if (_flags.GetFlagValue(Row, Column, CellFlags.RichText))
            {
                return (T)(object)Cells[Row, Column].RichText.Text;
            }

            return ConvertUtil.GetTypedCellValue<T>(v);
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
            SetValueInner(Row, Column, Value);
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
            SetValueInner(row, col, Value);           
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
               if(!string.IsNullOrEmpty( _mergedCells[i]))
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
            }
            return 0;
        }

#endregion
#endregion // END Worksheet Public Methods

#region Worksheet Private Methods
        private void UpdateCrossSheetReferences(string sheetWhoseReferencesShouldBeUpdated, int rowFrom, int rows, int columnFrom, int columns)
        {
          lock (this)
          {
            foreach (var f in _sharedFormulas.Values)
            {
              f.Formula = ExcelCellBase.UpdateFormulaReferences(f.Formula, rows, columns, rowFrom, columnFrom, this.Name, sheetWhoseReferencesShouldBeUpdated);
            }
            var cse = new CellsStoreEnumerator<object>(_formulas);
            while (cse.Next())
            {
              if (cse.Value is string)
              {
                cse.Value = ExcelCellBase.UpdateFormulaReferences(cse.Value.ToString(), rows, columns, rowFrom, columnFrom, this.Name, sheetWhoseReferencesShouldBeUpdated);
              }
            }
          }
        }

        private void UpdateCrossSheetReferenceNames(string oldName, string newName)
        {
          if (string.IsNullOrEmpty(oldName))
            throw new ArgumentNullException(nameof(oldName));
          if (string.IsNullOrEmpty(newName))
            throw new ArgumentNullException(nameof(newName));
          lock (this)
          {
            foreach (var f in _sharedFormulas.Values)
            {
              f.Formula = ExcelCellBase.UpdateFormulaSheetReferences(f.Formula, oldName, newName);
            }
            var cse = new CellsStoreEnumerator<object>(_formulas);
            while (cse.Next())
            {
              if (cse.Value is string)
              {
                cse.Value = ExcelCellBase.UpdateFormulaSheetReferences(cse.Value.ToString(), oldName, newName);
              }
            }
          }
        }
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


                        if (Drawings.Count == 0)
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
                        foreach (ExcelDrawing d in Drawings)
                        {
                            d.AdjustPositionAndSize();
                            if (d is ExcelChart)
                            {
                                ExcelChart c = (ExcelChart)d;
                                c.ChartXml.Save(c.Part.GetStream(FileMode.Create, FileAccess.Write));
                            }
                        }
                        Packaging.ZipPackagePart partPack = Drawings.Part;
                        Drawings.DrawingXml.Save(partPack.GetStream(FileMode.Create, FileAccess.Write));
                }
            }
        }
        internal void SaveHandler(ZipOutputStream stream, CompressionLevel compressionLevel, string fileName)
        {
                    //Init Zip
                    stream.CodecBufferSize = 8096;
                    stream.CompressionLevel = (OfficeOpenXml.Packaging.Ionic.Zlib.CompressionLevel)compressionLevel;
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
                    _comments.CommentXml.Save(_comments.Part.GetStream(FileMode.Create));
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
                    _vmlDrawings.VmlDrawingXml.Save(_vmlDrawings.Part.GetStream(FileMode.Create));
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
                        string n=col.Name.ToLowerInvariant();
                        if (tbl.ShowHeader)
                        {
                            n = tbl.WorkSheet.GetValue<string>(tbl.Address._fromRow,
                                tbl.Address._fromCol + col.Position);
                            if (string.IsNullOrEmpty(n))
                            {
                                n = col.Name.ToLowerInvariant();
                                SetValueInner(tbl.Address._fromRow, colNum, ConvertUtil.ExcelDecodeString(col.Name));
                            }
                            else
                            {
                                col.Name = n;
                            }
                        }
                        else
                        {
                            n = col.Name.ToLowerInvariant();
                        }
                    
                        if(colVal.Contains(n))
                        {
                            throw(new InvalidDataException(string.Format("Table {0} Column {1} does not have a unique name.", tbl.Name, col.Name)));
                        }                        
                        colVal.Add(n);
                        if (!string.IsNullOrEmpty(col.CalculatedColumnFormula))
                        {
                            int fromRow = tbl.ShowHeader ? tbl.Address._fromRow + 1 : tbl.Address._fromRow;
                            int toRow = tbl.ShowTotal ? tbl.Address._toRow - 1 : tbl.Address._toRow;
                            string r1c1Formula = ExcelCellBase.TranslateToR1C1(col.CalculatedColumnFormula, fromRow, colNum);
                            bool needsTranslation = r1c1Formula != col.CalculatedColumnFormula;

                            for (int row = fromRow; row <= toRow; row++)
                            {
                                SetFormula(row, colNum, needsTranslation ? ExcelCellBase.TranslateFromR1C1(r1c1Formula, row, colNum) : r1c1Formula);
                            }
                        }
                        colNum++;
                    }
                }                
                if (tbl.Part == null)
                {
                    var id = tbl.Id;
                    tbl.TableUri = GetNewUri(_package.Package, @"/xl/tables/table{0}.xml", ref id);
                    tbl.Id = id;
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
            if (tbl.ShowTotal == false) return;
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
                    case RowFunctions.CountNums:
                        SetFormula(tbl.Address._toRow, colNum, GetTotalFunction(col, "102"));
                        break;
                    case RowFunctions.Count:
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
                SetValueInner(tbl.Address._toRow, colNum, col.TotalsRowLabel);

            }
        }

        internal void SetFormula(int row, int col, object value)
        {
            _formulas.SetValue(row, col, value);
            if (!ExistsValueInner(row, col)) SetValueInner(row, col, null);
        }
        //internal void SetStyle(int row, int col, int value)
        //{
        //    SetStyleInner(row, col, value);
        //    if(!_values.Exists(row,col)) SetValueInner(row, col, null);
        //}
        
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

                //Rewrite the pivottable address again if any rows or columns have been inserted or deleted
                pt.SetXmlNodeString("d:location/@ref", pt.Address.Address);
                var ws = Workbook.Worksheets[pt.CacheDefinition.SourceRange.WorkSheet];
                var t = ws.Tables.GetFromRange(pt.CacheDefinition.SourceRange);
                if (pt.CacheDefinition.SourceRange!=null && !pt.CacheDefinition.SourceRange.IsName && t==null)
                {
                    pt.CacheDefinition.SetXmlNodeString(ExcelPivotCacheDefinition._sourceAddressPath, pt.CacheDefinition.SourceRange.Address);
                }

                var fields =
                    pt.CacheDefinition.CacheDefinitionXml.SelectNodes(
                        "d:pivotCacheDefinition/d:cacheFields/d:cacheField", NameSpaceManager);
                int ix = 0;
                if (fields != null)
                {
                    var flds = new HashSet<string>();
                    foreach (XmlElement node in fields)
                    {
                        if (ix >= pt.CacheDefinition.SourceRange.Columns) break;
                        var fldName = node.GetAttribute("name");                        //Fixes issue 15295 dup name error
                        if (string.IsNullOrEmpty(fldName))
                        {
                            fldName = (t == null
                                ? pt.CacheDefinition.SourceRange.Offset(0, ix++, 1, 1).Value.ToString()
                                : t.Columns[ix++].Name);
                        }
                        if (flds.Contains(fldName))
                        {                            
                            fldName = GetNewName(flds, fldName);
                        }
                        flds.Add(fldName);
                        node.SetAttribute("name", fldName);
                    }
                    foreach (var df in pt.DataFields)
                    {
                        if (string.IsNullOrEmpty(df.Name))
                        {
                            string name;
                            if (df.Function == DataFieldFunctions.None)
                            {
                                name = df.Field.Name; //Name must be set or Excel will crash on rename.                                
                            }
                            else
                            {
                                name = df.Function.ToString() + " of " + df.Field.Name; //Name must be set or Excel will crash on rename.
                            }
                            //Make sure name is unique
                            var newName = name;
                            var i = 2;
                            while (pt.DataFields.ExistsDfName(newName, df))
                            {
                                newName = name + (i++).ToString(CultureInfo.InvariantCulture);
                            }
                            df.Name = newName;
                        }
                    }
                }
                pt.PivotTableXml.Save(pt.Part.GetStream(FileMode.Create));
                pt.CacheDefinition.CacheDefinitionXml.Save(pt.CacheDefinition.Part.GetStream(FileMode.Create));
            }
        }

        private string GetNewName(HashSet<string> flds, string fldName)
        {
            int ix = 2;
            while (flds.Contains(fldName + ix.ToString(CultureInfo.InvariantCulture)))
            {
                ix++;
            }
            return fldName + ix.ToString(CultureInfo.InvariantCulture);
        }

        private static string GetTotalFunction(ExcelTableColumn col, string FunctionNum)
        {
            string escapedColumn = Regex.Replace(col.Name, @"[\[\]#']", new MatchEvaluator(m => "'" + m.Value));
            return string.Format("SUBTOTAL({0},{1}[{2}])", FunctionNum, col._tbl.Name, escapedColumn);
        }

        private void SaveXml(Stream stream)
        {
            //Create the nodes if they do not exist.
            StreamWriter sw = new StreamWriter(stream, System.Text.Encoding.UTF8, 65536);
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

                CleanupMergedCells(_mergedCells);
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

        private void CleanupMergedCells(MergeCellsCollection _mergedCells)
        {
            int i=0;
            while (i < _mergedCells.List.Count)
            {
                if (_mergedCells[i] == null)
                {
                    _mergedCells.List.RemoveAt(i);
                }
                else
                {
                    i++;
                }
            }
        }
        private void UpdateColBreaks(StreamWriter sw)
        {
            StringBuilder breaks = new StringBuilder();
            int count = 0;
            var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 0, 0, ExcelPackage.MaxColumns);
            //foreach (ExcelColumn col in _columns)
            while(cse.Next())
            {
                var col=cse.Value._value as ExcelColumn;
                if (col != null && col.PageBreak)
                {
                    breaks.AppendFormat("<brk id=\"{0}\" max=\"16383\" man=\"1\"/>", cse.Column);
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
            var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 0, ExcelPackage.MaxRows, 0);
            //foreach(ExcelRow row in _rows)            
            while(cse.Next())
            {
                var row=cse.Value._value as RowInternal;
                if (row != null && row.PageBreak)
                {
                    breaks.AppendFormat("<brk id=\"{0}\" max=\"1048575\" man=\"1\"/>", cse.Row);
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
            var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 1, 0, ExcelPackage.MaxColumns);
            bool first = true;
            while(cse.Next())
            {
                if (first)
                {
                    sw.Write("<cols>");
                    first = false;
                }
                var col = cse.Value._value as ExcelColumn;
                ExcelStyleCollection<ExcelXfs> cellXfs = _package.Workbook.Styles.CellXfs;

                sw.Write("<col min=\"{0}\" max=\"{1}\"", col.ColumnMin, col.ColumnMax);
                if (col.Hidden == true)
                {
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
                sw.Write("/>");
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
            
            int row = -1;

            StringBuilder sbXml = new StringBuilder();
            var ss = _package.Workbook._sharedStrings;
            var styles = _package.Workbook.Styles;
            var cache = new StringBuilder();
            cache.Append("<sheetData>");
            
            ////Set a value for cells with style and no value set.
            //var cseStyle = new CellsStoreEnumerator<ExcelCoreValue>(_values, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            //foreach (var s in cseStyle)
            //{
            //    if(!ExistsValueInner(cseStyle.Row, cseStyle.Column))
            //    {
            //        SetValueInner(cseStyle.Row, cseStyle.Column, null);
            //    }
            //}

            columnStyles = new Dictionary<int, int>();
            var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values, 1, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
            //foreach (IRangeID r in _cells)
            while(cse.Next())
            {
                if (cse.Column>0)
                {
                    var val = cse.Value;
                    //int styleID = cellXfs[styles.GetStyleId(this, cse.Row, cse.Column)].newID;
                    int styleID = cellXfs[(val._styleId == 0 ? GetStyleIdDefaultWithMemo(cse.Row, cse.Column) : val._styleId)].newID;
                    //Add the row element if it's a new row
                    if (cse.Row != row)
                    {
                        WriteRow(cache, cellXfs, row, cse.Row);
                        row = cse.Row;
                    }
                    object v = val._value;
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
                                    cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, f.Address, ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v,true));
                                }
                                else
                                {
                                    cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{6}><f ref=\"{2}\" t=\"shared\" si=\"{3}\">{4}</f>{5}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, f.Address, sfId, ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v,true));
                                }

                            }
                            else if (f.IsArray)
                            {
                                cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cse.CellAddress, styleID < 0 ? 0 : styleID);
                            }
                            else
                            {
                                cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{4}><f t=\"shared\" si=\"{2}\"/>{3}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, sfId, GetFormulaValue(v), GetCellType(v,true));
                            }
                        }
                        else
                        {
                            // We can also have a single cell array formula
                            if(f.IsArray)
                            {
                                cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{5}><f ref=\"{2}\" t=\"array\">{3}</f>{4}</c>", cse.CellAddress, styleID < 0 ? 0 : styleID, string.Format("{0}:{1}", f.Address, f.Address), ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v), GetCellType(v,true));
                            }
                            else
                            {
                                cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", f.Address, styleID < 0 ? 0 : styleID, GetCellType(v, true));
                                cache.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(f.Formula), GetFormulaValue(v));
                            }
                        }
                    }
                    else if (formula!=null && formula.ToString()!="")
                    {
                        cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(v,true));
                        cache.AppendFormat("<f>{0}</f>{1}</c>", ConvertUtil.ExcelEscapeString(formula.ToString()), GetFormulaValue(v));
                    }
                    else
                    {
                        if (v == null && styleID > 0)
                        {
                            cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"/>", cse.CellAddress, styleID < 0 ? 0 : styleID);
                        }
                        else if(v != null)
                        {
                            // Fix for issue 15460
                            var enumerableResult = v as System.Collections.IEnumerable;
                            if (enumerableResult != null && !(v is string))
                            {
                              var enumerator = enumerableResult.GetEnumerator();
                              if (enumerator.MoveNext() && enumerator.Current != null)
                                v = enumerator.Current;
                              else
                                v = string.Empty;
                            }
                            if ((TypeCompat.IsPrimitive(v) || v is double || v is decimal || v is DateTime || v is TimeSpan))
                            {
                                //string sv = GetValueForXml(v);
                                cache.AppendFormat("<c r=\"{0}\" s=\"{1}\"{2}>", cse.CellAddress, styleID < 0 ? 0 : styleID, GetCellType(v));
                                cache.AppendFormat("{0}</c>", GetFormulaValue(v));
                            }
                            else
                            {
                                var vString = Convert.ToString(v);
                                int ix;
                                if (!ss.ContainsKey(vString))
                                {
                                    ix = ss.Count;
                                    ss.Add(vString, new ExcelWorkbook.SharedStringItem() { isRichText = _flags.GetFlagValue(cse.Row,cse.Column,CellFlags.RichText), pos = ix });
                                }
                                else
                                {
                                    ix = ss[vString].pos;
                                }
                                cache.AppendFormat("<c r=\"{0}\" s=\"{1}\" t=\"s\">", cse.CellAddress, styleID < 0 ? 0 : styleID);
                                cache.AppendFormat("<v>{0}</v></c>", ix);
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
                    WriteRow(cache, cellXfs, row, cse.Row);
                    row = cse.Row;
                }
                if (cache.Length > 0x600000)
                {
                    sw.Write(cache.ToString());
                    sw.Flush();
                    cache.Length = 0;
                }
            }
            columnStyles = null;

            if (row != -1) cache.Append("</row>");
            cache.Append("</sheetData>");
            sw.Write(cache.ToString());
            sw.Flush();
        }
        private Dictionary<int, int> columnStyles = null;
        // get StyleID without cell style for UpdateRowCellData
        internal int GetStyleIdDefaultWithMemo(int row, int col)
        {
            int v = 0;
            if (ExistsStyleInner(row, 0, ref v)) //First Row
            {
                return v;
            }
            else // then column
            {
                if (!columnStyles.ContainsKey(col))
                {
                    if (ExistsStyleInner(0, col, ref v))
                    {
                        columnStyles.Add(col, v);
                    }
                    else
                    {
                        int r = 0, c = col;
                        if (_values.PrevCell(ref r, ref c))
                        {
                            //var column=ws.GetValueInner(0,c) as ExcelColumn;
                            var val = _values.GetValue(0, c);
                            var column = (ExcelColumn)(val._value);
                            if (column != null && column.ColumnMax >= col) //Fixes issue 15174
                            {
                                //return ws.GetStyleInner(0, c);
                                columnStyles.Add(col, val._styleId);
                            }
                            else
                            {
                                columnStyles.Add(col, 0);
                            }
                        }
                        else
                        {
                            columnStyles.Add(col, 0);
                        }
                    }
                }
                return columnStyles[col];
            }
        }

        private object GetFormulaValue(object v)
        {
            //if (_package.Workbook._isCalculated)
            //{                    
            if (v != null && v.ToString()!="")
            {
                return "<v>" + ConvertUtil.ExcelEscapeString(GetValueForXml(v)) + "</v>"; //Fixes issue 15071
            }            
            else
            {
                return "";
            }
        }

        private string GetCellType(object v, bool allowStr=false)
        {
            if (v is bool)
            {
                return " t=\"b\"";
            }
            else if ((v is double && double.IsInfinity((double)v)) || v is ExcelErrorValue)
            {
                return " t=\"e\"";
            }
            else if(allowStr && v!=null && !(TypeCompat.IsPrimitive(v) || v is double || v is decimal || v is DateTime || v is TimeSpan))
            {
                return " t=\"str\"";
            }
            else
            {
                return "";
            }
        }

        private string GetValueForXml(object v)
        {
            string s;
            try
            {
                if (v is DateTime)
                {
                    double sdv = ((DateTime)v).ToOADate();

                    if (Workbook.Date1904)
                    {
                        sdv -= ExcelWorkbook.date1904Offset;
                    }

                    s = sdv.ToString(CultureInfo.InvariantCulture);
                }
                else if (v is TimeSpan)
                {
                    s = DateTime.FromOADate(0).Add(((TimeSpan)v)).ToOADate().ToString(CultureInfo.InvariantCulture);
                }
                else if(TypeCompat.IsPrimitive(v) || v is double || v is decimal)
                {
                    if (v is double && double.IsNaN((double)v))
                    {
                        s = "";
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
        private void WriteRow(StringBuilder cache, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
        {
            if (prevRow != -1) cache.Append("</row>");
            //ulong rowID = ExcelRow.GetRowID(SheetID, row);
            cache.AppendFormat("<row r=\"{0}\"", row);
            RowInternal currRow = GetValueInner(row, 0) as RowInternal;
            if (currRow != null)
            {

                // if hidden, add hidden attribute and preserve ht/customHeight (Excel compatible)
                if (currRow.Hidden == true)
                {
                    cache.Append(" hidden=\"1\"");
                }
                if (currRow.Height >= 0)
                {
                    cache.AppendFormat(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", currRow.Height));
                    if (currRow.CustomHeight)
                    {
                        cache.Append(" customHeight=\"1\"");
                    }
                }

                if (currRow.OutlineLevel > 0)
                {
                    cache.AppendFormat(" outlineLevel =\"{0}\"", currRow.OutlineLevel);
                    if (currRow.Collapsed)
                    {
                        if (currRow.Hidden)
                        {
                            cache.Append(" collapsed=\"1\"");
                        }
                        else
                        {
                            cache.Append(" collapsed=\"1\" hidden=\"1\""); //Always hidden
                        }
                    }
                }
                if (currRow.Phonetic)
                {
                    cache.Append(" ph=\"1\"");
                }
            }
            var s = GetStyleInner(row, 0);
            if (s > 0)
            {
                cache.AppendFormat(" s=\"{0}\" customFormat=\"1\"", cellXfs[s].newID);
            }
            cache.Append(">");
        }
        private void WriteRow(StreamWriter sw, ExcelStyleCollection<ExcelXfs> cellXfs, int prevRow, int row)
        {
            if (prevRow != -1) sw.Write("</row>");
            //ulong rowID = ExcelRow.GetRowID(SheetID, row);
            sw.Write("<row r=\"{0}\"", row);
            RowInternal currRow = GetValueInner(row, 0) as RowInternal;
            if (currRow!=null)
            {

                // if hidden, add hidden attribute and preserve ht/customHeight (Excel compatible)
                if (currRow.Hidden == true)
                {
                    sw.Write(" hidden=\"1\"");
                }
                if (currRow.Height >= 0)
                {
                    sw.Write(string.Format(CultureInfo.InvariantCulture, " ht=\"{0}\"", currRow.Height));
                    if (currRow.CustomHeight)
                    {
                        sw.Write(" customHeight=\"1\"");
                    }
                }

                if (currRow.OutlineLevel > 0)
                {
                    sw.Write(" outlineLevel =\"{0}\"", currRow.OutlineLevel);
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
                    sw.Write(" ph=\"1\"");
                }
            }
            var s = GetStyleInner(row, 0);
            if (s > 0)
            {
                sw.Write(" s=\"{0}\" customFormat=\"1\"", cellXfs[s].newID);
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
                    var uri = _hyperLinks.GetValue(cse.Row, cse.Column);
                    if (first && uri != null)
                    {
                        sw.Write("<hyperlinks>");
                        first = false;
                    }

                    //int row, col;
                    //ExcelCell cell = _cells[cellId] as ExcelCell;
                    if (uri is ExcelHyperLink && !string.IsNullOrEmpty((uri as ExcelHyperLink).ReferenceAddress))
                    {
                        ExcelHyperLink hl = uri as ExcelHyperLink;
                        sw.Write("<hyperlink ref=\"{0}\" location=\"{1}\"{2}{3}/>",
                                Cells[cse.Row, cse.Column, cse.Row + hl.RowSpann, cse.Column + hl.ColSpann].Address, 
                                ExcelCellBase.GetFullAddress(SecurityElement.Escape(Name), SecurityElement.Escape(hl.ReferenceAddress)),
                                    string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"",
                                    string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"");
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
                                sw.Write("<hyperlink ref=\"{0}\"{2}{3} r:id=\"{1}\"/>", ExcelCellBase.GetAddress(cse.Row, cse.Column), relationship.Id,                                
                                    string.IsNullOrEmpty(hl.Display) ? "" : " display=\"" + SecurityElement.Escape(hl.Display) + "\"",
                                    string.IsNullOrEmpty(hl.ToolTip) ? "" : " tooltip=\"" + SecurityElement.Escape(hl.ToolTip) + "\"");
                            }
                            else
                            {
                                sw.Write("<hyperlink ref=\"{0}\" r:id=\"{1}\"/>", ExcelCellBase.GetAddress(cse.Row, cse.Column), relationship.Id);
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
                    var addr = new ExcelAddressBase(fromRow, fromCol, toRow, toCol);
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
                if (Workbook._nextTableID == int.MinValue) Workbook.ReadAllTables();
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
                    if (Workbook._nextPivotTableID == int.MinValue) Workbook.ReadAllTables();
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

        internal void ClearValidations()
        {
            _dataValidation = null;
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

        internal void UpdateCellsWithDate1904Setting()
        {
            var cse = new CellsStoreEnumerator<ExcelCoreValue>(_values);
            var offset = Workbook.Date1904 ? -ExcelWorkbook.date1904Offset : ExcelWorkbook.date1904Offset;
            while(cse.MoveNext())
            {
                if (cse.Value._value is DateTime)
                {
                    try
                    {
                        double sdv = ((DateTime)cse.Value._value).ToOADate();
                        sdv += offset;

                        //cse.Value._value = DateTime.FromOADate(sdv);
                        SetValueInner(cse.Row, cse.Column, DateTime.FromOADate(sdv));
                    }
                    catch
                    {
                    }
                }
            }
        }
        internal string GetFormula(int row, int col)
        {
            var v = _formulas.GetValue(row, col);
            if (v is int)
            {
                return _sharedFormulas[(int)v].GetFormula(row,col, Name);
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

        private void DisposeInternal(IDisposable candidateDisposable)
        {
            if (candidateDisposable != null)
            {
                candidateDisposable.Dispose();
            }
        }


        public void Dispose()
        {
            DisposeInternal(_values);
            DisposeInternal(_formulas);
            DisposeInternal(_flags);
            DisposeInternal(_hyperLinks);
            //DisposeInternal(_styles);
            //DisposeInternal(_types);
            DisposeInternal(_commentsStore);
            DisposeInternal(_formulaTokens);

            _values = null;
            _formulas = null;
            _flags = null;
            _hyperLinks = null;
            //_styles = null;
            //_types = null;
            _commentsStore = null;
            _formulaTokens = null;

            _package = null;
            _pivotTables = null;
            _protection = null;
            if(_sharedFormulas != null) _sharedFormulas.Clear();
            _sharedFormulas = null;
            _sheetView = null;
            _tables = null;
            _vmlDrawings = null;
            _conditionalFormatting = null;
            _dataValidation = null;
            _drawings = null;
        }

        /// <summary>
        /// Get the ExcelColumn for column (span ColumnMin and ColumnMax)
        /// </summary>
        /// <param name="column"></param>
        /// <returns></returns>
        internal ExcelColumn GetColumn(int column)
        {
            var c = GetValueInner(0, column) as ExcelColumn;
            if (c == null)
            {
                int row = 0, col = column;
                if (_values.PrevCell(ref row, ref col))
                {
                    c = GetValueInner(0, col) as ExcelColumn;
                    if (c != null && c.ColumnMax >= column)
                    {
                        return c;
                    }
                    return null;
                }
            }
            return c;

        }

        public bool Equals(ExcelWorksheet x, ExcelWorksheet y)
        {
            return x.Name == y.Name && x.SheetID == y.SheetID && x.WorksheetXml.OuterXml == y.WorksheetXml.OuterXml;
        }

        public int GetHashCode(ExcelWorksheet obj)
        {
            return obj.WorksheetXml.OuterXml.GetHashCode();
        }

        #region Worksheet internal Accessor
        /// <summary>
        /// Get accessor of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>cell value</returns>
        internal object GetValueInner(int row, int col)
        {
            return _values.GetValue(row, col)._value;
        }
        /// <summary>
        /// Get accessor of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>cell styleId</returns>
        internal int GetStyleInner(int row, int col)
        {
            return _values.GetValue(row, col)._styleId;
        }

        /// <summary>
        /// Set accessor of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="value">value</param>
        internal void SetValueInner(int row, int col, object value)
        {
            _values.SetValueSpecial(row, col, _setValueInnerUpdateDelegate, value);
        }
        private static CellStore<ExcelCoreValue>.SetValueDelegate _setValueInnerUpdateDelegate = SetValueInnerUpdate;
        private static void SetValueInnerUpdate(List<ExcelCoreValue> list, int index, object value)
        {
            list[index] = new ExcelCoreValue { _value = value, _styleId = list[index]._styleId };
        }
        /// <summary>
        /// Set accessor of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="styleId">styleId</param>
        internal void SetStyleInner(int row, int col, int styleId)
        {
            _values.SetValueSpecial(row, col, (CellStore<ExcelCoreValue>.SetValueDelegate)SetStyleInnerUpdate, styleId);
        }
        void SetStyleInnerUpdate(List<ExcelCoreValue> list, int index, object styleId)
        {
            list[index] = new ExcelCoreValue { _value = list[index]._value, _styleId = (int)styleId };
        }

        /// <summary>
        /// Bulk(Range) set accessor of sheet value, for value array
        /// </summary>
        /// <param name="fromRow">start row</param>
        /// <param name="fromColumn">start column</param>
        /// <param name="toRow">end row</param>
        /// <param name="toColumn">end column</param>
        /// <param name="values">set values</param>
        internal void SetRangeValueInner(int fromRow, int fromColumn, int toRow, int toColumn, object[,] values)
        {
            var rowBound = values.GetUpperBound(0);
            var colBound = values.GetUpperBound(1);
            _values.SetRangeValueSpecial(fromRow, fromColumn, toRow, toColumn,
                (List<ExcelCoreValue> list, int index, int row, int column, object value) => {
                    object val = null;
                    if (rowBound >= row - fromRow && colBound >= column - fromColumn)
                    {
                        val = ((object[,])values)[row - fromRow, column - fromColumn];
                    }
                    list[index] = new ExcelCoreValue { _value = val, _styleId = list[index]._styleId };
                },
                values);
        }
        private void SetRangeValueUpdate(List<ExcelCoreValue> list, int index, int row, int column, object values)
        {
            list[index] = new ExcelCoreValue { _value = ((object[,])values)[row, column], _styleId = list[index]._styleId };
        }

        /// <summary>
        /// Existance check of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>is exists</returns>
        internal bool ExistsValueInner(int row, int col)
        {
            return (_values.GetValue(row, col)._value != null);
        }
        /// <summary>
        /// Existance check of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>is exists</returns>
        internal bool ExistsStyleInner(int row, int col)
        {
            return (_values.GetValue(row, col)._styleId != 0);
        }
        /// <summary>
        /// Existance check of sheet value
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="value"></param>
        /// <returns>is exists</returns>
        internal bool ExistsValueInner(int row, int col, ref object value)
        {
            value = _values.GetValue(row, col)._value;
            return (value != null);
        }
        /// <summary>
        /// Existance check of sheet styleId
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="styleId"></param>
        /// <returns>is exists</returns>
        internal bool ExistsStyleInner(int row, int col, ref int styleId)
        {
            styleId = _values.GetValue(row, col)._styleId;
            return (styleId != 0);
        }
#endregion
    }  // END class Worksheet
}
