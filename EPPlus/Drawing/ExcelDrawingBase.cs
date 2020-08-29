/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Chart;
using System.Drawing;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Text anchoring
    /// </summary>
    public enum eTextAnchoringType
    {
        Bottom,
        Center,
        Distributed,
        Justify,
        Top
    }
    /// <summary>
    /// Vertical text type
    /// </summary>
    public enum eTextVerticalType
    {
        EastAsianVertical,
        Horizontal,
        MongolianVertical,
        Vertical,
        Vertical270,
        WordArtVertical,
        WordArtVerticalRightToLeft

    }
    /// <summary>
    /// How the drawing will be resized.
    /// </summary>
    public enum eEditAs
    {
        /// <summary>
        /// Specifies that the current start and end positions shall
        /// be maintained with respect to the distances from the
        /// absolute start point of the worksheet.
        /// </summary>
        Absolute,
        /// <summary>
        /// Specifies that the current drawing shall move with its
        ///row and column (i.e. the object is anchored to the
        /// actual from row and column), but that the size shall
        ///remain absolute.
        /// </summary>
        OneCell,
        /// <summary>
        /// Specifies that the current drawing shall move and
        /// resize to maintain its row and column anchors (i.e. the
        /// object is anchored to the actual from and to row and column).
        /// </summary>
        TwoCell
    }
    /// <summary>
    /// Base class for twoanchored drawings. 
    /// Drawings are Charts, shapes and Pictures.
    /// </summary>
    public class ExcelDrawing : XmlHelper, IDisposable 
    {
        /// <summary>
        /// Position of the a drawing.
        /// </summary>
        public class ExcelPosition : XmlHelper
        {
            internal delegate void SetWidthCallback();
            XmlNode _node;
            XmlNamespaceManager _ns;
            SetWidthCallback _setWidthCallback;
            internal ExcelPosition(XmlNamespaceManager ns, XmlNode node, SetWidthCallback setWidthCallback) :
                base (ns,node)
            {
                _node = node;
                _ns = ns;
                _setWidthCallback = setWidthCallback;
            }
            const string colPath="xdr:col";
            public int Column
            {
                get
                {
                    return GetXmlNodeInt(colPath);
                }
                set
                {
                    SetXmlNodeString(colPath, value.ToString());
                    _setWidthCallback();
                }
            }
            const string rowPath="xdr:row";
            public int Row
            {
                get
                {
                    return GetXmlNodeInt(rowPath);
                }
                set
                {
                    SetXmlNodeString(rowPath, value.ToString());
                    _setWidthCallback();
                }
            }
            const string colOffPath = "xdr:colOff";
            /// <summary>
            /// Column Offset
            /// 
            /// EMU units   1cm         =   1/360000 
            ///             1US inch    =   1/914400
            ///             1pixel      =   1/9525
            /// </summary>
            public int ColumnOff
            {
                get
                {
                    return GetXmlNodeInt(colOffPath);
                }
                set
                {
                    SetXmlNodeString(colOffPath, value.ToString());
                    _setWidthCallback();
                }
            }
            const string rowOffPath = "xdr:rowOff";
            /// <summary>
            /// Row Offset
            /// 
            /// EMU units   1cm         =   1/360000 
            ///             1US inch    =   1/914400
            ///             1pixel      =   1/9525
            /// </summary>
            public int RowOff
            {
                get
                {
                    return GetXmlNodeInt(rowOffPath);
                }
                set
                {
                    SetXmlNodeString(rowOffPath, value.ToString());
                    _setWidthCallback();
                }
            }
        }
        protected ExcelDrawings _drawings;
        protected XmlNode _topNode;
        string _nameXPath;
        protected internal int _id;
        const float STANDARD_DPI = 96;
        public const int EMU_PER_PIXEL = 9525;
        protected internal int _width = int.MinValue, _height = int.MinValue, _top=int.MinValue, _left=int.MinValue;
        bool _doNotAdjust = false;
        internal ExcelDrawing(ExcelDrawings drawings, XmlNode node, string nameXPath) :
            base(drawings.NameSpaceManager, node)
        {
            _drawings = drawings;
            _topNode = node;
            _id = drawings.Worksheet.Workbook._nextDrawingID++;
            XmlNode posNode = node.SelectSingleNode("xdr:from", drawings.NameSpaceManager);
            if (node != null)
            {
                From = new ExcelPosition(drawings.NameSpaceManager, posNode, GetPositionSize);
            }
            posNode = node.SelectSingleNode("xdr:to", drawings.NameSpaceManager);
            if (node != null)
            {
                To = new ExcelPosition(drawings.NameSpaceManager, posNode, GetPositionSize);
            }
            else
            {
                To = null;
            }
            GetPositionSize();
            _nameXPath = nameXPath;
            SchemaNodeOrder = new string[] { "from", "to", "graphicFrame", "sp", "clientData"  };
        }
        /// <summary>
        /// The name of the drawing object
        /// </summary>
        public string Name 
        {
            get
            {
                try
                {
                    if (_nameXPath == "") return "";
                    return GetXmlNodeString(_nameXPath);
                }
                catch
                {
                    return ""; 
                }
            }
            set
            {
                try
                {
                    if (_nameXPath == "") throw new NotImplementedException();
                    SetXmlNodeString(_nameXPath, value);
                }
                catch
                {
                    throw new NotImplementedException();
                }
            }
        }
        /// <summary>
        /// How Excel resize drawings when the column width is changed within Excel.
        /// The width of drawings are currently NOT resized in EPPLus when the column width changes
        /// </summary>
        public eEditAs EditAs
        {
            get
            {
                try
                {
                    string s = GetXmlNodeString("@editAs");
                    if (s == "")
                    {
                        return eEditAs.TwoCell;
                    }
                    else
                    {
                        return (eEditAs)Enum.Parse(typeof(eEditAs), s,true);
                    }
                }
                catch
                {
                    return eEditAs.TwoCell;
                }
            }
            set
            {
                string s=value.ToString();
                SetXmlNodeString("@editAs", s.Substring(0,1).ToLower(CultureInfo.InvariantCulture)+s.Substring(1,s.Length-1));
            }
        }
        const string lockedPath="xdr:clientData/@fLocksWithSheet";
        /// <summary>
        /// Lock drawing
        /// </summary>
        public bool Locked
        {
            get
            {
                return GetXmlNodeBool(lockedPath, true);
            }
            set
            {
                SetXmlNodeBool(lockedPath, value);
            }
        }
        const string printPath = "xdr:clientData/@fPrintsWithSheet";
        /// <summary>
        /// Print drawing with sheet
        /// </summary>
        public bool Print
        {
            get
            {
                return GetXmlNodeBool(printPath, true);
            }
            set
            {
                SetXmlNodeBool(printPath, value);
            }
        }        /// <summary>
        /// Top Left position
        /// </summary>
        public ExcelPosition From
        {
            get;
            private set;
        }
        /// <summary>
        /// Bottom right position
        /// </summary>
        ExcelPosition _to = null;
        public ExcelPosition To
        {
            get
            {
                return _to;
            }
            private set
            {
                _to = value;
            }
        }
        internal void ReSetWidthHeight()
        {
            if (!_drawings.Worksheet.Workbook._package.DoAdjustDrawings &&
                EditAs==eEditAs.Absolute)
            {
                SetPixelWidth(_width);
                SetPixelHeight(_height);
            }
        }
        /// <summary>
        /// Add new Drawing types here
        /// </summary>
        /// <param name="drawings">The drawing collection</param>
        /// <param name="node">Xml top node</param>
        /// <returns>The Drawing object</returns>
        internal static ExcelDrawing GetDrawing(ExcelDrawings drawings, XmlNode node)
        {
            if (node.SelectSingleNode("xdr:sp", drawings.NameSpaceManager) != null)
            {
                return new ExcelShape(drawings, node);
            }
            //}
            else if (node.SelectSingleNode("xdr:pic", drawings.NameSpaceManager) != null)
            {
                return new ExcelPicture(drawings, node);
            }
            else if (node.SelectSingleNode("xdr:graphicFrame", drawings.NameSpaceManager) != null)
            {
                return ExcelChart.GetChart(drawings, node);
            }
            else
            {
                return new ExcelDrawing(drawings, node, "");
            }
        }
        internal string Id
        {
            get { return _id.ToString(); }
        }
        internal static string GetTextAchoringText(eTextAnchoringType value)
        {
            switch (value)
            {
                case eTextAnchoringType.Bottom:
                    return "b";
                case eTextAnchoringType.Center:
                    return "ctr";
                case eTextAnchoringType.Distributed:
                    return "dist";
                case eTextAnchoringType.Justify:
                    return "just";
                default:
                    return "t";
            }
        }
        internal static eTextAnchoringType GetTextAchoringEnum(string text)
        {
            switch (text)
            {
                case "b":
                    return eTextAnchoringType.Bottom;
                case "ctr":
                    return eTextAnchoringType.Center;
                case "dist":
                    return eTextAnchoringType.Distributed;
                case "just":
                    return eTextAnchoringType.Justify;
                default:
                    return eTextAnchoringType.Top;
            }
        }
        internal static string GetTextVerticalText(eTextVerticalType value)
        {
            switch (value)
            {
                case eTextVerticalType.EastAsianVertical:
                    return "eaVert";
                case eTextVerticalType.MongolianVertical:
                    return "mongolianVert";
                case eTextVerticalType.Vertical:
                    return "vert";
                case eTextVerticalType.Vertical270:
                    return "vert270";
                case eTextVerticalType.WordArtVertical:
                    return "wordArtVert";
                case eTextVerticalType.WordArtVerticalRightToLeft:
                    return "wordArtVertRtl";
                default:
                    return "horz";
            }
        }
        internal static eTextVerticalType GetTextVerticalEnum(string text)
        {
            switch (text)
            {
                case "eaVert":
                    return eTextVerticalType.EastAsianVertical;
                case "mongolianVert":
                    return eTextVerticalType.MongolianVertical;
                case "vert":
                    return eTextVerticalType.Vertical;
                case "vert270":
                    return eTextVerticalType.Vertical270;
                case "wordArtVert":
                    return eTextVerticalType.WordArtVertical;
                case "wordArtVertRtl":
                    return eTextVerticalType.WordArtVerticalRightToLeft;
                default:
                    return eTextVerticalType.Horizontal;
            }
        }
        #region "Internal sizing functions"
        internal int GetPixelLeft()
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;

            int pix = 0;
            for (int col = 0; col < From.Column; col++)
            {
                pix += (int)decimal.Truncate(((256 * GetColumnWidth(col + 1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            }
            pix += From.ColumnOff / EMU_PER_PIXEL;
            return pix;
        }
        internal int GetPixelTop()
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            int pix = 0;
            for (int row = 0; row < From.Row; row++)
            {
                pix += (int)(GetRowHeight(row + 1) / 0.75);
            }
            pix += From.RowOff / EMU_PER_PIXEL;
            return pix;
        }
        internal int GetPixelWidth()
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;

            int pix = -From.ColumnOff / EMU_PER_PIXEL;
            for (int col = From.Column + 1; col <= To.Column; col++)
            {
                pix += (int)decimal.Truncate(((256 * GetColumnWidth(col) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            }
            pix += Convert.ToInt32(Math.Round(Convert.ToDouble(To.ColumnOff) / EMU_PER_PIXEL,0));
            return pix;
        }
        internal int GetPixelHeight()
        {
            ExcelWorksheet ws = _drawings.Worksheet;

            int pix = -(From.RowOff / EMU_PER_PIXEL);
            for (int row = From.Row + 1; row <= To.Row; row++)
            {
                pix += (int)(GetRowHeight(row) / 0.75);
            }
            pix += Convert.ToInt32(Math.Round(Convert.ToDouble(To.RowOff) / EMU_PER_PIXEL, 0));
            return pix;
        }

        private decimal GetColumnWidth(int col)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            var column = ws.GetValueInner(0, col) as ExcelColumn;
            if (column == null)   //Check that the column exists
            {
                return (decimal)ws.DefaultColWidth;
            }
            else
            {
                return (decimal)ws.Column(col).VisualWidth;
            }
        }
        private double GetRowHeight(int row)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            object o = null;
            if (ws.ExistsValueInner(row, 0, ref o) && o != null)   //Check that the row exists
            {
                var internalRow = (RowInternal)o;
                if (internalRow.Height >= 0 && internalRow.CustomHeight)
                {
                    return internalRow.Height;
                }
                else
                {
                    return GetRowHeightFromCellFonts(row, ws);
                }
            }
            else
            {
                //The row exists, check largest font in row

                /**** Default row height is assumed here. Excel calcualtes the row height from the larges font on the line. The formula to this calculation is undocumented, so currently its implemented with constants... ****/
                return GetRowHeightFromCellFonts(row, ws);
            }
        }

        private double GetRowHeightFromCellFonts(int row, ExcelWorksheet ws)
        {
            var dh = ws.DefaultRowHeight;
            if (double.IsNaN(dh) || ws.CustomHeight==false)
            {
                var height = dh;

                var cse = new CellsStoreEnumerator<ExcelCoreValue>(_drawings.Worksheet._values, row, 0, row, ExcelPackage.MaxColumns);
                var styles = _drawings.Worksheet.Workbook.Styles;
                while (cse.Next())
                {
                    var xfs = styles.CellXfs[cse.Value._styleId];
                    var f = styles.Fonts[xfs.FontId];
                    var rh = ExcelFontXml.GetFontHeight(f.Name, f.Size) * 0.75;
                    if (rh > height)
                    {
                        height = rh;
                    }
                }
                return height;
            }
            else
            {
                return dh;
            }
        }

        internal void SetPixelTop(int pixels)
        {
            _doNotAdjust = true;
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;
            int prevPix = 0;
            int pix = (int)(GetRowHeight(1) / 0.75);
            int row = 2;
            while (pix < pixels)
            {
                prevPix = pix;
                pix += (int)(GetRowHeight(row++) / 0.75);
            }

            if (pix == pixels)
            {
                From.Row = row - 1;
                From.RowOff = 0;
            }
            else
            {
                From.Row = row - 2;
                From.RowOff = (pixels - prevPix) * EMU_PER_PIXEL;
            }
            _top = pixels;
            _doNotAdjust = false;
        }
        internal void SetPixelLeft(int pixels)
        {
            _doNotAdjust = true;
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;
            int prevPix = 0;
            int pix = (int)decimal.Truncate(((256 * GetColumnWidth(1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            int col = 2;

            while (pix < pixels)
            {
                prevPix = pix;
                pix += (int)decimal.Truncate(((256 * GetColumnWidth(col++) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            }
            if (pix == pixels)
            {
                From.Column = col - 1;
                From.ColumnOff = 0;
            }
            else
            {
                From.Column = col - 2;
                From.ColumnOff = (pixels - prevPix) * EMU_PER_PIXEL;
            }
            _left = pixels;
            _doNotAdjust = false;
        }
        internal void SetPixelHeight(int pixels)
        {
            SetPixelHeight(pixels, STANDARD_DPI);
        }
        internal void SetPixelHeight(int pixels, float dpi)
        {
            _doNotAdjust = true;
            ExcelWorksheet ws = _drawings.Worksheet;
            //decimal mdw = ws.Workbook.MaxFontWidth;
            pixels = (int)(pixels / (dpi / STANDARD_DPI) + .5);
            int pixOff = pixels - ((int)(GetRowHeight(From.Row + 1) / 0.75) - (int)(From.RowOff / EMU_PER_PIXEL));
            int prevPixOff = pixels;
            int row = From.Row + 1;

            while (pixOff >= 0)
            {
                prevPixOff = pixOff;
                pixOff -= (int)(GetRowHeight(++row) / 0.75);
            }
            To.Row = row - 1;
            if (From.Row == To.Row)
            {
                To.RowOff = From.RowOff + (pixels) * EMU_PER_PIXEL;
            }
            else
            {
                 To.RowOff = prevPixOff * EMU_PER_PIXEL;
            }
            _doNotAdjust = false;
        }
        internal void SetPixelWidth(int pixels)
        {
            SetPixelWidth(pixels, STANDARD_DPI);
        }
        internal void SetPixelWidth(int pixels, float dpi)
        {
            _doNotAdjust = true;
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;

            pixels = (int)(pixels / (dpi / STANDARD_DPI) + .5);
            int pixOff = (int)pixels - ((int)decimal.Truncate(((256 * GetColumnWidth(From.Column + 1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw) - From.ColumnOff / EMU_PER_PIXEL);
            int prevPixOff = From.ColumnOff / EMU_PER_PIXEL + (int)pixels;
            int col = From.Column + 2;

            while (pixOff >= 0)
            {
                prevPixOff = pixOff;
                pixOff -= (int)decimal.Truncate(((256 * GetColumnWidth(col++) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            }

            To.Column = col - 2;
            To.ColumnOff = prevPixOff * EMU_PER_PIXEL;
            _doNotAdjust = false;
        }
        #endregion
        #region "Public sizing functions"
        /// <summary>
        /// Set the top left corner of a drawing. 
        /// Note that resizing columns / rows after using this function will effect the position of the drawing
        /// </summary>
        /// <param name="PixelTop">Top pixel</param>
        /// <param name="PixelLeft">Left pixel</param>
        public void SetPosition(int PixelTop, int PixelLeft)
        {
            _doNotAdjust = true;
            if (_width == int.MinValue)
            {
                _width = GetPixelWidth();
                _height = GetPixelHeight();
            }

            SetPixelTop(PixelTop);
            SetPixelLeft(PixelLeft);

            SetPixelWidth(_width);
            SetPixelHeight(_height);
            _doNotAdjust = false;
        }
        /// <summary>
        /// Set the top left corner of a drawing. 
        /// Note that resizing columns / rows after using this function will effect the position of the drawing
        /// </summary>
        /// <param name="Row">Start row - 0-based index.</param>
        /// <param name="RowOffsetPixels">Offset in pixels</param>
        /// <param name="Column">Start Column - 0-based index.</param>
        /// <param name="ColumnOffsetPixels">Offset in pixels</param>
        public void SetPosition(int Row, int RowOffsetPixels, int Column, int ColumnOffsetPixels)
        {
            if (RowOffsetPixels < -60)
            {
                throw new ArgumentException("Minimum negative offset is -60.", nameof(RowOffsetPixels));
            }
            if (ColumnOffsetPixels < -60)
            {
                throw new ArgumentException("Minimum negative offset is -60.", nameof(ColumnOffsetPixels));
            }
            _doNotAdjust = true;

            if (_width == int.MinValue)
            {
                _width = GetPixelWidth();
                _height = GetPixelHeight();
            }

            From.Row = Row;
            From.RowOff = RowOffsetPixels * EMU_PER_PIXEL;
            From.Column = Column;
            From.ColumnOff = ColumnOffsetPixels * EMU_PER_PIXEL;

            SetPixelWidth(_width);
            SetPixelHeight(_height);
            _doNotAdjust = false;
        }
        /// <summary>
        /// Set size in Percent
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="Percent"></param>
        public virtual void SetSize(int Percent)
        {
            _doNotAdjust = true;
            if (_width == int.MinValue)
            {
                _width = GetPixelWidth();
                _height = GetPixelHeight();
            }
            _width = (int)(_width * ((decimal)Percent / 100));
            _height = (int)(_height * ((decimal)Percent / 100));

            SetPixelWidth(_width, 96);
            SetPixelHeight(_height, 96);
            _doNotAdjust = false;
        }
        /// <summary>
        /// Set size in pixels
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="PixelWidth">Width in pixels</param>
        /// <param name="PixelHeight">Height in pixels</param>
        public void SetSize(int PixelWidth, int PixelHeight)
        {
            _doNotAdjust = true;
            _width = PixelWidth;
            _height = PixelHeight;
            SetPixelWidth(PixelWidth);
            SetPixelHeight(PixelHeight);
            _doNotAdjust = false;
        }
        #endregion
        internal virtual void DeleteMe()
        {
            TopNode.ParentNode.RemoveChild(TopNode);
        }

        public virtual void Dispose()
        {
            _topNode = null;
        }
        internal void GetPositionSize()
        {
            if (_doNotAdjust) return;
            _top = GetPixelTop();
            _left = GetPixelLeft();
            _height = GetPixelHeight();
            _width = GetPixelWidth();
        }
        /// <summary>
        /// Will adjust the position and size of the drawing according to chages in font of rows and to the Normal style.
        /// This method will be called before save, so use it only if you need the coordinates of the drawing.
        /// </summary>
        public void AdjustPositionAndSize()
        {
            if (_drawings.Worksheet.Workbook._package.DoAdjustDrawings == false) return;
            if (EditAs==eEditAs.Absolute)
            {
                SetPixelLeft(_left);
                SetPixelTop(_top);
            }
            if(EditAs == eEditAs.Absolute || EditAs == eEditAs.OneCell)
            {
                SetPixelHeight(_height);
                SetPixelWidth(_width);
            }
        }
    }
}
