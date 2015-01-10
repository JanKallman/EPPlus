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
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Chart;
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
            XmlNode _node;
            XmlNamespaceManager _ns;            
            internal ExcelPosition(XmlNamespaceManager ns, XmlNode node) :
                base (ns,node)
            {
                _node = node;
                _ns = ns;
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
                }
            }
        }
        protected ExcelDrawings _drawings;
        protected XmlNode _topNode;
        string _nameXPath;
        protected internal int _id;
        const float STANDARD_DPI = 96;
        public const int EMU_PER_PIXEL = 9525;

        internal ExcelDrawing(ExcelDrawings drawings, XmlNode node, string nameXPath) :
            base(drawings.NameSpaceManager, node)
        {
            _drawings = drawings;
            _topNode = node;
            _id = drawings.Worksheet.Workbook._nextDrawingID++;
            XmlNode posNode = node.SelectSingleNode("xdr:from", drawings.NameSpaceManager);
            if (node != null)
            {
                From = new ExcelPosition(drawings.NameSpaceManager, posNode);
            }
            posNode = node.SelectSingleNode("xdr:to", drawings.NameSpaceManager);
            if (node != null)
            {
                To = new ExcelPosition(drawings.NameSpaceManager, posNode);
            }
            else
            {
                To = null;
            }
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
                SetXmlNodeString("@editAs", s.Substring(0,1).ToLower()+s.Substring(1,s.Length-1));
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
        public ExcelPosition From { get; set; }
        /// <summary>
        /// Bottom right position
        /// </summary>
        public ExcelPosition To
        {
            get;
            set;
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
                pix += (int)(GetRowWidth(row + 1) / 0.75);
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
            pix += To.ColumnOff / EMU_PER_PIXEL;
            return pix;
        }
        internal int GetPixelHeight()
        {
            ExcelWorksheet ws = _drawings.Worksheet;

            int pix = -(From.RowOff / EMU_PER_PIXEL);
            for (int row = From.Row + 1; row <= To.Row; row++)
            {
                pix += (int)(GetRowWidth(row) / 0.75);
            }
            pix += To.RowOff / EMU_PER_PIXEL;
            return pix;
        }

        private decimal GetColumnWidth(int col)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            var column = ws._values.GetValue(0, col) as ExcelColumn;
            if (column == null)   //Check that the column exists
            {
                return (decimal)ws.DefaultColWidth;
            }
            else
            {
                return (decimal)ws.Column(col).VisualWidth;
            }
        }
        private double GetRowWidth(int row)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            object o = null;
            if (ws._values.Exists(row, 0, ref o) && o != null)   //Check that the row exists
            {
                var internalRow = (RowInternal)o;
                if (internalRow.Height >= 0)
                {
                    return internalRow.Height;
                }
            }
            return (double)ws.DefaultRowHeight;
        }
        internal void SetPixelTop(int pixels)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw = ws.Workbook.MaxFontWidth;
            int prevPix = 0;
            int pix = (int)(GetRowWidth(1) / 0.75);
            int row = 2;

            while (pix < pixels)
            {
                prevPix = pix;
                pix += (int)(GetRowWidth(row++) / 0.75);
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
        }
        internal void SetPixelLeft(int pixels)
        {
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
        }
        internal void SetPixelHeight(int pixels)
        {
            SetPixelHeight(pixels, STANDARD_DPI);
        }
        internal void SetPixelHeight(int pixels, float dpi)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            //decimal mdw = ws.Workbook.MaxFontWidth;
            pixels = (int)(pixels / (dpi / STANDARD_DPI) + .5);
            int pixOff = pixels - ((int)(ws.Row(From.Row + 1).Height / 0.75) - (int)(From.RowOff / EMU_PER_PIXEL));
            int prevPixOff = pixels;
            int row = From.Row + 1;

            while (pixOff >= 0)
            {
                prevPixOff = pixOff;
                pixOff -= (int)(GetRowWidth(++row) / 0.75);
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
        }
        internal void SetPixelWidth(int pixels)
        {
            SetPixelWidth(pixels, STANDARD_DPI);
        }
        internal void SetPixelWidth(int pixels, float dpi)
        {
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
            int width = GetPixelWidth();
            int height = GetPixelHeight();

            SetPixelTop(PixelTop);
            SetPixelLeft(PixelLeft);

            SetPixelWidth(width);
            SetPixelHeight(height);
        }
        /// <summary>
        /// Set the top left corner of a drawing. 
        /// Note that resizing columns / rows after using this function will effect the position of the drawing
        /// </summary>
        /// <param name="Row">Start row</param>
        /// <param name="RowOffsetPixels">Offset in pixels</param>
        /// <param name="Column">Start Column</param>
        /// <param name="ColumnOffsetPixels">Offset in pixels</param>
        public void SetPosition(int Row, int RowOffsetPixels, int Column, int ColumnOffsetPixels)
        {
            int width = GetPixelWidth();
            int height = GetPixelHeight();

            From.Row = Row;
            From.RowOff = RowOffsetPixels * EMU_PER_PIXEL;
            From.Column = Column;
            From.ColumnOff = ColumnOffsetPixels * EMU_PER_PIXEL;

            SetPixelWidth(width);
            SetPixelHeight(height);
        }
        /// <summary>
        /// Set size in Percent
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="Percent"></param>
        public virtual void SetSize(int Percent)
        {
            int width = GetPixelWidth();
            int height = GetPixelHeight();

            width = (int)(width * ((decimal)Percent / 100));
            height = (int)(height * ((decimal)Percent / 100));

            SetPixelWidth(width, 96);
            SetPixelHeight(height, 96);
        }
        /// <summary>
        /// Set size in pixels
        /// Note that resizing columns / rows after using this function will effect the size of the drawing
        /// </summary>
        /// <param name="PixelWidth">Width in pixels</param>
        /// <param name="PixelHeight">Height in pixels</param>
        public void SetSize(int PixelWidth, int PixelHeight)
        {
            SetPixelWidth(PixelWidth);
            SetPixelHeight(PixelHeight);
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
    }
}
