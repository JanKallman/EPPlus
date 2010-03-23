/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 *
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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Drawing.Chart;
namespace OfficeOpenXml.Drawing
{
    public enum eTextAnchoringType
    {
        Bottom,
        Center,
        Distributed,
        Justify,
        Top
    }
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
    public class ExcelDrawing : XmlHelper 
    {
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
                    SetXmlNode(colPath, value.ToString());
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
                    SetXmlNode(rowPath, value.ToString());
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
                    SetXmlNode(colOffPath, value.ToString());
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
                    SetXmlNode(rowOffPath, value.ToString());
                }
            }
        }
        protected ExcelDrawings _drawings;
        protected XmlNode _topNode;
        string _nameXPath;
        int _id;
        const float STANDARD_DPI = 96;
        public const int EMU_PER_PIXEL = 9525;

        //int _top, _left, _height, _width;

        internal ExcelDrawing(XmlNamespaceManager nameSpaceManager, XmlNode node, string nameXPath) :
            base(nameSpaceManager, node)
        {
            _nameXPath = nameXPath;
        }
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
            _nameXPath = nameXPath;
        }
        public ExcelDrawing(XmlNamespaceManager nameSpaceManager, XmlNode node) :
            base(nameSpaceManager, node)
        {
        }
        public string Name 
        {
            get
            {
                try
                {
                    if (_nameXPath == "") return "";
                    return GetXmlNode(_nameXPath);
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
                    SetXmlNode(_nameXPath, value);
                }
                catch
                {
                    throw new NotImplementedException();
                }
            }
        }
        /// <summary>
        /// Top Left position
        /// </summary>
        public ExcelPosition From { get; set; }
        /// <summary>
        /// Bottom right position
        /// </summary>
        public ExcelPosition To { get; set; }
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
                return new ExcelChart(drawings, node);
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
        #region "Internal sizing functions"
        internal int GetPixelWidth()
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            decimal mdw=ws.Workbook.MaxFontWidth;

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
            decimal mdw = ws.Workbook.MaxFontWidth;

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
            if (ws._columns.ContainsKey(ExcelColumn.GetColumnID(ws.SheetID, col)))   //Check that the column exists
            {
                return (decimal)ws.Column(col).Width;
            }
            else
            {
                return (decimal)ws.defaultColWidth;
            }
        }
        private double GetRowWidth(int row)
        {
            ExcelWorksheet ws = _drawings.Worksheet;
            if (ws._rows.ContainsKey(ExcelRow.GetRowID(ws.SheetID, row)))   //Check that the row exists
            {
                return (double)ws.Row(row).Height;
            }
            else
            {
                return (double)ws.defaultRowHeight;
            }
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
            int pix = (int)decimal.Truncate(((256 * (decimal)ws.Column(1).Width + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            int col = 2;

            while (pix < pixels)
            {
                prevPix = pix;
                pix += (int)decimal.Truncate(((256 * (decimal)ws.Column(col++).Width + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
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
            decimal mdw = ws.Workbook.MaxFontWidth;
            pixels = (int)(pixels / (dpi / STANDARD_DPI));
            int pixOff = pixels - ((int)(ws.Row(From.Row + 1).Height / 0.75) - (int)(From.RowOff / EMU_PER_PIXEL));
            int prevPixOff = pixels;
            int row = From.Row+2;

            while (pixOff >= 0)
            {
                prevPixOff = pixOff;
                pixOff -= (int)(ws.Row(++row).Height / 0.75);
            }
            To.Row = row - 2;
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

            pixels = (int)(pixels / (dpi / STANDARD_DPI));
            int pixOff = (int)pixels - ((int)decimal.Truncate(((256 * GetColumnWidth(From.Column + 1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw) - From.ColumnOff / EMU_PER_PIXEL);
            int prevPixOff = From.ColumnOff / EMU_PER_PIXEL + (int)pixels;
            int col = From.Column + 2;

            while (pixOff >= 0)
            {
                prevPixOff = pixOff;
                pixOff -= (int)decimal.Truncate(((256 * GetColumnWidth(++col) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
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
    }
}
