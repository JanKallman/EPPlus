using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// The position of a VML drawing. Used for comments
    /// </summary>
    public class ExcelVmlDrawingPosition : XmlHelper
    {
        int _startPos;
        public ExcelVmlDrawingPosition(XmlNamespaceManager ns, XmlNode topNode, int startPos) : 
            base(ns, topNode)
        {
            _startPos = startPos;
        }
        /// <summary>
        /// Row. Zero based
        /// </summary>
        public int Row
        {
            get
            {
                return GetNumber(2);
            }
            set
            {
                SetNumber(2, value);
            } 
        }
        /// <summary>
        /// Row offset in pixels. Zero based
        /// </summary>
        public int RowOffset
        {
            get
            {
                return GetNumber(3);
            }
            set
            {
                SetNumber(3, value);
            }
        }
        /// <summary>
        /// Column. Zero based
        /// </summary>
        public int Column
        {
            get
            {
                return GetNumber(0);
            }
            set
            {
                SetNumber(0, value);
            }
        }
        /// <summary>
        /// Column offset. Zero based
        /// </summary>
        public int ColumnOffset
        {
            get
            {
                return GetNumber(1);
            }
            set
            {
                SetNumber(1, value);
            }
        }
        private void SetNumber(int pos, int value)
        {
            string anchor = GetXmlNodeString("x:Anchor");
            string[] numbers = anchor.Split(',');
            if (numbers.Length == 8)
            {
                numbers[_startPos + pos] = value.ToString();
            }
            else
            {
                throw (new Exception("Anchor element is invalid in vmlDrawing"));
            }
            SetXmlNodeString("x:Anchor", string.Join(",",numbers));
        }

        private int GetNumber(int pos)
        {
            string anchor = GetXmlNodeString("x:Anchor");
            string[] numbers = anchor.Split(',');
            if (numbers.Length == 8)
            {
                int ret;
                if (int.TryParse(numbers[_startPos + pos], out ret))
                {
                    return ret;
                }
            }
            throw(new Exception("Anchor element is invalid in vmlDrawing"));
        }
    }
}
