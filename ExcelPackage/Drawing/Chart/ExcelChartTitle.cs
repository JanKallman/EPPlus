using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// The title of a chart
    /// </summary>
    public class ExcelChartTitle : XmlHelper
    {
        internal ExcelChartTitle(XmlNamespaceManager nameSpaceManager, XmlNode node) :
            base(nameSpaceManager, node)
        {
            XmlNode topNode = node.SelectSingleNode("c:title", NameSpaceManager);
            if (topNode == null)
            {
                topNode = node.OwnerDocument.CreateElement("c", "title", ExcelPackage.schemaChart);
                node.InsertBefore(topNode, node.ChildNodes[0]);
                topNode.InnerXml = "<c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:r><a:t /></a:r></a:p></c:rich></c:tx><c:layout /><c:overlay val=\"0\" />";
            }
            TopNode = topNode;
            SchemaNodeOrder = new string[] { "tx", "layout", "overlay" };
        }
        const string titlePath = "c:tx/c:rich/a:p/a:r/a:t";
        /// <summary>
        /// The text
        /// </summary>
        public string Text
        {
            get
            {
                //return GetXmlNode(titlePath);
                return RichText.Text;
            }
            set
            {
                //SetXmlNode(titlePath, value);
                RichText.Text = value;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// A reference to the border properties
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "c:spPr/a:ln");
                }
                return _border;
            }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// A reference to the fill properties
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "c:spPr");
                }
                return _fill;
            }
        }
        //ExcelTextFont _font = null;
        /// <summary>
        /// A reference to the font properties
        /// </summary>
        public ExcelTextFont Font
        {
            get
            {
                //if (_font == null)
                //{
                //    _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:tx/c:rich/a:p/a:r/a:rPr", new string[] { "rPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" });
                //}
                //return _font;
                if (_richText==null || _richText.Count == 0)
                {
                    RichText.Add("");
                }
                return _richText[0];
            }
        }
        string[] paragraphNodeOrder = new string[] { "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" };
        ExcelParagraphCollection _richText = null;
        /// <summary>
        /// Richtext
        /// </summary>
        public ExcelParagraphCollection RichText
        {
            get
            {
                if (_richText == null)
                {
                    _richText = new ExcelParagraphCollection(NameSpaceManager, TopNode, "c:tx/c:rich/a:p", paragraphNodeOrder);
                }
                return _richText;
            }
        }
        /// <summary>
        /// Show without overlaping the chart.
        /// </summary>
        public bool Overlay
        {
            get
            {
                return GetXmlNodeBool("overlay/@val");
            }
            set
            {
                SetXmlNodeBool("overlay/@val", value);
            }
        }
    }
}
