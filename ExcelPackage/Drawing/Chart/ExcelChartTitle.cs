using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Chart
{
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
                topNode.InnerXml = "<c:tx><c:rich><a:bodyPr /><a:lstStyle /><a:p><a:r><a:t /></a:r></a:p></c:rich></c:tx><c:layout />";
            }
            TopNode = topNode;
        }
        const string titlePath = "c:tx/c:rich/a:p/a:r/a:t";
        public string Text
        {
            get
            {
                return GetXmlNode(titlePath);
            }
            set
            {
                SetXmlNode(titlePath, value);
            }
        }
        ExcelDrawingFill _fill = null;
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
    }
}
