using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Style;
using System.Xml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;

namespace OfficeOpenXml
{
    /// <summary>
    /// An Excel Cell Comment
    /// </summary>
    public class ExcelComment : ExcelVmlDrawingComment
    {
        internal XmlHelper _commentHelper;
        internal ExcelComment(XmlNamespaceManager ns, XmlNode commentTopNode, ExcelRangeBase cell)
            : base(null, cell, cell.Worksheet.VmlDrawingsComments.NameSpaceManager)
        {
            //_commentHelper = new XmlHelper(ns, commentTopNode);
            _commentHelper = XmlHelperFactory.Create(ns, commentTopNode);
            var textElem=commentTopNode.SelectSingleNode("d:text", ns);
            if (textElem == null)
            {
                textElem = commentTopNode.OwnerDocument.CreateElement("text", ExcelPackage.schemaMain);
                commentTopNode.AppendChild(textElem);
            }
            if (!cell.Worksheet.VmlDrawingsComments.ContainsKey(ExcelAddress.GetCellID(cell.Worksheet.SheetID, cell.Start.Row, cell.Start.Column)))
            {
                cell.Worksheet.VmlDrawingsComments.Add(cell);
            }

            TopNode = cell.Worksheet.VmlDrawingsComments[ExcelCellBase.GetCellID(cell.Worksheet.SheetID, cell.Start.Row, cell.Start.Column)].TopNode;
            RichText = new ExcelRichTextCollection(ns,textElem);
        }
        const string AUTHORS_PATH = "d:comments/d:authors";
        const string AUTHOR_PATH = "d:comments/d:authors/d:author";
        /// <summary>
        /// Author
        /// </summary>
        public string Author
        {
            get
            {
                int authorRef = _commentHelper.GetXmlNodeInt("@authorId");
                return _commentHelper.TopNode.OwnerDocument.SelectSingleNode(string.Format("{0}[{1}]", AUTHOR_PATH, authorRef+1), _commentHelper.NameSpaceManager).InnerText;
            }
            set
            {
                int authorRef = GetAuthor(value);
                _commentHelper.SetXmlNodeString("@authorId", authorRef.ToString());
            }
        }
        private int GetAuthor(string value)
        {
            int authorRef = 0;
            bool found = false;
            foreach (XmlElement node in _commentHelper.TopNode.OwnerDocument.SelectNodes(AUTHOR_PATH, _commentHelper.NameSpaceManager))
            {
                if (node.InnerText == value)
                {
                    found = true;
                    break;
                }
                authorRef++;
            }
            if (!found)
            {
                var elem = _commentHelper.TopNode.OwnerDocument.CreateElement("d", "author", ExcelPackage.schemaMain);
                _commentHelper.TopNode.OwnerDocument.SelectSingleNode(AUTHORS_PATH, _commentHelper.NameSpaceManager).AppendChild(elem);
                elem.InnerText = value;
            }
            return authorRef;
        }
        /// <summary>
        /// The comment text 
        /// </summary>
        public string Text
        {
            get
            {
                return RichText.Text;
            }
            set
            {
                RichText.Text = value;
            }
        }
        /// <summary>
        /// Sets the font of the first richtext item.
        /// </summary>
        public ExcelRichText Font
        {
            get
            {
                if (RichText.Count > 0)
                {
                    return RichText[0];
                }
                return null;
            }
        }
        /// <summary>
        /// Richtext collection
        /// </summary>
        public ExcelRichTextCollection RichText 
        { 
           get; 
           set; 
        }
    }
}
