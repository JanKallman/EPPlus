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
 *******************************************************************************
 * Jan Källman		Initial Release		     
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
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
            if (!cell.Worksheet._vmlDrawings.ContainsKey(ExcelAddress.GetCellID(cell.Worksheet.SheetID, cell.Start.Row, cell.Start.Column)))
            {
                cell.Worksheet._vmlDrawings.Add(cell);
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
