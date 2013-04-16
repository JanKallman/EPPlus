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
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO.Packaging;
using System.Collections;

namespace OfficeOpenXml
{
    /// <summary>
    /// Collection of Excelcomment objects
    /// </summary>  
    public class ExcelCommentCollection : IEnumerable
    {
        internal RangeCollection _comments;
        internal ExcelCommentCollection(ExcelPackage pck, ExcelWorksheet ws, XmlNamespaceManager ns)
        {
            CommentXml = new XmlDocument();
            CommentXml.PreserveWhitespace = false;
            NameSpaceManager=ns;
            Worksheet=ws;
            CreateXml(pck);
            AddCommentsFromXml();
        }
        private void CreateXml(ExcelPackage pck)
        {
            var commentParts = Worksheet.Part.GetRelationshipsByType(ExcelPackage.schemaComment);
            bool isLoaded=false;
            CommentXml=new XmlDocument();
            foreach(var commentPart in commentParts)
            {
                Uri = PackUriHelper.ResolvePartUri(commentPart.SourceUri, commentPart.TargetUri);
                Part = pck.Package.GetPart(Uri);
                XmlHelper.LoadXmlSafe(CommentXml, Part.GetStream()); 
                RelId = commentPart.Id;
                isLoaded=true;
            }
            //Create a new document
            if(!isLoaded)
            {
                CommentXml.LoadXml("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors /><commentList /></comments>");
                Uri = null;
            }
        }
        private void AddCommentsFromXml()
        {
            var lst = new List<IRangeID>();
            foreach (XmlElement node in CommentXml.SelectNodes("//d:commentList/d:comment", NameSpaceManager))
            {
                var comment = new ExcelComment(NameSpaceManager, node, new ExcelRangeBase(Worksheet, node.GetAttribute("ref")));
                lst.Add(comment);
            }
            _comments = new RangeCollection(lst);
        }
        /// <summary>
        /// Access to the comment xml document
        /// </summary>
        public XmlDocument CommentXml { get; set; }
        internal Uri Uri { get; set; }
        internal string RelId { get; set; }
        internal XmlNamespaceManager NameSpaceManager { get; set; }
        internal PackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// A reference to the worksheet object
        /// </summary>
        public ExcelWorksheet Worksheet
        {
            get;
            set;
        }
        /// <summary>
        /// Number of comments in the collection
        /// </summary>
        public int Count
        {
            get
            {
                return _comments.Count;
            }
        }
        /// <summary>
        /// Indexer for the comments collection
        /// </summary>
        /// <param name="Index">The index</param>
        /// <returns>The comment</returns>
        public ExcelComment this[int Index]
        {
            get
            {
                if (Index < 0 || Index >= _comments.Count)
                {
                    throw(new ArgumentOutOfRangeException("Comment index out of range"));
                }
                return _comments[Index] as ExcelComment;
            }
        }
        /// <summary>
        /// Indexer for the comments collection
        /// </summary>
        /// <param name="cell">The cell</param>
        /// <returns>The comment</returns>
        public ExcelComment this[ExcelCellAddress cell]
        {
            get
            {
                ulong cellID=ExcelCellBase.GetCellID(Worksheet.SheetID, cell.Row, cell.Column);
                if (_comments.IndexOf(cellID) >= 0)
                {
                    return _comments[cellID] as ExcelComment;
                }
                else
                {
                    return null;
                }
            }
        }
        /// <summary>
        /// Adds a comment to the top left cell of the range
        /// </summary>
        /// <param name="cell">The cell</param>
        /// <param name="Text">The comment text</param>
        /// <param name="author">Author</param>
        /// <returns>The comment</returns>
        public ExcelComment Add(ExcelRangeBase cell, string Text, string author)
        {            
            var elem = CommentXml.CreateElement("comment", ExcelPackage.schemaMain);
            int ix=_comments.IndexOf(ExcelAddress.GetCellID(Worksheet.SheetID, cell._fromRow, cell._fromCol));
            //Make sure the nodes come on order.
            if (ix < 0 && (~ix < _comments.Count))
            {
                ix = ~ix;
                var preComment = _comments[ix] as ExcelComment;
                preComment._commentHelper.TopNode.ParentNode.InsertBefore(elem, preComment._commentHelper.TopNode);
            }
            else
            {
                CommentXml.SelectSingleNode("d:comments/d:commentList", NameSpaceManager).AppendChild(elem);
            }
            elem.SetAttribute("ref", cell.Start.Address);
            ExcelComment comment = new ExcelComment(NameSpaceManager, elem , cell);
            comment.RichText.Add(Text);
            if(author!="") 
            {
                comment.Author=author;
            }
            _comments.Add(comment);
            return comment;
        }
        /// <summary>
        /// Removes the comment
        /// </summary>
        /// <param name="comment">The comment to remove</param>
        public void Remove(ExcelComment comment)
        {
            ulong id = ExcelAddress.GetCellID(Worksheet.SheetID, comment.Range._fromRow, comment.Range._fromCol);
            int ix=_comments.IndexOf(id);
            if (ix >= 0 && comment == _comments[ix])
            {
                comment.TopNode.ParentNode.RemoveChild(comment.TopNode); //Remove VML
                comment._commentHelper.TopNode.ParentNode.RemoveChild(comment._commentHelper.TopNode); //Remove Comment

                Worksheet.VmlDrawingsComments._drawings.Delete(id);
                _comments.Delete(id);
            }
            else
            {
                throw (new ArgumentException("Comment does not exist in the worksheet"));
            }
        }
        /// <summary>
        /// Removes the comment at the specified position
        /// </summary>
        /// <param name="Index">The index</param>
        public void RemoveAt(int Index)
        {
            Remove(this[Index]);
        }
        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _comments;
        }
        #endregion


    }
}
