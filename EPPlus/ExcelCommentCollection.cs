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
 *******************************************************************************
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using OfficeOpenXml.Utils;
namespace OfficeOpenXml
{
    /// <summary>
    /// Collection of Excelcomment objects
    /// </summary>  
    public class ExcelCommentCollection : IEnumerable, IDisposable
    {
        //internal RangeCollection _comments;
        List<ExcelComment> _list=new List<ExcelComment>();
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
                Uri = UriHelper.ResolvePartUri(commentPart.SourceUri, commentPart.TargetUri);
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
            //var lst = new List<IRangeID>();
            foreach (XmlElement node in CommentXml.SelectNodes("//d:commentList/d:comment", NameSpaceManager))
            {
                var comment = new ExcelComment(NameSpaceManager, node, new ExcelRangeBase(Worksheet, node.GetAttribute("ref")));
                //lst.Add(comment);
                _list.Add(comment);
                Worksheet._commentsStore.SetValue(comment.Range._fromRow, comment.Range._fromCol, _list.Count-1);
            }
            //_comments = new RangeCollection(lst);
        }
        /// <summary>
        /// Access to the comment xml document
        /// </summary>
        public XmlDocument CommentXml { get; set; }
        internal Uri Uri { get; set; }
        internal string RelId { get; set; }
        internal XmlNamespaceManager NameSpaceManager { get; set; }
        internal Packaging.ZipPackagePart Part
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
                return _list.Count;
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
                if (Index < 0 || Index >= _list.Count)
                {
                    throw(new ArgumentOutOfRangeException("Comment index out of range"));
                }
                return _list[Index] as ExcelComment;
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
                //ulong cellID=ExcelCellBase.GetCellID(Worksheet.SheetID, cell.Row, cell.Column);
                //if (_comments.IndexOf(cellID) >= 0)
                //{
                //    return _comments[cellID] as ExcelComment;
                //}
                //else
                //{
                //    return null;
                //}
                int i=-1;
                if (Worksheet._commentsStore.Exists(cell.Row, cell.Column, ref i))
                {
                    return _list[i];
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
            //int ix=_comments.IndexOf(ExcelAddress.GetCellID(Worksheet.SheetID, cell._fromRow, cell._fromCol));
            //Make sure the nodes come on order.
            int row=cell.Start.Row, column= cell.Start.Column;
            ExcelComment nextComment = null;
            if (Worksheet._commentsStore.NextCell(ref row, ref column))
            {
                nextComment = _list[Worksheet._commentsStore.GetValue(row, column)];
            }
            if(nextComment==null)
            {
                CommentXml.SelectSingleNode("d:comments/d:commentList", NameSpaceManager).AppendChild(elem);
            }
            else
            {
                nextComment._commentHelper.TopNode.ParentNode.InsertBefore(elem, nextComment._commentHelper.TopNode);
            }
            elem.SetAttribute("ref", cell.Start.Address);
            ExcelComment comment = new ExcelComment(NameSpaceManager, elem , cell);
            comment.RichText.Add(Text);
            if(author!="") 
            {
                comment.Author=author;
            }
            _list.Add(comment);
            Worksheet._commentsStore.SetValue(cell.Start.Row, cell.Start.Column, _list.Count-1);
            //Check if a value exists otherwise add one so it is saved when the cells collection is iterated
            if (!Worksheet.ExistsValueInner(cell._fromRow, cell._fromCol))
            {
                Worksheet.SetValueInner(cell._fromRow, cell._fromCol, null);
            }
            return comment;
        }
        /// <summary>
        /// Removes the comment
        /// </summary>
        /// <param name="comment">The comment to remove</param>
        public void Remove(ExcelComment comment)
        {
            ulong id = ExcelAddress.GetCellID(Worksheet.SheetID, comment.Range._fromRow, comment.Range._fromCol);
            //int ix=_comments.IndexOf(id);
            int i = -1;
            ExcelComment c=null;
            if (Worksheet._commentsStore.Exists(comment.Range._fromRow, comment.Range._fromCol, ref i))
            {
                c = _list[i];
            }
            if (comment==c)
            {
                comment.TopNode.ParentNode.RemoveChild(comment.TopNode); //Remove VML
                comment._commentHelper.TopNode.ParentNode.RemoveChild(comment._commentHelper.TopNode); //Remove Comment

                Worksheet.VmlDrawingsComments._drawings.Delete(id);
                _list.RemoveAt(i);                
                Worksheet._commentsStore.Delete(comment.Range._fromRow, comment.Range._fromCol, 1, 1, false);   //Issue 15549, Comments should not be shifted 
                var ci = new CellsStoreEnumerator<int>(Worksheet._commentsStore);
                while(ci.Next())
                {
                    if(ci.Value>i)
                    {
                        ci.Value -= 1;
                    }
                }
            }
            else
            {
                throw (new ArgumentException("Comment does not exist in the worksheet"));
            }
        }

        /// <summary>
        /// Shifts all comments based on their address and the location of inserted rows and columns.
        /// </summary>
        /// <param name="fromRow">The start row.</param>
        /// <param name="fromCol">The start column.</param>
        /// <param name="rows">The number of rows to insert.</param>
        /// <param name="columns">The number of columns to insert.</param>
        internal void Delete(int fromRow, int fromCol, int rows, int columns)
        {
            List<ExcelComment> deletedComments = new List<ExcelComment>();
            ExcelAddressBase address = null;
            foreach (ExcelComment comment in _list)
            {
                address = new ExcelAddressBase(comment.Address);
                if (fromCol>0 && address._fromCol >= fromCol)
                {
                    address = address.DeleteColumn(fromCol, columns);
                }
                if(fromRow > 0 && address._fromRow >= fromRow)
                {
                    address = address.DeleteRow(fromRow, rows);
                }
                if(address==null || address.Address=="#REF!")
                {
                    deletedComments.Add(comment);
                }
                else
                {
                    comment.Reference = address.Address;
                }
            }
            foreach(var comment in deletedComments)
            {
                Remove(comment);
            }
        }
        /// <summary>
        /// Shifts all comments based on their address and the location of inserted rows and columns.
        /// </summary>
        /// <param name="fromRow">The start row.</param>
        /// <param name="fromCol">The start column.</param>
        /// <param name="rows">The number of rows to insert.</param>
        /// <param name="columns">The number of columns to insert.</param>
        public void Insert(int fromRow, int fromCol, int rows, int columns)
        {
          //List<ExcelComment> commentsToShift = new List<ExcelComment>();
          foreach (ExcelComment comment in _list)
          {
              var address = new ExcelAddressBase(comment.Address);
              if (rows > 0 && address._fromRow >= fromRow)
              {
                  comment.Reference = comment.Range.AddRow(fromRow, rows).Address;
              }
              if(columns>0 && address._fromCol >= fromCol)
              {
                 comment.Reference = comment.Range.AddColumn(fromCol, columns).Address;
              }
          }
          //foreach (ExcelComment comment in commentsToShift)
          //{
          //  Remove(comment);
          //  var address = new ExcelAddressBase(comment.Address);
          //  if (address._fromRow >= fromRow)
          //    address._fromRow += rows;
          //  if (address._fromCol >= fromCol)
          //    address._fromCol += columns;
          //  Add(Worksheet.Cells[address._fromRow, address._fromCol], comment.Text, comment.Author);
          //}
        }

        void IDisposable.Dispose() 
        { 
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
            return _list.GetEnumerator();
        }
        #endregion

        internal void Clear()
        {
            while(Count>0)
            {
                RemoveAt(0);
            }
        }
    }
}
