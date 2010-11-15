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
                CommentXml.Load(Part.GetStream());
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
        public XmlDocument CommentXml { get; set; }
        internal Uri Uri { get; set; }
        internal string RelId { get; set; }
        internal XmlNamespaceManager NameSpaceManager { get; set; }
        internal PackagePart Part
        {
            get;
            set;
        }
        public ExcelWorksheet Worksheet
        {
            get;
            set;
        }
        public int Count
        {
            get
            {
                return _comments.Count;
            }
        }
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
        public ExcelComment this[ExcelCellAddress cell]
        {
            get
            {
                ulong cellID=ExcelCellBase.GetCellID(Worksheet.SheetID, cell.Row, cell.Column);
                if (_comments.IndexOf(cellID) > 0)
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
            CommentXml.SelectSingleNode("d:comments/d:commentList", NameSpaceManager).AppendChild(elem);
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
        #region IEnumerable Members

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _comments;
        }

        #endregion
    
}
}
