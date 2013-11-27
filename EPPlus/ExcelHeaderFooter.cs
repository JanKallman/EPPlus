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
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman                      Total rewrite               2010-03-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 * *******************************************************************************/
using System;
using System.Xml;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Collections.Generic;
using OfficeOpenXml.Drawing.Vml;
using System.IO;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Utils;
namespace OfficeOpenXml
{    
    /// <summary>
    /// How a picture will be aligned in the header/footer
    /// </summary>
    public enum PictureAlignment
    {
        /// <summary>
        /// The picture will be added to the left aligned text
        /// </summary>
        Left,
        /// <summary>
        /// The picture will be added to the centered text
        /// </summary>
        Centered,
        /// <summary>
        /// The picture will be added to the right aligned text
        /// </summary>
        Right
    }
    #region class ExcelHeaderFooterText
	/// <summary>
    /// Print header and footer 
    /// </summary>
	public class ExcelHeaderFooterText
	{
        ExcelWorksheet _ws;
        string _hf;
        internal ExcelHeaderFooterText(XmlNode TextNode, ExcelWorksheet ws, string hf)
        {
            _ws = ws;
            _hf = hf;
            if (TextNode == null || string.IsNullOrEmpty(TextNode.InnerText)) return;
            string text = TextNode.InnerText;
            string code = text.Substring(0, 2);  
            int startPos=2;
            for (int pos=startPos;pos<text.Length-2;pos++)
            {
                string newCode = text.Substring(pos, 2);
                if (newCode == "&C" || newCode == "&R")
                {
                    SetText(code, text.Substring(startPos, pos-startPos));
                    startPos = pos+2;
                    pos = startPos;
                    code = newCode;
                }
            }
            SetText(code, text.Substring(startPos, text.Length - startPos));
        }
        private void SetText(string code, string text)
        {
            switch (code)
            {
                case "&L":
                    LeftAlignedText=text;
                    break;
                case "&C":
                    CenteredText=text;
                    break;
                default:
                    RightAlignedText=text;
                    break;
            }
        }
		/// <summary>
		/// Get/set the text to appear on the left hand side of the header (or footer) on the worksheet.
		/// </summary>
		public string LeftAlignedText = null;
		/// <summary>
        /// Get/set the text to appear in the center of the header (or footer) on the worksheet.
		/// </summary>
		public string CenteredText = null;
		/// <summary>
        /// Get/set the text to appear on the right hand side of the header (or footer) on the worksheet.
		/// </summary>
		public string RightAlignedText = null;
        /// <summary>
        /// Inserts a picture at the end of the text in the header or footer
        /// </summary>
        /// <param name="Picture">The image object containing the Picture</param>
        /// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
        public ExcelVmlDrawingPicture InsertPicture(Image Picture, PictureAlignment Alignment)
        {
            string id = ValidateImage(Alignment);
            
            //Add the image
            ImageConverter ic = new ImageConverter();
            byte[] img = (byte[])ic.ConvertTo(Picture, typeof(byte[]));
            var ii = _ws.Workbook._package.AddImage(img);

            return AddImage(Picture, id, ii);
        }
        /// <summary>
        /// Inserts a picture at the end of the text in the header or footer
        /// </summary>
        /// <param name="PictureFile">The image object containing the Picture</param>
        /// <param name="Alignment">Alignment. The image object will be inserted at the end of the Text.</param>
        public ExcelVmlDrawingPicture InsertPicture(FileInfo PictureFile, PictureAlignment Alignment)
        {
            string id = ValidateImage(Alignment);

            Image Picture;
            try
            {
                if (!PictureFile.Exists)
                {
                    throw (new FileNotFoundException(string.Format("{0} is missing", PictureFile.FullName)));
                }
                Picture = Image.FromFile(PictureFile.FullName);
            }
            catch (Exception ex)
            {
                throw (new InvalidDataException("File is not a supported image-file or is corrupt", ex));
            }

            ImageConverter ic = new ImageConverter();
            string contentType = ExcelPicture.GetContentType(PictureFile.Extension);
            var uriPic = XmlHelper.GetNewUri(_ws._package.Package, "/xl/media/"+PictureFile.Name.Substring(0, PictureFile.Name.Length-PictureFile.Extension.Length) + "{0}" + PictureFile.Extension);
            byte[] imgBytes = (byte[])ic.ConvertTo(Picture, typeof(byte[]));
            var ii = _ws.Workbook._package.AddImage(imgBytes, uriPic, contentType);

            return AddImage(Picture, id, ii);
        }

        private ExcelVmlDrawingPicture AddImage(Image Picture, string id, ExcelPackage.ImageInfo ii)
        {
            double width = Picture.Width * 72 / Picture.HorizontalResolution,      //Pixel --> Points
                   height = Picture.Height * 72 / Picture.VerticalResolution;      //Pixel --> Points
            //Add VML-drawing            
            return _ws.HeaderFooter.Pictures.Add(id, ii.Uri, "", width, height);
        }
        private string ValidateImage(PictureAlignment Alignment)
        {
            string id = string.Concat(Alignment.ToString()[0], _hf);
            foreach (ExcelVmlDrawingPicture image in _ws.HeaderFooter.Pictures)
            {
                if (image.Id == id)
                {
                    throw (new InvalidOperationException("A picture already exists in this section"));
                }
            }
            //Add the image placeholder to the end of the text
            switch (Alignment)
            {
                case PictureAlignment.Left:
                    LeftAlignedText += ExcelHeaderFooter.Image;
                    break;
                case PictureAlignment.Centered:
                    CenteredText += ExcelHeaderFooter.Image;
                    break;
                default:
                    RightAlignedText += ExcelHeaderFooter.Image;
                    break;
            }
            return id;
        }
	}
	#endregion

	#region ExcelHeaderFooter
	/// <summary>
	/// Represents the Header and Footer on an Excel Worksheet
	/// </summary>
	public sealed class ExcelHeaderFooter : XmlHelper
	{
		#region Static Properties
		/// <summary>
        /// The code for "current page #"
		/// </summary>
		public const string PageNumber = @"&P";
		/// <summary>
        /// The code for "total pages"
		/// </summary>
		public const string NumberOfPages = @"&N";
        /// <summary>
        /// The code for "text font color"
        /// RGB Color is specified as RRGGBB
        /// Theme Color is specified as TTSNN where TT is the theme color Id, S is either "+" or "-" of the tint/shade value, NN is the tint/shade value.
        /// </summary>
        public const string FontColor = @"&K";
		/// <summary>
        /// The code for "sheet tab name"
		/// </summary>
		public const string SheetName = @"&A";
		/// <summary>
        /// The code for "this workbook's file path"
		/// </summary>
		public const string FilePath = @"&Z";
		/// <summary>
        /// The code for "this workbook's file name"
		/// </summary>
		public const string FileName = @"&F";
		/// <summary>
        /// The code for "date"
		/// </summary>
		public const string CurrentDate = @"&D";
		/// <summary>
        /// The code for "time"
		/// </summary>
		public const string CurrentTime = @"&T";
        /// <summary>
        /// The code for "picture as background"
        /// </summary>
        public const string Image = @"&G";
        /// <summary>
        /// The code for "outline style"
        /// </summary>
        public const string OutlineStyle = @"&O";
        /// <summary>
        /// The code for "shadow style"
        /// </summary>
        public const string ShadowStyle = @"&H";
		#endregion

		#region ExcelHeaderFooter Private Properties
		internal ExcelHeaderFooterText _oddHeader;
        internal ExcelHeaderFooterText _oddFooter;
		internal ExcelHeaderFooterText _evenHeader;
        internal ExcelHeaderFooterText _evenFooter;
        internal ExcelHeaderFooterText _firstHeader;
        internal ExcelHeaderFooterText _firstFooter;
        private ExcelWorksheet _ws;
        #endregion

		#region ExcelHeaderFooter Constructor
		/// <summary>
		/// ExcelHeaderFooter Constructor
		/// </summary>
		/// <param name="nameSpaceManager"></param>
        /// <param name="topNode"></param>
        /// <param name="ws">The worksheet</param>
		internal ExcelHeaderFooter(XmlNamespaceManager nameSpaceManager, XmlNode topNode, ExcelWorksheet ws) :
            base(nameSpaceManager, topNode)
		{
            _ws = ws;
            SchemaNodeOrder = new string[] { "headerFooter", "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter" };
		}
		#endregion

		#region alignWithMargins
        const string alignWithMarginsPath="@alignWithMargins";
        /// <summary>
		/// Gets/sets the alignWithMargins attribute
		/// </summary>
		public bool AlignWithMargins
		{
			get
			{
                return GetXmlNodeBool(alignWithMarginsPath);
			}
			set
			{
                SetXmlNodeString(alignWithMarginsPath, value ? "1" : "0");
			}
		}
		#endregion

        #region differentOddEven
        const string differentOddEvenPath = "@differentOddEven";
        /// <summary>
		/// Gets/sets the flag that tells Excel to display different headers and footers on odd and even pages.
		/// </summary>
		public bool differentOddEven
		{
			get
			{
                return GetXmlNodeBool(differentOddEvenPath);
			}
			set
			{
                SetXmlNodeString(differentOddEvenPath, value ? "1" : "0");
			}
		}
		#endregion

		#region differentFirst
        const string differentFirstPath = "@differentFirst";

		/// <summary>
		/// Gets/sets the flag that tells Excel to display different headers and footers on the first page of the worksheet.
		/// </summary>
		public bool differentFirst
		{
			get
			{
                return GetXmlNodeBool(differentFirstPath);
			}
			set
			{
                SetXmlNodeString(differentFirstPath, value ? "1" : "0");
			}
		}
		#endregion

		#region ExcelHeaderFooter Public Properties
		/// <summary>
		/// Provides access to the header on odd numbered pages of the document.
		/// If you want the same header on both odd and even pages, then only set values in this ExcelHeaderFooterText class.
		/// </summary>
		public ExcelHeaderFooterText OddHeader 
        { 
            get 
            {
                if (_oddHeader == null)
                {
                    _oddHeader = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:oddHeader", NameSpaceManager), _ws, "H");
                }
                return _oddHeader; } 
        }
		/// <summary>
		/// Provides access to the footer on odd numbered pages of the document.
		/// If you want the same footer on both odd and even pages, then only set values in this ExcelHeaderFooterText class.
		/// </summary>
		public ExcelHeaderFooterText OddFooter 
        { 
            get 
            {
                if (_oddFooter == null)
                {
                    _oddFooter = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:oddFooter", NameSpaceManager), _ws, "F"); ;
                }
                return _oddFooter; 
            } 
        }
		// evenHeader and evenFooter set differentOddEven = true
		/// <summary>
		/// Provides access to the header on even numbered pages of the document.
		/// </summary>
		public ExcelHeaderFooterText EvenHeader 
        { 
            get 
            {
                if (_evenHeader == null)
                {
                    _evenHeader = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:evenHeader", NameSpaceManager), _ws, "HEVEN");
                    differentOddEven = true;
                }
                return _evenHeader; 
            } 
        }
		/// <summary>
		/// Provides access to the footer on even numbered pages of the document.
		/// </summary>
		public ExcelHeaderFooterText EvenFooter
        { 
            get 
            {
                if (_evenFooter == null)
                {
                    _evenFooter = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:evenFooter", NameSpaceManager), _ws, "FEVEN");
                    differentOddEven = true;
                }
                return _evenFooter ; 
            } 
        }
		/// <summary>
		/// Provides access to the header on the first page of the document.
		/// </summary>
		public ExcelHeaderFooterText FirstHeader
        { 
            get 
            {
                if (_firstHeader == null)
                {
                    _firstHeader = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:firstHeader", NameSpaceManager), _ws, "HFIRST"); 
                     differentFirst = true;
                }
                return _firstHeader; 
            } 
        }
		/// <summary>
		/// Provides access to the footer on the first page of the document.
		/// </summary>
		public ExcelHeaderFooterText FirstFooter
        { 
            get 
            {
                if (_firstFooter == null)
                {
                    _firstFooter = new ExcelHeaderFooterText(TopNode.SelectSingleNode("d:firstFooter", NameSpaceManager), _ws, "FFIRST"); 
                    differentFirst = true;
                }
                return _firstFooter; 
            } 
        }
        private ExcelVmlDrawingPictureCollection _vmlDrawingsHF = null;
        /// <summary>
        /// Vml drawings. Underlaying object for Header footer images
        /// </summary>
        public ExcelVmlDrawingPictureCollection Pictures
        {
            get
            {
                if (_vmlDrawingsHF == null)
                {
                    var vmlNode = _ws.WorksheetXml.SelectSingleNode("d:worksheet/d:legacyDrawingHF/@r:id", NameSpaceManager);
                    if (vmlNode == null)
                    {
                        _vmlDrawingsHF = new ExcelVmlDrawingPictureCollection(_ws._package, _ws, null);
                    }
                    else
                    {
                        if (_ws.Part.RelationshipExists(vmlNode.Value))
                        {
                            var rel = _ws.Part.GetRelationship(vmlNode.Value);
                            var vmlUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);

                            _vmlDrawingsHF = new ExcelVmlDrawingPictureCollection(_ws._package, _ws, vmlUri);
                            _vmlDrawingsHF.RelId = rel.Id;
                        }
                    }
                }
                return _vmlDrawingsHF;
            }
        }
		#endregion
		#region Save  //  ExcelHeaderFooter
		/// <summary>
		/// Saves the header and footer information to the worksheet XML
		/// </summary>
		internal void Save()
		{
			if (_oddHeader != null)
			{
                SetXmlNodeString("d:oddHeader", GetText(OddHeader));
			}
			if (_oddFooter != null)
			{
                SetXmlNodeString("d:oddFooter", GetText(OddFooter));
			}

			// only set evenHeader and evenFooter 
			if (differentOddEven)
			{
				if (_evenHeader != null)
				{
                    SetXmlNodeString("d:evenHeader", GetText(EvenHeader));
				}
				if (_evenFooter != null)
				{
                    SetXmlNodeString("d:evenFooter", GetText(EvenFooter));
				}
			}

			// only set firstHeader and firstFooter
			if (differentFirst)
			{
				if (_firstHeader != null)
				{
                    SetXmlNodeString("d:firstHeader", GetText(FirstHeader));
				}
				if (_firstFooter != null)
				{
                    SetXmlNodeString("d:firstFooter", GetText(FirstFooter));
				}
			}
		}
        internal void SaveHeaderFooterImages()
        {
            if (_vmlDrawingsHF != null)
            {
                if (_vmlDrawingsHF.Count == 0)
                {
                    if (_vmlDrawingsHF.Uri != null)
                    {
                        _ws.Part.DeleteRelationship(_vmlDrawingsHF.RelId);
                        _ws._package.Package.DeletePart(_vmlDrawingsHF.Uri);
                    }
                }
                else
                {
                    if (_vmlDrawingsHF.Uri == null)
                    {
                        _vmlDrawingsHF.Uri = XmlHelper.GetNewUri(_ws._package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
                    }
                    if (_vmlDrawingsHF.Part == null)
                    {
                        _vmlDrawingsHF.Part = _ws._package.Package.CreatePart(_vmlDrawingsHF.Uri, "application/vnd.openxmlformats-officedocument.vmlDrawing", _ws._package.Compression);
                        var rel = _ws.Part.CreateRelationship(UriHelper.GetRelativeUri(_ws.WorksheetUri, _vmlDrawingsHF.Uri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/vmlDrawing");
                        _ws.SetHFLegacyDrawingRel(rel.Id);
                        _vmlDrawingsHF.RelId = rel.Id;
                        foreach (ExcelVmlDrawingPicture draw in _vmlDrawingsHF)
                        {
                            rel = _vmlDrawingsHF.Part.CreateRelationship(UriHelper.GetRelativeUri(_vmlDrawingsHF.Uri, draw.ImageUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                            draw.RelId = rel.Id;
                        }
                    }
                    _vmlDrawingsHF.VmlDrawingXml.Save(_vmlDrawingsHF.Part.GetStream());
                }
            }
        }
		private string GetText(ExcelHeaderFooterText headerFooter)
		{
			string ret = "";
			if (headerFooter.LeftAlignedText != null)
				ret += "&L" + headerFooter.LeftAlignedText;
			if (headerFooter.CenteredText != null)
				ret += "&C" + headerFooter.CenteredText;
			if (headerFooter.RightAlignedText != null)
				ret += "&R" + headerFooter.RightAlignedText;
			return ret;
		}
		#endregion
	}
	#endregion
}
