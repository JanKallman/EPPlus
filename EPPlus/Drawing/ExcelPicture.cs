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
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.IO.Packaging;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// An image object
    /// </summary>
    public sealed class ExcelPicture : ExcelDrawing
    {
        #region "Constructors"
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlNode picNode = node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip", drawings.NameSpaceManager);
            if (picNode != null)
            {
                RelPic = drawings.Part.GetRelationship(picNode.Attributes["r:embed"].Value);
                UriPic = PackUriHelper.ResolvePartUri(drawings.UriDrawing, RelPic.TargetUri);

                Part = drawings.Part.Package.GetPart(UriPic);
                FileInfo f = new FileInfo(UriPic.OriginalString);
                ContentType = GetContentType(f.Extension);
                _image = Image.FromStream(Part.GetStream());
                ImageConverter ic=new ImageConverter();
                var iby=(byte[])ic.ConvertTo(_image, typeof(byte[]));
                var ii = _drawings._package.LoadImage(iby, UriPic, Part);
                ImageHash = ii.Hash;

                string relID = GetXmlNodeString("xdr:pic/xdr:nvPicPr/xdr:cNvPr/a:hlinkClick/@r:id");
                if (!string.IsNullOrEmpty(relID))
                {
                    HypRel = drawings.Part.GetRelationship(relID);
                    if (HypRel.TargetUri.IsAbsoluteUri)
                    {
                        _hyperlink = new ExcelHyperLink(HypRel.TargetUri.AbsoluteUri);
                    }
                    else
                    {
                        _hyperlink = new ExcelHyperLink(HypRel.TargetUri.OriginalString, UriKind.Relative);
                    }
                    ((ExcelHyperLink)_hyperlink).ToolTip = GetXmlNodeString("xdr:pic/xdr:nvPicPr/xdr:cNvPr/a:hlinkClick/@tooltip");
                }
            }
        }
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Image image) :
           this(drawings, node, image, null)
        {
        }
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, Image image, Uri hyperlink) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlElement picNode = node.OwnerDocument.CreateElement("xdr", "pic", ExcelPackage.schemaSheetDrawings);
            node.InsertAfter(picNode,node.SelectSingleNode("xdr:to",NameSpaceManager));
            _hyperlink = hyperlink;
            picNode.InnerXml = PicStartXml();

            node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings), picNode);

            Package package = drawings.Worksheet._package.Package;
            //Get the picture if it exists or save it if not.
            _image = image;
            string relID = SavePicture(image);

            //Create relationship
            node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;

            SetPosDefaults(image);
            package.Flush();
        }
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, FileInfo imageFile) :
           this(drawings,node,imageFile,null)
        {
        }
        internal ExcelPicture(ExcelDrawings drawings, XmlNode node, FileInfo imageFile, Uri hyperlink) :
            base(drawings, node, "xdr:pic/xdr:nvPicPr/xdr:cNvPr/@name")
        {
            XmlElement picNode = node.OwnerDocument.CreateElement("xdr", "pic", ExcelPackage.schemaSheetDrawings);
            node.InsertAfter(picNode, node.SelectSingleNode("xdr:to", NameSpaceManager));
            _hyperlink = hyperlink;
            picNode.InnerXml = PicStartXml();

            node.InsertAfter(node.OwnerDocument.CreateElement("xdr", "clientData", ExcelPackage.schemaSheetDrawings), picNode);

            //Changed to stream 2/4-13 (issue 14834). Thnx SClause
            var package = drawings.Worksheet._package.Package;
            ContentType = GetContentType(imageFile.Extension);
            var imagestream = new FileStream(imageFile.FullName, FileMode.Open, FileAccess.Read);
            _image = Image.FromStream(imagestream);
            ImageConverter ic = new ImageConverter();
            var img = (byte[])ic.ConvertTo(_image, typeof(byte[]));
            imagestream.Close();

            UriPic = GetNewUri(package, "/xl/media/{0}" + imageFile.Name);
            var ii = _drawings._package.AddImage(img, UriPic, ContentType);
            string relID;
            if(!drawings._hashes.ContainsKey(ii.Hash))
            {
                Part = ii.Part;
                RelPic = drawings.Part.CreateRelationship(PackUriHelper.GetRelativeUri(drawings.UriDrawing, ii.Uri), TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                relID = RelPic.Id;
                _drawings._hashes.Add(ii.Hash, relID);
                AddNewPicture(img, relID);
            }
            else
            {
                relID = drawings._hashes[ii.Hash];
                var rel = _drawings.Part.GetRelationship(relID);
                UriPic = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            }
            SetPosDefaults(Image);
            //Create relationship
            node.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
            package.Flush();
        }

        internal static string GetContentType(string extension)
        {
            switch (extension.ToLower())
            {
                case ".bmp":
                    return  "image/bmp";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".png":
                    return "image/png";
                case ".cgm":
                    return "image/cgm";
                case ".emf":
                    return "image/x-emf";
                case ".eps":
                    return "image/x-eps";
                case ".pcx":
                    return "image/x-pcx";
                case ".tga":
                    return "image/x-tga";
                case ".tif":
                case ".tiff":
                    return "image/x-tiff";
                case ".wmf":
                    return "image/x-wmf";
                default:
                    return "image/jpeg";

            }
        }
        //Add a new image to the compare collection
        private void AddNewPicture(byte[] img, string relID)
        {
            var newPic = new ExcelDrawings.ImageCompare();
            newPic.image = img;
            newPic.relID = relID;
            //_drawings._pics.Add(newPic);
        }
        #endregion
        private string SavePicture(Image image)
        {
            ImageConverter ic = new ImageConverter();
            byte[] img = (byte[])ic.ConvertTo(image, typeof(byte[]));
            var ii = _drawings._package.AddImage(img);

            if (_drawings._hashes.ContainsKey(ii.Hash))
            {
                var relID = _drawings._hashes[ii.Hash];
                var rel = _drawings.Part.GetRelationship(relID);
                UriPic = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                return relID;
            }
            else
            {
                UriPic = ii.Uri;
            }

            //Set the Image and save it to the package.
            RelPic = _drawings.Part.CreateRelationship(PackUriHelper.GetRelativeUri(_drawings.UriDrawing, UriPic), TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            
            //AddNewPicture(img, picRelation.Id);
            _drawings._hashes.Add(ii.Hash, RelPic.Id);
            ImageHash = ii.Hash;

            return RelPic.Id;
        }
        private void SetPosDefaults(Image image)
        {
            EditAs = eEditAs.OneCell;
            SetPixelWidth(image.Width, image.HorizontalResolution);
            SetPixelHeight(image.Height, image.VerticalResolution);
        }

        private string PicStartXml()
        {
            StringBuilder xml = new StringBuilder();
            
            xml.Append("<xdr:nvPicPr>");
            
            if (_hyperlink == null)
            {
                xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\" />", _id);
            }
            else
            {
               HypRel = _drawings.Part.CreateRelationship(_hyperlink, TargetMode.External, ExcelPackage.schemaHyperlink);
               xml.AppendFormat("<xdr:cNvPr id=\"{0}\" descr=\"\">", _id);
               if (HypRel != null)
               {
                   if (_hyperlink is ExcelHyperLink)
                   {
                       xml.AppendFormat("<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" tooltip=\"{1}\"/>",
                         HypRel.Id, ((ExcelHyperLink)_hyperlink).ToolTip);
                   }
                   else
                   {
                       xml.AppendFormat("<a:hlinkClick xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{0}\" />",
                         HypRel.Id);
                   }
               }
               xml.Append("</xdr:cNvPr>");
            }
           
            xml.Append("<xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\" /></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"\" cstate=\"print\" /><a:stretch><a:fillRect /> </a:stretch> </xdr:blipFill> <xdr:spPr> <a:xfrm> <a:off x=\"0\" y=\"0\" />  <a:ext cx=\"0\" cy=\"0\" /> </a:xfrm> <a:prstGeom prst=\"rect\"> <a:avLst /> </a:prstGeom> </xdr:spPr>");

            return xml.ToString();
        }

        internal string ImageHash { get; set; }
        Image _image = null;
        /// <summary>
        /// The Image
        /// </summary>
        public Image Image 
        {
            get
            {
                return _image;
            }
            set
            {
                if (value != null)
                {
                    _image = value;
                    try
                    {
                        string relID = SavePicture(value);

                        //Create relationship
                        TopNode.SelectSingleNode("xdr:pic/xdr:blipFill/a:blip/@r:embed", NameSpaceManager).Value = relID;
                        //_image.Save(Part.GetStream(FileMode.Create, FileAccess.Write), _imageFormat);   //Always JPEG here at this point. 
                    }
                    catch(Exception ex)
                    {
                        throw(new Exception("Can't save image - " + ex.Message, ex));
                    }
                }
            }
        }
        ImageFormat _imageFormat=ImageFormat.Jpeg;
        /// <summary>
        /// Image format
        /// If the picture is created from an Image this type is always Jpeg
        /// </summary>
        public ImageFormat ImageFormat
        {
            get
            {
                return _imageFormat;
            }
            internal set
            {
                _imageFormat = value;
            }
        }
        internal string ContentType
        {
            get;
            set;
        }
        /// <summary>
        /// Set the size of the image in percent from the orginal size
        /// Note that resizing columns / rows after using this function will effect the size of the picture
        /// </summary>
        /// <param name="Percent">Percent</param>
        public override void SetSize(int Percent)
        {
            if(Image == null)
            {
                base.SetSize(Percent);
            }
            else
            {
                int width = Image.Width;
                int height = Image.Height;

                width = (int)(width * ((decimal)Percent / 100));
                height = (int)(height * ((decimal)Percent / 100));

                SetPixelWidth(width, Image.HorizontalResolution);
                SetPixelHeight(height, Image.VerticalResolution);
            }
        }
        internal Uri UriPic { get; set; }
        internal PackageRelationship RelPic {get; set;}
        internal PackageRelationship HypRel { get; set; }
        internal PackagePart Part;

        internal new string Id
        {
            get { return Name; }
        }
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Fill
        /// </summary>
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "xdr:pic/xdr:spPr");
                }
                return _fill;
            }
        }
        ExcelDrawingBorder _border = null;
        /// <summary>
        /// Border
        /// </summary>
        public ExcelDrawingBorder Border
        {
            get
            {
                if (_border == null)
                {
                    _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "xdr:pic/xdr:spPr/a:ln");
                }
                return _border;
            }
        }

        private Uri _hyperlink = null;
        /// <summary>
        /// Hyperlink
        /// </summary>
        public Uri Hyperlink
        {
           get
           {
              return _hyperlink;
           }
        }
        internal override void DeleteMe()
        {
            _drawings._package.RemoveImage(ImageHash);
            base.DeleteMe();
        }
    }
}