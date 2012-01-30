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
 * Jan Källman		Added		10-SEP-2009
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Drawing;
using System.IO.Packaging;
using System.IO;
using OfficeOpenXml.Drawing;
namespace OfficeOpenXml
{
    /// <summary>
    /// An image that fills the background of the worksheet.
    /// </summary>
    public class ExcelBackgroundImage : XmlHelper
    {
        ExcelWorksheet _workSheet;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="nsm"></param>
        /// <param name="topNode">The topnode of the worksheet</param>
        /// <param name="workSheet">Worksheet reference</param>
        internal  ExcelBackgroundImage(XmlNamespaceManager nsm, XmlNode topNode, ExcelWorksheet workSheet) :
            base(nsm, topNode)
        {
            _workSheet = workSheet;
        }
        
        const string BACKGROUNDPIC_PATH = "d:picture/@r:id";
        /// <summary>
        /// The background image of the worksheet. 
        /// The image will be saved internally as a jpg.
        /// </summary>
        public Image Image
        {
            get
            {
                string relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
                if (!string.IsNullOrEmpty(relID))
                {
                    var rel = _workSheet.Part.GetRelationship(relID);
                    var imagePart = _workSheet.Part.Package.GetPart(PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
                    return Image.FromStream(imagePart.GetStream());
                }
                return null;
            }
            set
            {
                DeletePrevImage();
                if (value == null)
                {
                    DeleteAllNode(BACKGROUNDPIC_PATH);
                }
                else
                {
                    ImageConverter ic = new ImageConverter();
                    byte[] img = (byte[])ic.ConvertTo(value, typeof(byte[]));
                    var ii = _workSheet.Workbook._package.AddImage(img);
                    var rel = _workSheet.Part.CreateRelationship(ii.Uri, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
                    SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
                }
            }
        }
        /// <summary>
        /// Set the picture from an image file. 
        /// The image file will be saved as a blob, so make sure Excel supports the image format.
        /// </summary>
        /// <param name="PictureFile">The image file.</param>
        public void SetFromFile(FileInfo PictureFile)
        {
            DeletePrevImage();

            Image img;
            try
            {
                img = Image.FromFile(PictureFile.FullName);
            }
            catch (Exception ex)
            {
                throw (new InvalidDataException("File is not a supported image-file or is corrupt", ex));
            }

            ImageConverter ic = new ImageConverter();
            string contentType = ExcelPicture.GetContentType(PictureFile.Extension);
            var imageURI = XmlHelper.GetNewUri(_workSheet._package.Package, "/xl/media/" + PictureFile.Name.Substring(0, PictureFile.Name.Length - PictureFile.Extension.Length) + "{0}" + PictureFile.Extension);

            byte[] fileBytes = (byte[])ic.ConvertTo(img, typeof(byte[]));
            var ii = _workSheet.Workbook._package.AddImage(fileBytes, imageURI, contentType);


            if (_workSheet.Part.Package.PartExists(imageURI) && ii.RefCount==1) //The file exists with another content, overwrite it.
            {
                //Remove the part if it exists
                _workSheet.Part.Package.DeletePart(imageURI);
            }

            var imagePart = _workSheet.Part.Package.CreatePart(imageURI, contentType, CompressionOption.NotCompressed);
            //Save the picture to package.

            var strm = imagePart.GetStream(FileMode.Create, FileAccess.Write);
            strm.Write(fileBytes, 0, fileBytes.Length);

            var rel = _workSheet.Part.CreateRelationship(imageURI, TargetMode.Internal, ExcelPackage.schemaRelationships + "/image");
            SetXmlNodeString(BACKGROUNDPIC_PATH, rel.Id);
        }
        private void DeletePrevImage()
        {
            var relID = GetXmlNodeString(BACKGROUNDPIC_PATH);
            if (relID != "")
            {
                var ic = new ImageConverter();
                byte[] img = (byte[])ic.ConvertTo(Image, typeof(byte[]));
                var ii = _workSheet.Workbook._package.GetImageInfo(img);

                //Delete the relation
                _workSheet.Part.DeleteRelationship(relID);
                
                //Delete the image if there are no other references.
                if (ii != null && ii.RefCount == 1)
                {
                    if (_workSheet.Part.Package.PartExists(ii.Uri))
                    {
                        _workSheet.Part.Package.DeletePart(ii.Uri);
                    }
                }
                
            }
        }
    }
}
