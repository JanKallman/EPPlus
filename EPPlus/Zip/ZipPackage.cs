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
 * Jan Källman		Added		25-Oct-2012
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Ionic.Zip;
using Ionic.Zlib;
using System.Xml;
namespace OfficeOpenXml.Zip
{
    /// <summary>
    /// Specifies whether the target is inside or outside the System.IO.Packaging.Package.
    /// </summary>
    public enum TargetMode
    {
        /// <summary>
        /// The relationship references a part that is inside the package.
        /// </summary>
        Internal = 0,
        /// <summary>
        /// The relationship references a resource that is external to the package.
        /// </summary>
        External = 1,
    }
    public class ZipPackage : ZipRelationshipBase
    {
        internal class ContentType
        {
            internal string Name;
            internal bool IsExtension;
            internal string Match;
            public ContentType(string name, bool isExtension, string match)
            {
                Name = name;
                IsExtension = isExtension;
                Match = match;
            }
        }
        Dictionary<string, ZipPackagePart> Parts = new Dictionary<string, ZipPackagePart>();
        internal Dictionary<string, ContentType> _contentTypes = new Dictionary<string, ContentType>();
        internal ZipPackage()
        {
            AddNew();
        }

        private void AddNew()
        {
            _contentTypes.Add("xml", new ContentType(ExcelPackage.schemaXmlExtension, true, "xml"));
            _contentTypes.Add("rels", new ContentType(ExcelPackage.schemaRelsExtension, true, "rels"));
        }
        internal ZipPackage(string filePath)
        {
            using (ZipFile zip = new ZipFile(filePath))
            {
                foreach (var e in zip.Entries)
                {
                    Parts.Add(GetUriKey(e.FileName), new ZipPackagePart(this, e));
                }
            }
        }

        internal ZipPackage(Stream stream)
        {
            if (stream == null || stream.Length == 0)
            {
                AddNew();
            }
            else
            {
                var rels = new Dictionary<string, string>();
                stream.Seek(0, SeekOrigin.Begin);
                using (ZipInputStream zip = new ZipInputStream(stream))
                {
                    var e = zip.GetNextEntry();
                    while (e != null)
                    {
                        if (e.UncompressedSize > 0)
                        {
                            var b = new byte[e.UncompressedSize];
                            var size = zip.Read(b, 0, (int)e.UncompressedSize);
                            if (e.FileName.ToLower() == "[content_types].xml")
                            {
                                AddContentTypes(Encoding.UTF8.GetString(b));
                            }
                            else if (e.FileName.ToLower() == "_rels/.rels")
                            {
                                ReadRelation(Encoding.UTF8.GetString(b), "");
                            }
                            else
                            {
                                if (e.FileName.ToLower().EndsWith(".rels"))
                                {
                                    rels.Add(GetUriKey(e.FileName.ToLower()), Encoding.UTF8.GetString(b));
                                }
                                else
                                {
                                    var part = new ZipPackagePart(this, e);
                                    part.Stream = new MemoryStream(b);
                                    Parts.Add(GetUriKey(e.FileName), part);
                                }
                            }
                        }
                        else
                        {
                        }
                        e = zip.GetNextEntry();
                    }

                    foreach (var p in Parts)
                    {
                        FileInfo fi = new FileInfo(p.Key);
                        string relFile = string.Format("{0}_rels/{1}.rels", p.Key.Substring(0, p.Key.Length - fi.Name.Length), fi.Name);
                        if (rels.ContainsKey(relFile))
                        {
                            p.Value.ReadRelation(rels[relFile], p.Value.Uri.OriginalString);
                        }
                        if (_contentTypes.ContainsKey(p.Key))
                        {
                            p.Value.ContentType = _contentTypes[p.Key].Name;
                        }
                        else if (fi.Extension.Length > 1 && _contentTypes.ContainsKey(fi.Extension.Substring(1)))
                        {
                            p.Value.ContentType = _contentTypes[fi.Extension.Substring(1)].Name;
                        }
                    }
                }
            }
        }

        private void AddContentTypes(string xml)
        {
            var doc = new XmlDocument();
            XmlHelper.LoadXmlSafe(doc, xml, Encoding.UTF8);

            foreach (XmlElement c in doc.DocumentElement.ChildNodes)
            {
                ContentType ct;
                if (string.IsNullOrEmpty(c.GetAttribute("Extension")))
                {
                    ct = new ContentType(c.GetAttribute("ContentType"), false, c.GetAttribute("PartName"));
                }
                else
                {
                    ct = new ContentType(c.GetAttribute("ContentType"), true, c.GetAttribute("Extension"));
                }
                _contentTypes.Add(GetUriKey(ct.Match), ct);
            }
        }

        #region Methods
        internal ZipPackagePart CreatePart(Uri partUri, string contentType)
        {
            return CreatePart(partUri, contentType, CompressionLevel.Default);
        }
        internal ZipPackagePart CreatePart(Uri partUri, string contentType, CompressionLevel compressionLevel)
        {
            if (PartExists(partUri))
            {
                throw (new InvalidOperationException("Part already exist"));
            }

            var part = new ZipPackagePart(this, partUri, contentType, compressionLevel);
            _contentTypes.Add(GetUriKey(part.Uri.OriginalString), new ContentType(contentType, false, part.Uri.OriginalString));
            Parts.Add(GetUriKey(part.Uri.OriginalString), part);
            return part;
        }
        internal ZipPackagePart GetPart(Uri partUri)
        {
            if (PartExists(partUri))
            {
                return Parts[GetUriKey(partUri.OriginalString)];
            }
            else
            {
                throw (new InvalidOperationException("Part does not exist."));
            }
        }

        internal string GetUriKey(string uri)
        {
            string ret = uri.ToLower();
            if (ret[0] != '/')
            {
                ret = "/" + ret;
            }
            return ret;
        }
        internal bool PartExists(Uri partUri)
        {
            return Parts.ContainsKey(GetUriKey(partUri.OriginalString));
        }
        #endregion

        internal void DeletePart(Uri Uri)
        {
            _contentTypes.Remove(GetUriKey(Uri.OriginalString));
            Parts.Remove(GetUriKey(Uri.OriginalString));
        }
        internal MemoryStream Save()
        {
            var ms = new MemoryStream();
            var enc = Encoding.UTF8;
            ZipOutputStream os = new ZipOutputStream(ms, true);
            /**** ContentType****/
            var entry = os.PutNextEntry("[Content_Types].xml");
            byte[] b = enc.GetBytes(GetContentTypeXml());
            os.Write(b, 0, b.Length);
            /**** Top Rels ****/
            _rels.WriteZip(os, "_rels\\.rels");

            foreach (var part in Parts.Values)
            {
                part.WriteZip(os);
            }

            os.Flush();
            os.Close();
            return ms;
        }

        private string GetContentTypeXml()
        {
            StringBuilder xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            foreach (ContentType ct in _contentTypes.Values)
            {
                if (ct.IsExtension)
                {
                    xml.AppendFormat("<Default ContentType=\"{0}\" Extension=\"{1}\"/>", ct.Name, ct.Match);
                }
                else
                {
                    xml.AppendFormat("<Override ContentType=\"{0}\" PartName=\"{1}\" />", ct.Name, GetUriKey(ct.Match));
                }
            }
            xml.Append("</Types>");
            return xml.ToString();
        }
        internal void Flush()
        {

        }
        internal void Close()
        {

        }
    }
}
