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
using Ionic.Zip;
using System.IO;
using System.Xml;
using Ionic.Zlib;
using System.Web;
namespace OfficeOpenXml.Packaging
{
    public abstract class ZipPackageRelationshipBase
    {
        protected ZipPackageRelationshipCollection _rels = new ZipPackageRelationshipCollection();
        protected internal 
        int maxRId = 1;
        internal void DeleteRelationship(string id)
        {
            _rels.Remove(id);
            UpdateMaxRId(id, ref maxRId);
        }
        protected void UpdateMaxRId(string id, ref int maxRId)
        {
            if (id.StartsWith("rId"))
            {
                int num;
                if (int.TryParse(id.Substring(3), out num))
                {
                    if (num == maxRId - 1)
                    {
                        maxRId--;
                    }
                }
            }
        }
        internal virtual ZipPackageRelationship CreateRelationship(Uri targetUri, TargetMode targetMode, string relationshipType)
        {
            var rel = new ZipPackageRelationship();
            rel.TargetUri = targetUri;
            rel.TargetMode = targetMode;
            rel.RelationshipType = relationshipType;
            rel.Id = "RId" + (maxRId++).ToString();
            _rels.Add(rel);
            return rel;
        }
        internal bool RelationshipExists(string id)
        {
            return _rels.ContainsKey(id);
        }
        internal ZipPackageRelationshipCollection GetRelationshipsByType(string schema)
        {
            return _rels.GetRelationshipsByType(schema);
        }
        internal ZipPackageRelationshipCollection GetRelationships()
        {
            return _rels;
        }
        internal ZipPackageRelationship GetRelationship(string id)
        {
            return _rels[id];
        }
        internal void ReadRelation(string xml, string source)
        {
            var doc = new XmlDocument();
            XmlHelper.LoadXmlSafe(doc, xml, Encoding.UTF8);

            foreach (XmlElement c in doc.DocumentElement.ChildNodes)
            {
                var rel = new ZipPackageRelationship();
                rel.Id = c.GetAttribute("Id");
                rel.RelationshipType = c.GetAttribute("Type");
                rel.TargetMode = c.GetAttribute("TargetMode").ToLower() == "external" ? TargetMode.External : TargetMode.Internal;
                try
                {
                    rel.TargetUri = new Uri(c.GetAttribute("Target"), UriKind.RelativeOrAbsolute);
                }
                catch
                {
                    //The URI is not a valid URI. Encode it to make i valid.
                    rel.TargetUri = new Uri(HttpUtility.UrlEncode("Invalid:URI "+c.GetAttribute("Target")), UriKind.RelativeOrAbsolute);
                }
                if (!string.IsNullOrEmpty(source))
                {
                    rel.SourceUri = new Uri(source, UriKind.Relative);
                }
                if (rel.Id.ToLower().StartsWith("rid"))
                {
                    int id;
                    if (int.TryParse(rel.Id.Substring(3), out id))
                    {
                        if (id >= maxRId && id < int.MaxValue - 10000) //Not likly to have this high id's but make sure we have space to avoid overflow.
                        {
                            maxRId = id + 1;
                        }
                    }
                }
                _rels.Add(rel);
            }
        }
    }
}