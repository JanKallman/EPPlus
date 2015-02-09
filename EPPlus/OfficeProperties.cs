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
 * Jan Källman                      Total rewrite               2010-03-01
 * Jan Källman		    License changed GPL-->LGPL  2011-12-27
 * Raziq York                       Added Created & Modified    2014-08-20
 *******************************************************************************/
using System;
using System.Xml;
using System.IO;
using System.Globalization;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml
{
    /// <summary>
    /// Provides access to the properties bag of the package
    /// </summary>
    public sealed class OfficeProperties : XmlHelper
    {
        #region Private Properties
        private XmlDocument _xmlPropertiesCore;
        private XmlDocument _xmlPropertiesExtended;
        private XmlDocument _xmlPropertiesCustom;

        private Uri _uriPropertiesCore = new Uri("/docProps/core.xml", UriKind.Relative);
        private Uri _uriPropertiesExtended = new Uri("/docProps/app.xml", UriKind.Relative);
        private Uri _uriPropertiesCustom = new Uri("/docProps/custom.xml", UriKind.Relative);

        XmlHelper _coreHelper;
        XmlHelper _extendedHelper;
        XmlHelper _customHelper;
        private ExcelPackage _package;
        #endregion

        #region ExcelProperties Constructor
        /// <summary>
        /// Provides access to all the office document properties.
        /// </summary>
        /// <param name="package"></param>
        /// <param name="ns"></param>
        internal OfficeProperties(ExcelPackage package, XmlNamespaceManager ns) :
            base(ns)
        {
            _package = package;

            _coreHelper = XmlHelperFactory.Create(ns, CorePropertiesXml.SelectSingleNode("cp:coreProperties", NameSpaceManager));
            _extendedHelper = XmlHelperFactory.Create(ns, ExtendedPropertiesXml);
            _customHelper = XmlHelperFactory.Create(ns, CustomPropertiesXml);

        }
        #endregion
        #region CorePropertiesXml
        /// <summary>
        /// Provides access to the XML document that holds all the code 
        /// document properties.
        /// </summary>
        public XmlDocument CorePropertiesXml
        {
            get
            {
                if (_xmlPropertiesCore == null)
                {
                    string xml = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><cp:coreProperties xmlns:cp=\"{0}\" xmlns:dc=\"{1}\" xmlns:dcterms=\"{2}\" xmlns:dcmitype=\"{3}\" xmlns:xsi=\"{4}\"></cp:coreProperties>",
                        ExcelPackage.schemaCore,
                        ExcelPackage.schemaDc,
                        ExcelPackage.schemaDcTerms,
                        ExcelPackage.schemaDcmiType,
                        ExcelPackage.schemaXsi);

                    _xmlPropertiesCore = GetXmlDocument(xml, _uriPropertiesCore, @"application/vnd.openxmlformats-package.core-properties+xml", @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
                }
                return (_xmlPropertiesCore);
            }
        }

        private XmlDocument GetXmlDocument(string startXml, Uri uri, string contentType, string relationship)
        {
            XmlDocument xmlDoc;
            if (_package.Package.PartExists(uri))
                xmlDoc = _package.GetXmlFromUri(uri);
            else
            {
                xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(startXml);

                // Create a the part and add to the package
                Packaging.ZipPackagePart part = _package.Package.CreatePart(uri, contentType);

                // Save it to the package
                StreamWriter stream = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                xmlDoc.Save(stream);
                //stream.Close();
                _package.Package.Flush();

                // create the relationship between the workbook and the new shared strings part
                _package.Package.CreateRelationship(UriHelper.GetRelativeUri(new Uri("/xl", UriKind.Relative), uri), Packaging.TargetMode.Internal, relationship);
                _package.Package.Flush();
            }
            return xmlDoc;
        }
        #endregion
        #region Core Properties
        const string TitlePath = "dc:title";
        /// <summary>
        /// Gets/sets the title property of the document (core property)
        /// </summary>
        public string Title
        {
            get { return _coreHelper.GetXmlNodeString(TitlePath); }
            set { _coreHelper.SetXmlNodeString(TitlePath, value); }
        }

        const string SubjectPath = "dc:subject";
        /// <summary>
        /// Gets/sets the subject property of the document (core property)
        /// </summary>
        public string Subject
        {
            get { return _coreHelper.GetXmlNodeString(SubjectPath); }
            set { _coreHelper.SetXmlNodeString(SubjectPath, value); }
        }

        const string AuthorPath = "dc:creator";
        /// <summary>
        /// Gets/sets the author property of the document (core property)
        /// </summary>
        public string Author
        {
            get { return _coreHelper.GetXmlNodeString(AuthorPath); }
            set { _coreHelper.SetXmlNodeString(AuthorPath, value); }
        }

        const string CommentsPath = "dc:description";
        /// <summary>
        /// Gets/sets the comments property of the document (core property)
        /// </summary>
        public string Comments
        {
            get { return _coreHelper.GetXmlNodeString(CommentsPath); }
            set { _coreHelper.SetXmlNodeString(CommentsPath, value); }
        }

        const string KeywordsPath = "cp:keywords";
        /// <summary>
        /// Gets/sets the keywords property of the document (core property)
        /// </summary>
        public string Keywords
        {
            get { return _coreHelper.GetXmlNodeString(KeywordsPath); }
            set { _coreHelper.SetXmlNodeString(KeywordsPath, value); }
        }

        const string LastModifiedByPath = "cp:lastModifiedBy";
        /// <summary>
        /// Gets/sets the lastModifiedBy property of the document (core property)
        /// </summary>
        public string LastModifiedBy
        {
            get { return _coreHelper.GetXmlNodeString(LastModifiedByPath); }
            set { _coreHelper.SetXmlNodeString(LastModifiedByPath, value); }
        }

        const string LastPrintedPath = "cp:lastPrinted";
        /// <summary>
        /// Gets/sets the lastPrinted property of the document (core property)
        /// </summary>
        public string LastPrinted
        {
            get { return _coreHelper.GetXmlNodeString(LastPrintedPath); }
            set { _coreHelper.SetXmlNodeString(LastPrintedPath, value); }
        }

        const string CreatedPath = "dcterms:created";

        /// <summary>
	    /// Gets/sets the created property of the document (core property)
	    /// </summary>
	    public DateTime Created
	    {
	        get
	        {
	            DateTime date;
	            return DateTime.TryParse(_coreHelper.GetXmlNodeString(CreatedPath), out date) ? date : DateTime.MinValue;
	        }
	        set
	        {
	            var dateString = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
	            _coreHelper.SetXmlNodeString(CreatedPath, dateString);
	        }
	    }

        const string CategoryPath = "cp:category";
        /// <summary>
        /// Gets/sets the category property of the document (core property)
        /// </summary>
        public string Category
        {
            get { return _coreHelper.GetXmlNodeString(CategoryPath); }
            set { _coreHelper.SetXmlNodeString(CategoryPath, value); }
        }

        const string ContentStatusPath = "cp:contentStatus";
        /// <summary>
        /// Gets/sets the status property of the document (core property)
        /// </summary>
        public string Status
        {
            get { return _coreHelper.GetXmlNodeString(ContentStatusPath); }
            set { _coreHelper.SetXmlNodeString(ContentStatusPath, value); }
        }
        #endregion

        #region Extended Properties
        #region ExtendedPropertiesXml
        /// <summary>
        /// Provides access to the XML document that holds the extended properties of the document (app.xml)
        /// </summary>
        public XmlDocument ExtendedPropertiesXml
        {
            get
            {
                if (_xmlPropertiesExtended == null)
                {
                    _xmlPropertiesExtended = GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
                            ExcelPackage.schemaVt,
                            ExcelPackage.schemaExtended),
                        _uriPropertiesExtended,
                        @"application/vnd.openxmlformats-officedocument.extended-properties+xml",
                        @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
                }
                return (_xmlPropertiesExtended);
            }
        }
        #endregion

        const string ApplicationPath = "xp:Properties/xp:Application";
        /// <summary>
        /// Gets the Application property of the document (extended property)
        /// </summary>
        public string Application
        {
            get { return _extendedHelper.GetXmlNodeString(ApplicationPath); }
        }

        const string HyperlinkBasePath = "xp:Properties/xp:HyperlinkBase";
        /// <summary>
        /// Gets/sets the HyperlinkBase property of the document (extended property)
        /// </summary>
        public Uri HyperlinkBase
        {
            get { return new Uri(_extendedHelper.GetXmlNodeString(HyperlinkBasePath), UriKind.Absolute); }
            set { _extendedHelper.SetXmlNodeString(HyperlinkBasePath, value.AbsoluteUri); }
        }

        const string AppVersionPath = "xp:Properties/xp:AppVersion";
        /// <summary>
        /// Gets the AppVersion property of the document (extended property)
        /// </summary>
        public string AppVersion
        {
            get { return _extendedHelper.GetXmlNodeString(AppVersionPath); }
        }
        const string CompanyPath = "xp:Properties/xp:Company";

        /// <summary>
        /// Gets/sets the Company property of the document (extended property)
        /// </summary>
        public string Company
        {
            get { return _extendedHelper.GetXmlNodeString(CompanyPath); }
            set { _extendedHelper.SetXmlNodeString(CompanyPath, value); }
        }

        const string ManagerPath = "xp:Properties/xp:Manager";
        /// <summary>
        /// Gets/sets the Manager property of the document (extended property)
        /// </summary>
        public string Manager
        {
            get { return _extendedHelper.GetXmlNodeString(ManagerPath); }
            set { _extendedHelper.SetXmlNodeString(ManagerPath, value); }
        }

        const string ModifiedPath = "dcterms:modified";
	    /// <summary>
	    /// Gets/sets the modified property of the document (core property)
	    /// </summary>
	    public DateTime Modified
	    {
	        get
	        {
	            DateTime date;
	            return DateTime.TryParse(_coreHelper.GetXmlNodeString(ModifiedPath), out date) ? date : DateTime.MinValue;
	        }
	        set
	        {
	            var dateString = value.ToUniversalTime().ToString("s", CultureInfo.InvariantCulture) + "Z";
	            _coreHelper.SetXmlNodeString(ModifiedPath, dateString);
	        }
	    }

        #region Get and Set Extended Properties
        private string GetExtendedPropertyValue(string propertyName)
        {
            string retValue = null;
            string searchString = string.Format("xp:Properties/xp:{0}", propertyName);
            XmlNode node = ExtendedPropertiesXml.SelectSingleNode(searchString, NameSpaceManager);
            if (node != null)
            {
                retValue = node.InnerText;
            }
            return retValue;
        }
        #endregion
        #endregion

        #region Custom Properties

        #region CustomPropertiesXml
        /// <summary>
        /// Provides access to the XML document which holds the document's custom properties
        /// </summary>
        public XmlDocument CustomPropertiesXml
        {
            get
            {
                if (_xmlPropertiesCustom == null)
                {
                    _xmlPropertiesCustom = GetXmlDocument(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
                            ExcelPackage.schemaVt,
                            ExcelPackage.schemaCustom),
                         _uriPropertiesCustom, 
                         @"application/vnd.openxmlformats-officedocument.custom-properties+xml",
                         @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
                }
                return (_xmlPropertiesCustom);
            }
        }
        #endregion

        #region Get and Set Custom Properties
        /// <summary>
        /// Gets the value of a custom property
        /// </summary>
        /// <param name="propertyName">The name of the property</param>
        /// <returns>The current value of the property</returns>
        public object GetCustomPropertyValue(string propertyName)
        {
            string searchString = string.Format("ctp:Properties/ctp:property[@name='{0}']", propertyName);
            XmlElement node = CustomPropertiesXml.SelectSingleNode(searchString, NameSpaceManager) as XmlElement;
            if (node != null)
            {
                string value = node.LastChild.InnerText;
                switch (node.LastChild.LocalName)
                {
                    case "filetime":
                        DateTime dt;
                        if (DateTime.TryParse(value, out dt))
                        {
                            return dt;
                        }
                        else
                        {
                            return null;
                        }
                    case "i4":
                        int i;
                        if (int.TryParse(value, out i))
                        {
                            return i;
                        }
                        else
                        {
                            return null;
                        }
                    case "r8":
                        double d;
                        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                        {
                            return d;
                        }
                        else
                        {
                            return null;
                        }
                    case "bool":
                        if (value == "true")
                        {
                            return true;
                        }
                        else if (value == "false")
                        {
                            return false;
                        }
                        else
                        {
                            return null;
                        }
                    default:
                        return value;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Allows you to set the value of a current custom property or create your own custom property.  
        /// </summary>
        /// <param name="propertyName">The name of the property</param>
        /// <param name="value">The value of the property</param>
        public void SetCustomPropertyValue(string propertyName, object value)
        {
            XmlNode allProps = CustomPropertiesXml.SelectSingleNode(@"ctp:Properties", NameSpaceManager);

            var prop = string.Format("ctp:Properties/ctp:property[@name='{0}']", propertyName);
            XmlElement node = CustomPropertiesXml.SelectSingleNode(prop, NameSpaceManager) as XmlElement;
            if (node == null)
            {
                int pid;
                var MaxNode = CustomPropertiesXml.SelectSingleNode("ctp:Properties/ctp:property[not(@pid <= preceding-sibling::ctp:property/@pid) and not(@pid <= following-sibling::ctp:property/@pid)]", NameSpaceManager);
                if (MaxNode == null)
                {
                    pid = 2;
                }
                else
                {
                    if (!int.TryParse(MaxNode.Attributes["pid"].Value, out pid))
                    {
                        pid = 2;
                    }
                    pid++;
                }
                node = CustomPropertiesXml.CreateElement("property", ExcelPackage.schemaCustom);
                node.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
                node.SetAttribute("pid", pid.ToString());  // custom property pid
                node.SetAttribute("name", propertyName);

                allProps.AppendChild(node);
            }
            else
            {
                while (node.ChildNodes.Count > 0) node.RemoveChild(node.ChildNodes[0]);
            }
            XmlElement valueElem;
            if (value is bool)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "bool", ExcelPackage.schemaVt);
                valueElem.InnerText = value.ToString().ToLower(CultureInfo.InvariantCulture);
            }
            else if (value is DateTime)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "filetime", ExcelPackage.schemaVt);
                valueElem.InnerText = ((DateTime)value).AddHours(-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
            }
            else if (value is short || value is int)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "i4", ExcelPackage.schemaVt);
                valueElem.InnerText = value.ToString();
            }
            else if (value is double || value is decimal || value is float || value is long)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "r8", ExcelPackage.schemaVt);
                if (value is double)
                {
                    valueElem.InnerText = ((double)value).ToString(CultureInfo.InvariantCulture);
                }
                else if (value is float)
                {
                    valueElem.InnerText = ((float)value).ToString(CultureInfo.InvariantCulture);
                }
                else if (value is decimal)
                {
                    valueElem.InnerText = ((decimal)value).ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    valueElem.InnerText = value.ToString();
                }
            }
            else
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "lpwstr", ExcelPackage.schemaVt);
                valueElem.InnerText = value.ToString();
            }
            node.AppendChild(valueElem);
        }
        #endregion
        #endregion

        #region Save
        /// <summary>
        /// Saves the document properties back to the package.
        /// </summary>
        internal void Save()
        {
            if (_xmlPropertiesCore != null)
            {
                _package.SavePart(_uriPropertiesCore, _xmlPropertiesCore);
            }
            if (_xmlPropertiesExtended != null)
            {
                _package.SavePart(_uriPropertiesExtended, _xmlPropertiesExtended);
            }
            if (_xmlPropertiesCustom != null)
            {
                _package.SavePart(_uriPropertiesCustom, _xmlPropertiesCustom);
            }

        }
        #endregion

    }
}
