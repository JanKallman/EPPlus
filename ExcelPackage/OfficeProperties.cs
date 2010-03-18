/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 * EPPlus is a fork of the ExcelPackage project
 * See http://www.codeplex.com/EPPlus for details.
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * The GNU General Public License can be viewed at http://www.opensource.org/licenses/gpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 * 
 * The code for this project may be used and redistributed by any means PROVIDING it is 
 * not sold for profit without the author's written consent, and providing that this notice 
 * and the author's name and all copyright notices remain intact.
 * 
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * 
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman                      Total rewrite               2010-03-01
 * *******************************************************************************/
using System;
using System.Xml;
using System.IO;
using System.IO.Packaging;
using System.Globalization;

namespace OfficeOpenXml
{
	/// <summary>
	/// Provides access to the properties bag of the package
	/// </summary>
	public class OfficeProperties : XmlHelper
	{
		#region Private Properties
		private Uri _uriPropertiesCore = new Uri("/docProps/core.xml", UriKind.Relative);
		private Uri _uriPropertiesExtended = new Uri("/docProps/app.xml", UriKind.Relative);
		private Uri _uriPropertiesCustom = new Uri("/docProps/custom.xml", UriKind.Relative);

		private XmlDocument _xmlPropertiesCore;
		private XmlDocument _xmlPropertiesExtended;
		private XmlDocument _xmlPropertiesCustom;

        XmlHelper _coreHelper;
        XmlHelper _extendedHelper;
        XmlHelper _customHelper;
        private ExcelPackage _package;
		#endregion 

		#region ExcelProperties Constructor
		/// <summary>
		/// Provides access to all the office document properties.
		/// </summary>
		/// <param name="xlPackage"></param>
		public OfficeProperties(ExcelPackage xlPackage, XmlNamespaceManager ns) : 
            base(ns)
		{
			_package = xlPackage;
            _coreHelper = new XmlHelper(ns, CorePropertiesXml.SelectSingleNode("cp:coreProperties", NameSpaceManager));             
            _extendedHelper = new XmlHelper(ns, ExtendedPropertiesXml);
            _customHelper = new XmlHelper(ns, CustomPropertiesXml);

		}
		#endregion

		#region Protected Internal Properties
		/// <summary>
		/// The URI to the core properties component (core.xml)
		/// </summary>
		protected internal Uri CorePropertiesUri { get { return (_uriPropertiesCore); } }
		/// <summary>
		/// The URI to the extended properties component (app.xml)
		/// </summary
		protected internal Uri ExtendedPropertiesUri { get { return (_uriPropertiesExtended); } }
		/// <summary>
		/// The URI to the custom properties component (custom.xml)
		/// </summary>
		protected internal Uri CustomPropertiesUri { get { return (_uriPropertiesCustom); } }
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
					if (_package.Package.PartExists(CorePropertiesUri))
						_xmlPropertiesCore = _package.GetXmlFromUri(CorePropertiesUri);
					else
					{
                        _xmlPropertiesCore = new XmlDocument();
                        _xmlPropertiesCore.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><cp:coreProperties xmlns:cp=\"{0}\" xmlns:dc=\"{1}\" xmlns:dcterms=\"{2}\" xmlns:dcmitype=\"{3}\" xmlns:xsi=\"{4}\"></cp:coreProperties>",
                            ExcelPackage.schemaCore,
                            ExcelPackage.schemaDc,
                            ExcelPackage.schemaDcTerms, 
                            ExcelPackage.schemaDcmiType,
                            ExcelPackage.schemaXsi));

                        // create a new document properties part and add to the package
						PackagePart partCore = _package.Package.CreatePart(CorePropertiesUri, @"application/vnd.openxmlformats-package.core-properties+xml");

						// create the document properties XML (with no entries in it)


						// save it to the package
						StreamWriter streamCore = new StreamWriter(partCore.GetStream(FileMode.Create, FileAccess.Write));
						_xmlPropertiesCore.Save(streamCore);
						streamCore.Close();
						_package.Package.Flush();

						// create the relationship between the workbook and the new shared strings part
						_package.Package.CreateRelationship(CorePropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
						_package.Package.Flush();
					}
				}
				return (_xmlPropertiesCore);
			}
		}
		#endregion
        #region "Core Properties"
        const string TitlePath = "dc:title";
        /// <summary>
		/// Gets/sets the title property of the document (core property)
		/// </summary>
		public string Title
		{
            get { return _coreHelper.GetXmlNode(TitlePath) ; }
            set { _coreHelper.SetXmlNode(TitlePath, value); }
		}

        const string SubjectPath = "dc:subject";
        /// <summary>
		/// Gets/sets the subject property of the document (core property)
		/// </summary>
        public string Subject
        {
            get { return _coreHelper.GetXmlNode(SubjectPath); }
            set { _coreHelper.SetXmlNode(SubjectPath, value); }
        }

        const string AuthorPath = "dc:creator";
        /// <summary>
		/// Gets/sets the author property of the document (core property)
		/// </summary>
		public string Author
		{
            get { return _coreHelper.GetXmlNode(AuthorPath); }
            set { _coreHelper.SetXmlNode(AuthorPath, value); }
		}

        const string CommentsPath = "dc:description";
        /// <summary>
		/// Gets/sets the comments property of the document (core property)
		/// </summary>
		public string Comments
		{
            get { return _coreHelper.GetXmlNode(CommentsPath); }
            set { _coreHelper.SetXmlNode(CommentsPath, value); }
        }

        const string KeywordsPath = "cp:keywords";
        /// <summary>
		/// Gets/sets the keywords property of the document (core property)
		/// </summary>
		public string Keywords
		{
            get { return _coreHelper.GetXmlNode(KeywordsPath); }
            set { _coreHelper.SetXmlNode(KeywordsPath, value); }
        }

        const string LastModifiedByPath = "cp:lastModifiedBy";
        /// <summary>
		/// Gets/sets the lastModifiedBy property of the document (core property)
		/// </summary>
		public string LastModifiedBy
		{
            get { return _coreHelper.GetXmlNode(LastModifiedByPath); }
            set { _coreHelper.SetXmlNode(LastModifiedByPath, value); }
        }

        const string LastPrintedPath = "cp:lastPrinted";
        /// <summary>
		/// Gets/sets the lastPrinted property of the document (core property)
		/// </summary>
		public string LastPrinted
		{
            get { return _coreHelper.GetXmlNode(LastPrintedPath); }
            set { _coreHelper.SetXmlNode(LastPrintedPath, value); }
        }

        const string CategoryPath = "cp:category";
        /// <summary>
		/// Gets/sets the category property of the document (core property)
		/// </summary>
		public string Category
		{
            get { return _coreHelper.GetXmlNode(CategoryPath); }
            set { _coreHelper.SetXmlNode(CategoryPath, value); }
        }

        const string ContentStatusPath = "cp:contentStatus";
        /// <summary>
		/// Gets/sets the status property of the document (core property)
		/// </summary>
		public string Status
		{
            get { return _coreHelper.GetXmlNode(ContentStatusPath); }
            set { _coreHelper.SetXmlNode(ContentStatusPath, value); }
        }
		
		#region Get and Set Core Properties
		/// <summary>
		/// Gets the value of a core property
		/// Private method, for internal use only!
		/// </summary>
		/// <param name="nameSpace">The namespace of the property</param>
		/// <param name="propertyName">The property name</param>
		/// <returns>The current value of the property</returns>
        //private string GetCorePropertyValue(string nameSpace, string propertyName)
        //{
        //    string retValue = null;
        //    string searchString = string.Format("//cp:coreProperties/{0}:{1}", nameSpace, propertyName);
        //    XmlNode node = CorePropertiesXml.SelectSingleNode(searchString, NameSpaceManager);
        //    if (node != null)
        //    {
        //        retValue = node.InnerText;
        //    }
        //    return retValue;
        //}

        ///// <summary>
        ///// Sets a core property value.
        ///// Private method, for internal use only!
        ///// </summary>
        ///// <param name="nameSpace">The property's namespace</param>
        ///// <param name="propertyName">The name of the property</param>
        ///// <param name="propValue">The value of the property</param>
        //private void SetCorePropertyValue(string nameSpace, string propertyName, string propValue)
        //{
        //    string searchString = string.Format("//cp:coreProperties/{0}:{1}", nameSpace, propertyName);
        //    XmlNode node = CorePropertiesXml.SelectSingleNode(searchString, NameSpaceManager);
        //    if (node == null)
        //    {
        //        // the property does not exist, so create the XML node
        //        string schema = ExcelPackage.schemaCore;
        //        switch (nameSpace)
        //        {
        //            case "cp": schema = ExcelPackage.schemaCore; break;
        //            case "dc": schema = ExcelPackage.schemaDc; break;
        //            case "dcterms": schema = ExcelPackage.schemaDcTerms; break;
        //            case "dcmitype": schema = ExcelPackage.schemaDcmiType; break;
        //            case "xsi": schema = ExcelPackage.schemaXsi; break;
        //        }
        //        node = (XmlNode) CorePropertiesXml.CreateElement(nameSpace, propertyName, schema);
        //        CorePropertiesXml.DocumentElement.AppendChild(node);
				
        //    }
        //    node.InnerText = propValue;
        //}
		#endregion

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
                    if (_package.Package.PartExists(ExtendedPropertiesUri))
                        _xmlPropertiesExtended = _package.GetXmlFromUri(ExtendedPropertiesUri);
                    else
                    {
                        _xmlPropertiesExtended = new XmlDocument();
                        _xmlPropertiesExtended.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
                            ExcelPackage.schemaVt,
                            ExcelPackage.schemaExtended));

                        // create a new document properties part and add to the package
                        PackagePart partExtended = _package.Package.CreatePart(ExtendedPropertiesUri, @"application/vnd.openxmlformats-officedocument.extended-properties+xml");

                        // save it to the package
                        StreamWriter streamExtended = new StreamWriter(partExtended.GetStream(FileMode.Create, FileAccess.Write));
                        _xmlPropertiesExtended.Save(streamExtended);
                        streamExtended.Close();
                        _package.Package.Flush();

                        // create the relationship between the workbook and the new shared strings part
                        _package.Package.CreateRelationship(ExtendedPropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
                        _package.Package.Flush();
                    }
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
			get { return _extendedHelper.GetXmlNode(ApplicationPath); }
		}

        const string HyperlinkBasePath = "xp:Properties/xp:HyperlinkBase";
        /// <summary>
		/// Gets/sets the HyperlinkBase property of the document (extended property)
		/// </summary>
		public Uri HyperlinkBase
		{
            get { return new Uri(_extendedHelper.GetXmlNode(HyperlinkBasePath), UriKind.Absolute); }
            set { _extendedHelper.SetXmlNode(HyperlinkBasePath, value.AbsoluteUri); }
		}

        const string AppVersionPath = "xp:Properties/xp:AppVersion";
        /// <summary>
		/// Gets the AppVersion property of the document (extended property)
		/// </summary>
		public string AppVersion
		{
            get { return _extendedHelper.GetXmlNode(AppVersionPath); }
        }
        const string CompanyPath = "xp:Properties/xp:Company";

		/// <summary>
		/// Gets/sets the Company property of the document (extended property)
		/// </summary>
		public string Company
		{
            get { return _extendedHelper.GetXmlNode(CompanyPath); }
            set { _extendedHelper.SetXmlNode(CompanyPath, value); }
        }

        const string ManagerPath = "xp:Properties/xp:Manager";
		/// <summary>
		/// Gets/sets the Manager property of the document (extended property)
		/// </summary>
		public string Manager
		{
            get { return _extendedHelper.GetXmlNode(ManagerPath); }
            set { _extendedHelper.SetXmlNode(ManagerPath, value); }
		}

		#region Get and Set Extended Properties
		private string GetExtendedPropertyValue(string propertyName)
		{
			string retValue = null;
			string searchString = string.Format("//xp:Properties/xp:{0}", propertyName);
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
                    if (_package.Package.PartExists(CustomPropertiesUri))
                        _xmlPropertiesCustom = _package.GetXmlFromUri(CustomPropertiesUri);
                    else
                    {
                        _xmlPropertiesCustom = new XmlDocument();
                        _xmlPropertiesCustom.LoadXml(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\" ?><Properties xmlns:vt=\"{0}\" xmlns=\"{1}\"></Properties>",
                            ExcelPackage.schemaVt,
                            ExcelPackage.schemaCustom));

                        // create a new document properties part and add to the package
                        PackagePart partCustom = _package.Package.CreatePart(CustomPropertiesUri, @"application/vnd.openxmlformats-officedocument.custom-properties+xml");

                        // save it to the package
                        StreamWriter streamCustom = new StreamWriter(partCustom.GetStream(FileMode.Create, FileAccess.Write));
                        _xmlPropertiesCustom.Save(streamCustom);
                        streamCustom.Close();
                        _package.Package.Flush();

                        // create the relationship between the workbook and the new shared strings part
                        _package.Package.CreateRelationship(CustomPropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
                        _package.Package.Flush();
                    }
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
            string searchString = string.Format("//ctp:Properties/ctp:property[@name='{0}']", propertyName);
            XmlElement node = CustomPropertiesXml.SelectSingleNode(searchString, NameSpaceManager) as XmlElement;
            if (node != null)
            {
                string value=node.LastChild.InnerText;
                switch (node.LastChild.LocalName)
                {
                    case "filetime":
                        DateTime dt;
                        if(DateTime.TryParse(value, out dt))
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
                        if (double.TryParse(value, NumberStyles.Any,ExcelWorksheet._ci, out d))
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
		/// Allows you to set the value of a current custom property or create 
		/// your own custom property.  
		/// Currently only supports string values.
		/// </summary>
		/// <param name="propertyName">The name of the property</param>
		/// <param name="propValue">The value of the property</param>
		public void SetCustomPropertyValue(string propertyName, object value)
		{
			// TODO:  provide support for other custom property data types
			string searchString = @"ctp:Properties";
            XmlNode allProps = CustomPropertiesXml.SelectSingleNode(searchString, NameSpaceManager);

            searchString = string.Format("//ctp:Properties/ctp:property[@name='{0}']", propertyName);
			XmlElement node = CustomPropertiesXml.SelectSingleNode(searchString, NameSpaceManager) as XmlElement;
            if (node == null)
            {
                int pid;
                var MaxNode = CustomPropertiesXml.SelectSingleNode("//ctp:Properties/ctp:property[not(@pid <= preceding-sibling::ctp:property/@pid) and not(@pid <= following-sibling::ctp:property/@pid)]", NameSpaceManager);
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
                while(node.ChildNodes.Count > 0) node.RemoveChild(node.ChildNodes[0]);
            }
            XmlElement valueElem;
            if (value is bool)
            {
                valueElem = CustomPropertiesXml.CreateElement("vt", "bool", ExcelPackage.schemaVt);
                valueElem.InnerText = value.ToString().ToLower();
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
                if(value is double)
                {
                    valueElem.InnerText = ((double)value).ToString(ExcelWorksheet._ci);
                }
                else if (value is float)
                {
                    valueElem.InnerText = ((float)value).ToString(ExcelWorksheet._ci);
                }
                else if (value is decimal)
                {
                    valueElem.InnerText = ((decimal)value).ToString(ExcelWorksheet._ci);
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

		#region Save  // OfficeProperties save
		/// <summary>
		/// Saves the office document properties back to the package.
		/// </summary>
		protected internal void Save()
		{
			if (_xmlPropertiesCore != null)
			{
				_package.WriteDebugFile(_xmlPropertiesCore, "docProps", "core.xml");
				_package.SavePart(CorePropertiesUri, _xmlPropertiesCore);
			}
			if (_xmlPropertiesExtended != null)
			{
				_package.WriteDebugFile(_xmlPropertiesExtended, "docProps", "app.xml");
				_package.SavePart(ExtendedPropertiesUri, _xmlPropertiesExtended);
			}
			if (_xmlPropertiesCustom != null)
			{
				_package.WriteDebugFile(_xmlPropertiesCustom, "docProps", "custom.xml");
				_package.SavePart(CustomPropertiesUri, _xmlPropertiesCustom);
			}

		}
		#endregion

	}
}
