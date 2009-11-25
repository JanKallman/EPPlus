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
 *******************************************************************************/

/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * ExcelPackage provides server-side generation of Excel 2007 spreadsheets.
 * See http://www.codeplex.com/ExcelPackage for details.
 * 
 * Copyright 2007 © Dr John Tunnicliffe 
 * mailto:dr.john.tunnicliffe@btinternet.com
 * All rights reserved.
 * 
 * ExcelPackage is an Open Source project provided under the 
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
 */

/*
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * John Tunnicliffe		Initial Release		01-Jan-2007
 * ******************************************************************************
 */
using System;
using System.Xml;
using System.IO;
using System.IO.Packaging;

namespace OfficeOpenXml
{
	/// <summary>
	/// Provides access to the properties bag of any office document (i.e. Word, Excel etc.)
	/// </summary>
	public class OfficeProperties
	{
		#region Private Properties
		
		private const string schemaCore = @"http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
		private const string schemeExtended = @"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
		private const string schemaCustom = @"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
		private const string schemaDc = @"http://purl.org/dc/elements/1.1/";
		private const string schemaDcTerms = @"http://purl.org/dc/terms/";
		private const string schemaDcmiType = @"http://purl.org/dc/dcmitype/";
		private const string schemaXsi = @"http://www.w3.org/2001/XMLSchema-instance";
		private const string schemaVt = @"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

		private Uri _uriPropertiesCore = new Uri("/docProps/core.xml", UriKind.Relative);
		private Uri _uriPropertiesExtended = new Uri("/docProps/app.xml", UriKind.Relative);
		private Uri _uriPropertiesCustom = new Uri("/docProps/custom.xml", UriKind.Relative);

		private XmlDocument _xmlPropertiesCore;
		private XmlDocument _xmlPropertiesExtended;
		private XmlDocument _xmlPropertiesCustom;
		private ExcelPackage _xlPackage;
		private XmlNamespaceManager _nsManager;
		#endregion 

		#region ExcelProperties Constructor
		/// <summary>
		/// Provides access to all the office document properties.
		/// </summary>
		/// <param name="xlPackage"></param>
		public OfficeProperties(ExcelPackage xlPackage)
		{
			_xlPackage = xlPackage;
			//  Create a NamespaceManager to handle the default namespace, 
			//  and create a prefix for the default namespace:
			NameTable nt = new NameTable();
			_nsManager = new XmlNamespaceManager(nt);
			// default namespace
			_nsManager.AddNamespace("d", ExcelPackage.schemaMain);
			_nsManager.AddNamespace("vt", schemaVt);
			// extended properties (app.xml)
			_nsManager.AddNamespace("xp", schemeExtended);
			// custom properties
			_nsManager.AddNamespace("ctp", schemaCustom);
			// core properties
			_nsManager.AddNamespace("cp", schemaCore);
			// core property namespaces
			_nsManager.AddNamespace("dc", schemaDc);
			_nsManager.AddNamespace("dcterms", schemaDcTerms);
			_nsManager.AddNamespace("dcmitype", schemaDcmiType);
			_nsManager.AddNamespace("xsi", schemaXsi);
		}
		#endregion

		#region Protected Internal Properties
		/// <summary>
		/// The URI to the core properties component (core.xml)
		/// </summary>
		protected internal Uri CorePropertiesUri { get { return (_uriPropertiesCore); } }
		/// <summary>
		/// The URI to the extended properties component (app.xml)
		/// </summary>
		protected internal Uri ExtendedPropertiesUri { get { return (_uriPropertiesExtended); } }
		/// <summary>
		/// The URI to the custom properties component (custom.xml)
		/// </summary>
		protected internal Uri CustomPropertiesUri { get { return (_uriPropertiesCustom); } }
		#endregion

		#region Core Properties

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
					if (_xlPackage.Package.PartExists(CorePropertiesUri))
						_xmlPropertiesCore = _xlPackage.GetXmlFromUri(CorePropertiesUri);
					else
					{
						// create a new document properties part and add to the package
						PackagePart partCore = _xlPackage.Package.CreatePart(CorePropertiesUri, @"application/vnd.openxmlformats-package.core-properties+xml");

						// create the document properties XML (with no entries in it)
						_xmlPropertiesCore = new XmlDocument();
						XmlElement root = _xmlPropertiesCore.CreateElement("cp:coreProperties", schemaCore);
						ExcelPackage.AddSchemaAttribute(root, schemaCore, "cp");
						ExcelPackage.AddSchemaAttribute(root, schemaDc, "dc");
						ExcelPackage.AddSchemaAttribute(root, schemaDcTerms, "dcterms");
						ExcelPackage.AddSchemaAttribute(root, schemaDcmiType, "dcmitype");
						ExcelPackage.AddSchemaAttribute(root, schemaXsi, "xsi");
						_xmlPropertiesCore.AppendChild(root);

						// save it to the package
						StreamWriter streamCore = new StreamWriter(partCore.GetStream(FileMode.Create, FileAccess.Write));
						_xmlPropertiesCore.Save(streamCore);
						streamCore.Close();
						_xlPackage.Package.Flush();

						// create the relationship between the workbook and the new shared strings part
						_xlPackage.Package.CreateRelationship(CorePropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
						_xlPackage.Package.Flush();
					}
				}
				return (_xmlPropertiesCore);
			}
		}
		#endregion

		/// <summary>
		/// Gets/sets the title property of the document (core property)
		/// </summary>
		public string Title
		{
			get	{	return GetCorePropertyValue("dc", "title"); }
			set { SetCorePropertyValue("dc", "title", value); }
		}

		/// <summary>
		/// Gets/sets the subject property of the document (core property)
		/// </summary>
		public string Subject
		{
			get { return GetCorePropertyValue("dc", "subject"); }
			set { SetCorePropertyValue("dc", "subject", value); }
		}

		/// <summary>
		/// Gets/sets the author property of the document (core property)
		/// </summary>
		public string Author
		{
			get { return GetCorePropertyValue("dc", "creator"); }
			set { SetCorePropertyValue("dc", "creator", value); }
		}

		/// <summary>
		/// Gets/sets the comments property of the document (core property)
		/// </summary>
		public string Comments
		{
			get { return GetCorePropertyValue("dc", "description"); }
			set { SetCorePropertyValue("dc", "description", value); }
		}

		/// <summary>
		/// Gets/sets the keywords property of the document (core property)
		/// </summary>
		public string Keywords
		{
			get { return GetCorePropertyValue("cp", "keywords"); }
			set { SetCorePropertyValue("cp", "keywords", value); }
		}

		/// <summary>
		/// Gets/sets the lastModifiedBy property of the document (core property)
		/// </summary>
		public string LastModifiedBy
		{
			get { return GetCorePropertyValue("cp", "lastModifiedBy"); }
			set { SetCorePropertyValue("cp", "lastModifiedBy", value); }
		}

		/// <summary>
		/// Gets/sets the lastPrinted property of the document (core property)
		/// </summary>
		public string LastPrinted
		{
			get { return GetCorePropertyValue("cp", "lastPrinted"); }
			set { SetCorePropertyValue("cp", "lastPrinted", value); }
		}

		/// <summary>
		/// Gets/sets the category property of the document (core property)
		/// </summary>
		public string Category
		{
			get { return GetCorePropertyValue("cp", "category"); }
			set { SetCorePropertyValue("cp", "category", value); }
		}

		/// <summary>
		/// Gets/sets the status property of the document (core property)
		/// </summary>
		public string Status
		{
			get { return GetCorePropertyValue("cp", "contentStatus"); }
			set { SetCorePropertyValue("cp", "contentStatus", value); }
		}
		
		#region Get and Set Core Properties
		/// <summary>
		/// Gets the value of a core property
		/// Private method, for internal use only!
		/// </summary>
		/// <param name="nameSpace">The namespace of the property</param>
		/// <param name="propertyName">The property name</param>
		/// <returns>The current value of the property</returns>
		private string GetCorePropertyValue(string nameSpace, string propertyName)
		{
			string retValue = null;
			string searchString = string.Format("//cp:coreProperties/{0}:{1}", nameSpace, propertyName);
			XmlNode node = CorePropertiesXml.SelectSingleNode(searchString, _nsManager);
			if (node != null)
			{
				retValue = node.InnerText;
			}
			return retValue;
		}

		/// <summary>
		/// Sets a core property value.
		/// Private method, for internal use only!
		/// </summary>
		/// <param name="nameSpace">The property's namespace</param>
		/// <param name="propertyName">The name of the property</param>
		/// <param name="propValue">The value of the property</param>
		private void SetCorePropertyValue(string nameSpace, string propertyName, string propValue)
		{
			string searchString = string.Format("//cp:coreProperties/{0}:{1}", nameSpace, propertyName);
			XmlNode node = CorePropertiesXml.SelectSingleNode(searchString, _nsManager);
			if (node == null)
			{
				// the property does not exist, so create the XML node
				string schema = schemaCore;
				switch (nameSpace)
				{
					case "cp": schema = schemaCore;	break;
					case "dc": schema = schemaDc;	break;
					case "dcterms": schema = schemaDcTerms;	break;
					case "dcmitype": schema = schemaDcmiType;	break;
					case "xsi": schema = schemaXsi;	break;
				}
				node = (XmlNode) CorePropertiesXml.CreateElement(nameSpace, propertyName, schema);
				CorePropertiesXml.DocumentElement.AppendChild(node);
				
			}
			node.InnerText = propValue;
		}
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
					if (_xlPackage.Package.PartExists(ExtendedPropertiesUri))
						_xmlPropertiesExtended = _xlPackage.GetXmlFromUri(ExtendedPropertiesUri);
					else
					{
						// create a new extended properties part and add to the package
						PackagePart partExtended = _xlPackage.Package.CreatePart(ExtendedPropertiesUri, @"application/vnd.openxmlformats-officedocument.extended-properties+xml");

						// create the extended properties XML (with no entries in it)
						_xmlPropertiesExtended = new XmlDocument();
						XmlElement root = _xmlPropertiesExtended.CreateElement("Properties", schemeExtended);
						ExcelPackage.AddSchemaAttribute(root, schemaVt, "vt");
						_xmlPropertiesExtended.AppendChild(root);

						// save it to the package
						StreamWriter streamExtended = new StreamWriter(partExtended.GetStream(FileMode.Create, FileAccess.Write));
						_xmlPropertiesExtended.Save(streamExtended);
						streamExtended.Close();
						_xlPackage.Package.Flush();

						// create the relationship between the workbook and the new shared strings part
						_xlPackage.Package.CreateRelationship(ExtendedPropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
						_xlPackage.Package.Flush();
					}
				}
				return (_xmlPropertiesExtended);
			}
		}
		#endregion

		/// <summary>
		/// Gets the Application property of the document (extended property)
		/// </summary>
		public string Application
		{
			get { return GetExtendedPropertyValue("Application"); }
		}

		/// <summary>
		/// Gets/sets the HyperlinkBase property of the document (extended property)
		/// </summary>
		public Uri HyperlinkBase
		{
			get { return new Uri(GetExtendedPropertyValue("HyperlinkBase")); }
			set	{	SetExtendedPropertyValue("HyperlinkBase", value.ToString());	}
		}

		/// <summary>
		/// Gets the AppVersion property of the document (extended property)
		/// </summary>
		public string AppVersion
		{
			get { return GetExtendedPropertyValue("AppVersion"); }
		}

		/// <summary>
		/// Gets/sets the Company property of the document (extended property)
		/// </summary>
		public string Company
		{
			get { return GetExtendedPropertyValue("Company"); }
			set { SetExtendedPropertyValue("Company", value); }
		}

		/// <summary>
		/// Gets/sets the Manager property of the document (extended property)
		/// </summary>
		public string Manager
		{
			get { return GetExtendedPropertyValue("Manager"); }
			set { SetExtendedPropertyValue("Manager", value); }
		}

		#region Get and Set Extended Properties
		private string GetExtendedPropertyValue(string propertyName)
		{
			string retValue = null;
			string searchString = string.Format("//xp:Properties/xp:{0}", propertyName);
			XmlNode node = ExtendedPropertiesXml.SelectSingleNode(searchString, _nsManager);
			if (node != null)
			{
				retValue = node.InnerText;
			}
			return retValue;
		}

		private void SetExtendedPropertyValue(string propertyName, string propValue)
		{
			string searchString = string.Format("//xp:Properties/xp:{0}", propertyName);
			XmlNode node = ExtendedPropertiesXml.SelectSingleNode(searchString, _nsManager);
			if (node == null)
			{
				// the property does not exist, so create the XML node
				node = (XmlNode)ExtendedPropertiesXml.CreateElement(propertyName, schemeExtended);
				ExtendedPropertiesXml.DocumentElement.AppendChild(node);
			}
			node.InnerText = propValue;
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
					if (_xlPackage.Package.PartExists(CustomPropertiesUri))
						_xmlPropertiesCustom = _xlPackage.GetXmlFromUri(CustomPropertiesUri);
					else
					{
						// create a new extended properties part and add to the package
						PackagePart partCustom = _xlPackage.Package.CreatePart(CustomPropertiesUri, @"application/vnd.openxmlformats-officedocument.custom-properties+xml");

						// create the extended properties XML (with no entries in it)
						_xmlPropertiesCustom = new XmlDocument();
						XmlElement root = _xmlPropertiesCustom.CreateElement("Properties", schemaCustom);
						ExcelPackage.AddSchemaAttribute(root, schemaVt, "vt");
						_xmlPropertiesCustom.AppendChild(root);

						// save it to the package
						StreamWriter streamCustom = new StreamWriter(partCustom.GetStream(FileMode.Create, FileAccess.Write));
						_xmlPropertiesCustom.Save(streamCustom);
						streamCustom.Close();
						_xlPackage.Package.Flush();

						// create the relationship between the workbook and the new shared strings part
						_xlPackage.Package.CreateRelationship(CustomPropertiesUri, TargetMode.Internal, @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties");
						_xlPackage.Package.Flush();
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
		public string GetCustomPropertyValue(string propertyName)
		{
			string retValue = null;
			string searchString = string.Format("//ctp:Properties/ctp:property/@name[.='{0}']", propertyName);
			XmlNode node = CustomPropertiesXml.SelectSingleNode(searchString, _nsManager);
			if (node != null)
			{
				retValue = node.LastChild.InnerText;
			}
			return retValue;
		}

		/// <summary>
		/// Allows you to set the value of a current custom property or create 
		/// your own custom property.  
		/// Currently only supports string values.
		/// </summary>
		/// <param name="propertyName">The name of the property</param>
		/// <param name="propValue">The value of the property</param>
		public void SetCustomPropertyValue(string propertyName, string propValue)
		{
			// TODO:  provide support for other custom property data types
			string searchString = @"//ctp:Properties";
			XmlNode allProps = CustomPropertiesXml.SelectSingleNode(searchString, _nsManager);

			searchString = string.Format("//ctp:Properties/ctp:property/@ctp:name[.='{0}']", propertyName);
			XmlNode node = CustomPropertiesXml.SelectSingleNode(searchString, _nsManager);
			if (node == null)
			{
				// the property does not exist, so first find the max PID
				int pid = 4;
				foreach (XmlNode prop in allProps.ChildNodes)
				{
					XmlAttribute attr = (XmlAttribute)prop.Attributes.GetNamedItem("pid");
					if (attr != null)
					{
						int attrValue = int.Parse(attr.Value);
						if (attrValue > pid)
							pid = attrValue;
					}
				}
				pid++;
				// the property does not exist, so create the XML node
				XmlElement element = CustomPropertiesXml.CreateElement("property", schemaCustom);
				element.SetAttribute("fmtid", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}");
				element.SetAttribute("pid", pid.ToString());  // custom property pid
				element.SetAttribute("name", propertyName);

				XmlElement valueElement = CustomPropertiesXml.CreateElement("vt", "lpwstr", schemaVt);
				valueElement.InnerText = propValue;
				element.AppendChild(valueElement);

				CustomPropertiesXml.DocumentElement.AppendChild(element);
			}
			else
			{
				node.LastChild.InnerText = propValue;
			}
			
		}
		#endregion
		#endregion

		#region Save  // OfficeProperties save
		/// <summary>
		/// Saves the office document properties back to the package (if they exist!).
		/// </summary>
		protected internal void Save()
		{
			if (_xmlPropertiesCore != null)
			{
				_xlPackage.WriteDebugFile(_xmlPropertiesCore, "docProps", "core.xml");
				_xlPackage.SavePart(CorePropertiesUri, _xmlPropertiesCore);
			}
			if (_xmlPropertiesExtended != null)
			{
				_xlPackage.WriteDebugFile(_xmlPropertiesExtended, "docProps", "app.xml");
				_xlPackage.SavePart(ExtendedPropertiesUri, _xmlPropertiesExtended);
			}
			if (_xmlPropertiesCustom != null)
			{
				_xlPackage.WriteDebugFile(_xmlPropertiesCustom, "docProps", "custom.xml");
				_xlPackage.SavePart(CustomPropertiesUri, _xmlPropertiesCustom);
			}

		}
		#endregion

	}
}
