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
 * Author							Change						         Date
 * ******************************************************************************
 * Jan Källman		    Initial Release		         2009-10-01
 * Jan Källman		    License changed GPL-->LGPL 2011-12-27
 * Eyal Seagull       Add "CreateComplexNode"    2012-04-03
 * Eyal Seagull       Add "DeleteTopNode"        2012-04-13
 *******************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;
using System.IO.Packaging;
using System.Globalization;
namespace OfficeOpenXml
{
	/// <summary>
	/// Help class containing XML functions. 
	/// Can be Inherited 
	/// </summary>
	public abstract class XmlHelper
	{
		internal delegate int ChangedEventHandler(StyleBase sender, Style.StyleChangeEventArgs e);

		internal XmlHelper(XmlNamespaceManager nameSpaceManager)
		{
			TopNode = null;
			NameSpaceManager = nameSpaceManager;
		}

		internal XmlHelper(XmlNamespaceManager nameSpaceManager, XmlNode topNode)
		{
			TopNode = topNode;
			NameSpaceManager = nameSpaceManager;
		}
		//internal bool ChangedFlag;
		internal XmlNamespaceManager NameSpaceManager { get; set; }
		internal XmlNode TopNode { get; set; }
		string[] _schemaNodeOrder = null;
		/// <summary>
		/// Schema order list
		/// </summary>
		internal string[] SchemaNodeOrder
		{
			get
			{
				return _schemaNodeOrder;
			}
			set
			{
				_schemaNodeOrder = value;
			}
		}
		internal XmlNode CreateNode(string path)
		{
			if (path == "")
				return TopNode;
			else
				return CreateNode(path, false);
		}
		internal XmlNode CreateNode(string path, bool insertFirst)
		{
			XmlNode node = TopNode;
			XmlNode prependNode = null;
			foreach (string subPath in path.Split('/'))
			{
				XmlNode subNode = node.SelectSingleNode(subPath, NameSpaceManager);
				if (subNode == null)
				{
					string nodeName;
					string nodePrefix;

					string nameSpaceURI = "";
					string[] nameSplit = subPath.Split(':');

					if (SchemaNodeOrder != null && subPath[0] != '@')
					{
						insertFirst = false;
						prependNode = GetPrependNode(subPath, node);
					}

					if (nameSplit.Length > 1)
					{
						nodePrefix = nameSplit[0];
						if (nodePrefix[0] == '@') nodePrefix = nodePrefix.Substring(1, nodePrefix.Length - 1);
						nameSpaceURI = NameSpaceManager.LookupNamespace(nodePrefix);
						nodeName = nameSplit[1];
					}
					else
					{
						nodePrefix = "";
						nameSpaceURI = "";
						nodeName = nameSplit[0];
					}
					if (subPath.StartsWith("@"))
					{
						XmlAttribute addedAtt = node.OwnerDocument.CreateAttribute(subPath.Substring(1, subPath.Length - 1), nameSpaceURI);  //nameSpaceURI
						node.Attributes.Append(addedAtt);
					}
					else
					{
						if (nodePrefix == "")
						{
							subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
						}
						else
						{
							if (nodePrefix == "" || (node.OwnerDocument != null && node.OwnerDocument.DocumentElement != null && node.OwnerDocument.DocumentElement.NamespaceURI == nameSpaceURI &&
									node.OwnerDocument.DocumentElement.Prefix == ""))
							{
								subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
							}
							else
							{
								subNode = node.OwnerDocument.CreateElement(nodePrefix, nodeName, nameSpaceURI);
							}
						}
						if (prependNode != null)
						{
							node.InsertBefore(subNode, prependNode);
							prependNode = null;
						}
						else if (insertFirst)
						{
							node.PrependChild(subNode);
						}
						else
						{
							node.AppendChild(subNode);
						}
					}
				}
				node = subNode;
			}
			return node;
		}

		/// <summary>
		/// Options to insert a node in the XmlDocument
		/// </summary>
		internal enum eNodeInsertOrder
		{
			/// <summary>
			/// Insert as first node of "topNode"
			/// </summary>
			First,

			/// <summary>
			/// Insert as the last child of "topNode"
			/// </summary>
			Last,

			/// <summary>
			/// Insert after the "referenceNode"
			/// </summary>
			After,

			/// <summary>
			/// Insert before the "referenceNode"
			/// </summary>
			Before,

			/// <summary>
			/// Use the Schema List to insert in the right order. If the Schema list
			/// is null or empty, consider "Last" as the selected option
			/// </summary>
			SchemaOrder
		}

		/// <summary>
		/// Create a complex node. Insert the node according to SchemaOrder
		/// using the TopNode as the parent
		/// </summary>
		/// <param name="path"></param>
		/// <returns></returns>
		internal XmlNode CreateComplexNode(
			string path)
		{
			return CreateComplexNode(
				TopNode,
				path,
				eNodeInsertOrder.SchemaOrder,
				null);
		}

		/// <summary>
		/// Create a complex node. Insert the node according to the <paramref name="path"/>
		/// using the <paramref name="topNode"/> as the parent
		/// </summary>
		/// <param name="topNode"></param>
		/// <param name="path"></param>
		/// <returns></returns>
		internal XmlNode CreateComplexNode(
			XmlNode topNode,
			string path)
		{
			return CreateComplexNode(
				topNode,
				path,
				eNodeInsertOrder.SchemaOrder,
				null);
		}

		/// <summary>
		/// Creates complex XML nodes
    /// </summary>
    /// <remarks>
		/// 1. "d:conditionalFormatting"
		///		1.1. Creates/find the first "conditionalFormatting" node
		/// 
		/// 2. "d:conditionalFormatting/@sqref"
		///		2.1. Creates/find the first "conditionalFormatting" node
		///		2.2. Creates (if not exists) the @sqref attribute
		///
		/// 3. "d:conditionalFormatting/@id='7'/@sqref='A9:B99'"
		///		3.1. Creates/find the first "conditionalFormatting" node
		///		3.2. Creates/update its @id attribute to "7"
		///		3.3. Creates/update its @sqref attribute to "A9:B99"
		///
		/// 4. "d:conditionalFormatting[@id='7']/@sqref='X1:X5'"
		///		4.1. Creates/find the first "conditionalFormatting" node with @id=7
		///		4.2. Creates/update its @sqref attribute to "X1:X5"
		///	
		/// 5. "d:conditionalFormatting[@id='7']/@id='8'/@sqref='X1:X5'/d:cfRule/@id='AB'"
		///		5.1. Creates/find the first "conditionalFormatting" node with @id=7
		///		5.2. Set its @id attribute to "8"
		///		5.2. Creates/update its @sqref attribute and set it to "X1:X5"
		///		5.3. Creates/find the first "cfRule" node (inside the node)
		///		5.4. Creates/update its @id attribute to "AB"
		///	
		/// 6. "d:cfRule/@id=''"
		///		6.1. Creates/find the first "cfRule" node
		///		6.1. Remove the @id attribute
    ///	</remarks>
		/// <param name="topNode"></param>
		/// <param name="path"></param>
		/// <param name="nodeInsertOrder"></param>
		/// <param name="referenceNode"></param>
		/// <returns>The last node creates/found</returns>
		internal XmlNode CreateComplexNode(
			XmlNode topNode,
			string path,
			eNodeInsertOrder nodeInsertOrder,
			XmlNode referenceNode)
		{
			// Path is obrigatory
			if ((path == null) || (path == string.Empty))
			{
				return topNode;
			}

			XmlNode node = topNode;
			string nameSpaceURI = string.Empty;

      //TODO: BUG: when the "path" contains "/" in an attrribue value, it gives an error.

			// Separate the XPath to Nodes and Attributes
			foreach (string subPath in path.Split('/'))
			{
				// The subPath can be any one of those:
				// nodeName
				// x:nodeName
				// nodeName[find criteria]
				// x:nodeName[find criteria]
				// @attribute
				// @attribute='attribute value'

				// Check if the subPath has at least one character
				if (subPath.Length > 0)
				{
					// Check if the subPath is an attribute (with or without value)
					if (subPath.StartsWith("@"))
					{
						// @attribute										--> Create attribute
						// @attribute=''								--> Remove attribute
						// @attribute='attribute value' --> Create attribute + update value
						string[] attributeSplit = subPath.Split('=');
						string attributeName = attributeSplit[0].Substring(1, attributeSplit[0].Length - 1);
						string attributeValue = null;	// Null means no attribute value

						// Check if we have an attribute value to set
						if (attributeSplit.Length > 1)
						{
							// Remove the ' or " from the attribute value
							attributeValue = attributeSplit[1].Replace("'", "").Replace("\"", "");
						}

						// Get the attribute (if exists)
						XmlAttribute attribute = (XmlAttribute)(node.Attributes.GetNamedItem(attributeName));

						// Remove the attribute if value is empty (not null)
						if (attributeValue == string.Empty)
						{
							// Only if the attribute exists
							if (attribute != null)
							{
								node.Attributes.Remove(attribute);
							}
						}
						else
						{
							// Create the attribue if does not exists
							if (attribute == null)
							{
								// Create the attribute
								attribute = node.OwnerDocument.CreateAttribute(
									attributeName);

								// Add it to the current node
								node.Attributes.Append(attribute);
							}

							// Update the attribute value
							if (attributeValue != null)
							{
								node.Attributes[attributeName].Value = attributeValue;
							}
						}
					}
					else
					{
						// nodeName
						// x:nodeName
						// nodeName[find criteria]
						// x:nodeName[find criteria]

						// Look for the node (with or without filter criteria)
						XmlNode subNode = node.SelectSingleNode(subPath, NameSpaceManager);

						// Check if the node does not exists
						if (subNode == null)
						{
							string nodeName;
							string nodePrefix;
							string[] nameSplit = subPath.Split(':');
							nameSpaceURI = string.Empty;

							// Check if the name has a prefix like "d:nodeName"
							if (nameSplit.Length > 1)
							{
								nodePrefix = nameSplit[0];
								nameSpaceURI = NameSpaceManager.LookupNamespace(nodePrefix);
								nodeName = nameSplit[1];
							}
							else
							{
								nodePrefix = string.Empty;
								nameSpaceURI = string.Empty;
								nodeName = nameSplit[0];
							}

							// Check if we have a criteria part in the node name
							if (nodeName.IndexOf("[") > 0)
							{
								// remove the criteria from the node name
								nodeName = nodeName.Substring(0, nodeName.IndexOf("["));
							}

							if (nodePrefix == string.Empty)
							{
								subNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
							}
							else
							{
								if (node.OwnerDocument != null
									&& node.OwnerDocument.DocumentElement != null
									&& node.OwnerDocument.DocumentElement.NamespaceURI == nameSpaceURI
									&& node.OwnerDocument.DocumentElement.Prefix == string.Empty)
								{
									subNode = node.OwnerDocument.CreateElement(
										nodeName,
										nameSpaceURI);
								}
								else
								{
									subNode = node.OwnerDocument.CreateElement(
										nodePrefix,
										nodeName,
										nameSpaceURI);
								}
							}

							// Check if we need to use the "SchemaOrder"
							if (nodeInsertOrder == eNodeInsertOrder.SchemaOrder)
							{
								// Check if the Schema Order List is empty
								if ((SchemaNodeOrder == null) || (SchemaNodeOrder.Length == 0))
								{
									// Use the "Insert Last" option when Schema Order List is empty
									nodeInsertOrder = eNodeInsertOrder.Last;
								}
								else
								{
									// Find the prepend node in order to insert
									referenceNode = GetPrependNode(nodeName, node);

									if (referenceNode != null)
									{
										nodeInsertOrder = eNodeInsertOrder.Before;
									}
									else
									{
										nodeInsertOrder = eNodeInsertOrder.Last;
									}
								}
							}

							switch (nodeInsertOrder)
							{
								case eNodeInsertOrder.After:
									node.InsertAfter(subNode, referenceNode);
									break;

								case eNodeInsertOrder.Before:
									node.InsertBefore(subNode, referenceNode);
									break;

								case eNodeInsertOrder.First:
									node.PrependChild(subNode);
									break;

								case eNodeInsertOrder.Last:
									node.AppendChild(subNode);
									break;
							}
						}

						// Make the newly created node the top node when the rest of the path
						// is being evaluated. So newly created nodes will be the children of the
						// one we just created.
						node = subNode;
					}
				}
			}

			// Return the last created/found node
			return node;
		}

		/// <summary>
		/// return Prepend node
		/// </summary>
		/// <param name="nodeName">name of the node to check</param>
		/// <param name="node">Topnode to check children</param>
		/// <returns></returns>
		private XmlNode GetPrependNode(string nodeName, XmlNode node)
		{
			int pos = GetNodePos(nodeName);
			if (pos < 0)
			{
				return null;
			}
			XmlNode prependNode = null;
			foreach (XmlNode childNode in node.ChildNodes)
			{
				int childPos = GetNodePos(childNode.Name);
				if (childPos > -1)  //Found?
				{
					if (childPos > pos) //Position is before
					{
						prependNode = childNode;
						break;
					}
				}
			}
			return prependNode;
		}
		private int GetNodePos(string nodeName)
		{
			int ix = nodeName.IndexOf(":");
			if (ix > 0)
			{
				nodeName = nodeName.Substring(ix + 1, nodeName.Length - (ix + 1));
			}
			for (int i = 0; i < _schemaNodeOrder.Length; i++)
			{
				if (nodeName == _schemaNodeOrder[i])
				{
					return i;
				}
			}
			return -1;
		}
		internal void DeleteAllNode(string path)
		{
			string[] split = path.Split('/');
			XmlNode node = TopNode;
			foreach (string s in split)
			{
				node = node.SelectSingleNode(s, NameSpaceManager);
				if (node != null)
				{
					if (node is XmlAttribute)
					{
						(node as XmlAttribute).OwnerElement.Attributes.Remove(node as XmlAttribute);
					}
					else
					{
						node.ParentNode.RemoveChild(node);
					}
				}
				else
				{
					break;
				}
			}
		}
		internal void DeleteNode(string path)
		{
			var node = TopNode.SelectSingleNode(path, NameSpaceManager);
			if (node != null)
			{
				if (node is XmlAttribute)
				{
					var att = (XmlAttribute)node;
					att.OwnerElement.Attributes.Remove(att);
				}
				else
				{
					node.ParentNode.RemoveChild(node);
				}
			}
		}
    internal void DeleteTopNode()
    {
      TopNode.ParentNode.RemoveChild(TopNode);
    }
		internal void SetXmlNodeString(string path, string value)
		{
			SetXmlNodeString(TopNode, path, value, false, false);
		}
		internal void SetXmlNodeString(string path, string value, bool removeIfBlank)
		{
			SetXmlNodeString(TopNode, path, value, removeIfBlank, false);
		}
		internal void SetXmlNodeString(XmlNode node, string path, string value)
		{
			SetXmlNodeString(node, path, value, false, false);
		}
		internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank)
		{
			SetXmlNodeString(node, path, value, removeIfBlank, false);
		}
		internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank, bool insertFirst)
		{
			if (node == null)
			{
				return;
			}
			if (value == "" && removeIfBlank)
			{
				DeleteAllNode(path);
			}
			else
			{
				XmlNode nameNode = node.SelectSingleNode(path, NameSpaceManager);
				if (nameNode == null)
				{
					CreateNode(path, insertFirst);
					nameNode = node.SelectSingleNode(path, NameSpaceManager);
				}
				//if (nameNode.InnerText != value) HasChanged();
				nameNode.InnerText = value;
			}
		}
		internal void SetXmlNodeBool(string path, bool value)
		{
			SetXmlNodeString(TopNode, path, value ? "1" : "0", false, false);
		}
		internal void SetXmlNodeBool(string path, bool value, bool removeIf)
		{
			if (value == removeIf)
			{
				var node = TopNode.SelectSingleNode(path, NameSpaceManager);
				if (node != null)
				{
					if (node is XmlAttribute)
					{
						var elem = (node as XmlAttribute).OwnerElement;
						elem.ParentNode.RemoveChild(elem);
					}
					else
					{
						TopNode.RemoveChild(node);
					}
				}
			}
			else
			{
				SetXmlNodeString(TopNode, path, value ? "1" : "0", false, false);
			}
		}
		internal bool ExistNode(string path)
		{
			if (TopNode == null || TopNode.SelectSingleNode(path, NameSpaceManager) == null)
			{
				return false;
			}
			else
			{
				return true;
			}
		}
		internal bool? GetXmlNodeBoolNullable(string path)
		{
			var value = GetXmlNodeString(path);
			if (string.IsNullOrEmpty(value))
			{
				return null;
			}
			return GetXmlNodeBool(path);
		}
		internal bool GetXmlNodeBool(string path)
		{
			return GetXmlNodeBool(path, false);
		}
		internal bool GetXmlNodeBool(string path, bool blankValue)
		{
			string value = GetXmlNodeString(path);
			if (value == "1" || value == "-1" || value == "True")
			{
				return true;
			}
			else if (value == "")
			{
				return blankValue;
			}
			else
			{
				return false;
			}
		}
		internal int GetXmlNodeInt(string path)
		{
			int i;
			if (int.TryParse(GetXmlNodeString(path), out i))
			{
				return i;
			}
			else
			{
				return int.MinValue;
			}
		}
        internal int? GetXmlNodeIntNull(string path)
        {
            int i;
            string s = GetXmlNodeString(path);
            if (s!="" && int.TryParse(s, out i))
            {
                return i;
            }
            else
            {
                return null;
            }
        }
                }
                else
                {
                    return null;
                }
            }
        }

		internal decimal GetXmlNodeDecimal(string path)
		{
			decimal d;
			if (decimal.TryParse(GetXmlNodeString(path), NumberStyles.Any, CultureInfo.InvariantCulture, out d))
			{
				return d;
			}
			else
			{
				return 0;
			}
		}
		internal double? GetXmlNodeDoubleNull(string path)
		{
			string s = GetXmlNodeString(path);
			if (s == "")
			{
				return null;
			}
			else
			{
				double v;
				if (double.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out v))
				{
					return v;
				}
				else
				{
					return null;
				}
			}
		}
		internal double GetXmlNodeDouble(string path)
		{
			string s = GetXmlNodeString(path);
			if (s == "")
			{
				return double.NaN;
			}
			else
			{
				double v;
				if (double.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out v))
				{
					return v;
				}
				else
				{
					return double.NaN;
				}
			}
		}
    internal string GetXmlNodeString(XmlNode node, string path)
    {
      if (node == null)
      {
        return "";
      }

      XmlNode nameNode = node.SelectSingleNode(path, NameSpaceManager);

      if (nameNode != null)
      {
        if (nameNode.NodeType == XmlNodeType.Attribute)
        {
          return nameNode.Value != null ? nameNode.Value : "";
        }
        else
        {
          return nameNode.InnerText;
        }
      }
      else
      {
        return "";
      }
    }
		internal string GetXmlNodeString(string path)
		{
     return GetXmlNodeString(TopNode, path);
		}
		internal static Uri GetNewUri(Package package, string sUri)
		{
			return GetNewUri(package, sUri, 1);
		}
		internal static Uri GetNewUri(Package package, string sUri, int id)
		{
			Uri uri;
			do
			{
				uri = new Uri(string.Format(sUri, id++), UriKind.Relative);
			}
			while (package.PartExists(uri));
			return uri;
		}
		/// <summary>
		/// Insert the new node before any of the nodes in the comma separeted list
		/// </summary>
		/// <param name="parentNode">Parent node</param>
		/// <param name="beforeNodes">comma separated list containing nodes to insert after. Left to right order</param>
		/// <param name="newNode">The new node to be inserterd</param>
		internal void InserAfter(XmlNode parentNode, string beforeNodes, XmlNode newNode)
		{
			string[] nodePaths = beforeNodes.Split(',');

			foreach (string nodePath in nodePaths)
			{
				XmlNode node = parentNode.SelectSingleNode(nodePath, NameSpaceManager);
				if (node != null)
				{
					parentNode.InsertAfter(newNode, node);
					return;
				}
			}
			parentNode.InsertAfter(newNode, null);
		}
	}
}
