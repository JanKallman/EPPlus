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

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;
using System.IO.Packaging;
namespace OfficeOpenXml
{
    /// <summary>
    /// Help class containing XML functions. 
    /// Can be Inherited 
    /// </summary>
    public class XmlHelper
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

        internal void CreateNode(string path)
        {
            CreateNode(path, false);
        }
        internal void CreateNode(string path, bool insertFirst)
        {
            XmlNode node = TopNode;
            foreach (string subPath in path.Split('/'))
            {
                if (node.SelectSingleNode(subPath, NameSpaceManager) == null)
                {
                    XmlNode addedNode;
                    string nodeName;
                    string nameSpaceURI = "";
                    string[] nameSplit = subPath.Split(':');

                    if (nameSplit.Length > 1)
                    {
                        nameSpaceURI = NameSpaceManager.LookupNamespace(nameSplit[0]);
                        nodeName=nameSplit[1];
                    }
                    else
                    {
                        nodeName=nameSplit[0];
                    }

                    if (subPath.StartsWith("@"))
                    {
                        XmlAttribute addedAtt = node.OwnerDocument.CreateAttribute(subPath.Substring(1,subPath.Length-1), nameSpaceURI);
                        node.Attributes.Append(addedAtt);
                    }
                    else
                    {
                        addedNode = node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
                        if (insertFirst)
                        {
                            node.PrependChild(addedNode);
                        }
                        else
                        {
                            node.AppendChild(addedNode);
                        }
                    }
                }
                node = node.SelectSingleNode(subPath, NameSpaceManager);
            }
        }
        internal void DeleteNode(string path)
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
        internal void SetXmlNode(string path, string value)
        {
            SetXmlNode(TopNode, path, value, false, false);
        }
        internal void SetXmlNode(XmlNode node, string path, string value)
        {
            SetXmlNode(node, path, value, false, false);
        }
        internal void SetXmlNode(XmlNode node, string path, string value, bool removeIfBlank)
        {
            SetXmlNode(node, path, value, removeIfBlank, false);
        }
        internal void SetXmlNode(XmlNode node, string path, string value, bool removeIfBlank, bool insertFirst)
        {
            if (node == null)
            {
                return;
            }
            if (value == "" && removeIfBlank)
            {
                DeleteNode(path);
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

        internal bool GetXmlNodeBool(string path)
        {
            string value=GetXmlNode(path);
            if (value == "1" || value == "-1" || value == "True")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        internal int GetXmlNodeInt(string path)
        {
            int i;
            if (int.TryParse(GetXmlNode(path), out i))
            {
                return i;
            }
            else
            {
                return int.MinValue;
            }
        }
        internal decimal GetXmlNodeDecimal(string path)
        {
            decimal d;
            if (decimal.TryParse(GetXmlNode(path), out d))
            {
                return d;
            }
            else
            {
                return 0;
            }
        }
        internal string GetXmlNode(string path)
        {
            if (TopNode == null)
            {
                return "";
            }
            XmlNode nameNode = TopNode.SelectSingleNode(path, NameSpaceManager);
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
        internal Uri GetNewUri(Package package, string sUri)
        {
            int id = 1;
            Uri uri;
            do
            {
                uri = new Uri(string.Format(sUri, id++), UriKind.Relative);
            }
            while (package.PartExists(uri));
            return uri;
        }       
    }
}
