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
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml helper class for cell style classes
    /// </summary>
    public abstract class  StyleXmlHelper : XmlHelper
    {
        internal StyleXmlHelper(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        { 

        }
        internal StyleXmlHelper(XmlNamespaceManager nameSpaceManager, XmlNode topNode) : base(nameSpaceManager, topNode)
        {
        }
        internal abstract XmlNode CreateXmlNode(XmlNode top);
        internal abstract string Id
        {
            get;
        }
        internal long useCnt=0;
        internal int newID=int.MinValue;
        protected bool GetBoolValue(XmlNode topNode, string path)
        {
            var node = topNode.SelectSingleNode(path, NameSpaceManager);
            if (node is XmlAttribute)
            {
                return node.Value != "0";
            }
            else
            {
                if (node != null && ((node.Attributes["val"] != null && node.Attributes["val"].Value != "0") || node.Attributes["val"] == null))
                {
                    return true;
                }
                else
                {
                    return false;
                }                
            }
        }

    }
}
