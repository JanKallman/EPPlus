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
using System.Globalization;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// 3D settings
    /// </summary>
    public sealed class ExcelView3D : XmlHelper
    {
       internal ExcelView3D(XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           SchemaNodeOrder = new string[] { "rotX", "hPercent", "rotY", "depthPercent","rAngAx", "perspective"};
       }
       const string perspectivePath = "c:perspective/@val";
       /// <summary>
       /// Degree of perspective 
       /// </summary>
       public decimal Perspective
       {
           get
           {
               return GetXmlNodeInt(perspectivePath);
           }
           set
           {
               SetXmlNodeString(perspectivePath, value.ToString(CultureInfo.InvariantCulture));
           }
       }
       const string rotXPath = "c:rotX/@val";
       /// <summary>
       /// Rotation X-axis
       /// </summary>
       public decimal RotX
       {
           get
           {
               return GetXmlNodeDecimal(rotXPath);
           }
           set
           {
               CreateNode(rotXPath);
               SetXmlNodeString(rotXPath, value.ToString(CultureInfo.InvariantCulture));
           }
       }
       const string rotYPath = "c:rotY/@val";
       /// <summary>
       /// Rotation Y-axis
       /// </summary>
       public decimal RotY
       {
           get
           {
               return GetXmlNodeDecimal(rotYPath);
           }
           set
           {
               CreateNode(rotYPath);
               SetXmlNodeString(rotYPath, value.ToString(CultureInfo.InvariantCulture));
           }
       }
       const string rAngAxPath = "c:rAngAx/@val";
       /// <summary>
       /// Right Angle Axes
       /// </summary>
       public bool RightAngleAxes
       {
           get
           {
               return GetXmlNodeBool(rAngAxPath);
           }
           set
           {
               SetXmlNodeBool(rAngAxPath, value);
           }
       }
       const string depthPercentPath = "c:depthPercent/@val";
       /// <summary>
       /// Depth % of base
       /// </summary>
       public int DepthPercent
       {
           get
           {
               return GetXmlNodeInt(depthPercentPath);
           }
           set
           {
               if (value < 0 || value > 2000)
               {
                   throw(new ArgumentOutOfRangeException("Value must be between 0 and 2000"));
               }
               SetXmlNodeString(depthPercentPath, value.ToString());
           }
       }
       const string heightPercentPath = "c:hPercent/@val";
       /// <summary>
       /// Height % of base
       /// </summary>
       public int HeightPercent
       {
           get
           {
               return GetXmlNodeInt(heightPercentPath);
           }
           set
           {
               if (value < 5 || value > 500)
               {
                   throw (new ArgumentOutOfRangeException("Value must be between 5 and 500"));
               }
               SetXmlNodeString(heightPercentPath, value.ToString());
           }
       }
    }
}
