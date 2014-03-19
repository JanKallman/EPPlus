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
 * Jan Källman		Added		21-MAR-2011
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Base class for pivot table field groups
    /// </summary>
    public class ExcelPivotTableFieldGroup : XmlHelper
    {
        internal ExcelPivotTableFieldGroup(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
            
        }
    }
    /// <summary>
    /// A date group
    /// </summary>
    public class ExcelPivotTableFieldDateGroup : ExcelPivotTableFieldGroup
    {
        internal ExcelPivotTableFieldDateGroup(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
        }
        const string groupByPath = "d:fieldGroup/d:rangePr/@groupBy";
        /// <summary>
        /// How to group the date field
        /// </summary>
        public eDateGroupBy GroupBy
        {
            get
            {
                string v = GetXmlNodeString(groupByPath);
                if (v != "")
                {
                    return (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), v, true);
                }
                else
                {
                    throw (new Exception("Invalid date Groupby"));
                }
            }
            private set
            {
                SetXmlNodeString(groupByPath, value.ToString().ToLower());
            }
        }
        /// <summary>
        /// Auto detect start date
        /// </summary>
        public bool AutoStart
        {
            get
            {
                return GetXmlNodeBool("@autoStart", false);
            }
        }
        /// <summary>
        /// Auto detect end date
        /// </summary>
        public bool AutoEnd
        {
            get
            {
                return GetXmlNodeBool("@autoStart", false);
            }
        }
    }
    /// <summary>
    /// A pivot table field numeric grouping
    /// </summary>
    public class ExcelPivotTableFieldNumericGroup : ExcelPivotTableFieldGroup
    {
        internal ExcelPivotTableFieldNumericGroup(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
        }
        const string startPath = "d:fieldGroup/d:rangePr/@startNum";
        /// <summary>
        /// Start value
        /// </summary>
        public double Start
        {
            get
            {
                return (double)GetXmlNodeDoubleNull(startPath);
            }
            private set
            {
                SetXmlNodeString(startPath,value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string endPath = "d:fieldGroup/d:rangePr/@endNum";
        /// <summary>
        /// End value
        /// </summary>
        public double End
        {
            get
            {
                return (double)GetXmlNodeDoubleNull(endPath);
            }
            private set
            {
                SetXmlNodeString(endPath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
        const string groupIntervalPath = "d:fieldGroup/d:rangePr/@groupInterval";
        /// <summary>
        /// Interval
        /// </summary>
        public double Interval
        {
            get
            {
                return (double)GetXmlNodeDoubleNull(groupIntervalPath);
            }
            private set
            {
                SetXmlNodeString(groupIntervalPath, value.ToString(CultureInfo.InvariantCulture));
            }
        }
    }
}
