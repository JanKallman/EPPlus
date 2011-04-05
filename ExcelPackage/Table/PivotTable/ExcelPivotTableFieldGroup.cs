/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * All rights reserved.
 * 
 * EPPlus is an Open Source project provided under the 
 * GNU General Public License (GPL) as published by the 
 * Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
 * 
 * See http://epplus.codeplex.com/ for details
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
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		21-MAR-2011
 *******************************************************************************/
using System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Globalization;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableFieldGroup : XmlHelper
    {
        internal ExcelPivotTableFieldGroup(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
            
        }
    }
    public class ExcelPivotTableFieldDateGroup : ExcelPivotTableFieldGroup
    {
        internal ExcelPivotTableFieldDateGroup(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
        }
        const string groupByPath = "d:fieldGroup/d:rangePr/@groupBy";
        public eDateGroupBy GroupBy
        {
            get
            {
                string v = GetXmlNodeString(groupByPath);
                if (v == "")
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
    }
    public class ExcelPivotTableFieldNumericGroup : ExcelPivotTableFieldGroup
    {
        internal ExcelPivotTableFieldNumericGroup(XmlNamespaceManager ns, XmlNode topNode) :
            base(ns, topNode)
        {
        }
        //internal ExcelPivotTableFieldNumericGroup(XmlNamespaceManager ns, XmlNode topNode, double start, double end, double interval) :
        //    base(ns, topNode)
        //{
        //    Start = start;
        //    End = end;
        //    Interval = interval;
        //}
        const string startPath = "d:fieldGroup/d:rangePr/@startNum";
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
