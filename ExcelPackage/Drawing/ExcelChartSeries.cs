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
using System.IO.Packaging;
using System.Collections;

namespace OfficeOpenXml.Drawing
{
   public class ExcelChartSeries : XmlHelper, IEnumerable
    {
       List<ExcelChartSerie> _list=new List<ExcelChartSerie>();
       ExcelChart _chart;
       XmlNode _node;
       XmlNamespaceManager _ns;
       internal ExcelChartSeries (ExcelChart chart, XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           _ns = ns;
           _chart=chart;
           _node=node;

           foreach(XmlNode n in node.SelectNodes("//c:ser",ns))
           {
               ExcelChartSerie s = new ExcelChartSerie(this, ns, n);
               _list.Add(s);
           }
       }

       #region IEnumerable Members

       public IEnumerator GetEnumerator()
       {
           return (_list.GetEnumerator());
       }
       /// <summary>
       /// Returns the serie at the specified position.  
       /// </summary>
       /// <param name="PositionID">The position of the series.</param>
       /// <returns></returns>
       public ExcelChartSerie this[int PositionID]
       {
           get
           {
               return (_list[PositionID]);
           }
       }
       public int Count
       {
           get
           {
               return _list.Count;
           }
       }
       public void Delete(int PositionID)
       {
           ExcelChartSerie ser = _list[PositionID];
           ser.TopNode.ParentNode.RemoveChild(ser.TopNode);
           _list.RemoveAt(PositionID);
       }
       #endregion
       public ExcelChart Chart
       {
           get
           {
               return _chart;
           }
       }
   }
}
