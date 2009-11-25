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
   public class ExcelChartSerie : XmlHelper
   {
       ExcelChartSeries _charts;
       XmlNode _node;
       XmlNamespaceManager _ns;
       public ExcelChartSerie(ExcelChartSeries charts, XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           _charts=charts;
           _node=node;
           _ns=ns;
       }
       const string headerPath="c:tx/c:v";
       /// <summary>
       /// Header for the serie.
       /// </summary>
       public string Header 
       {
           get
           {
                return GetXmlNode(headerPath);
            }
            set
            {
                //Where need this one 
                CreateNode(headerPath);
                SetXmlNode(headerPath, value);            
            }
        }
       const string seriesPath = "c:val/c:numRef/c:f";       
       /// <summary>
       /// Set this to a valid address or the drawing will be invalid.
       /// </summary>
       public string Series
       {
           get
           {
               return GetXmlNode(seriesPath);
           }
           set
           {
               if (_charts.Chart.ChartType == eChartType.xlBubble)
               {
                   throw(new Exception("Bubble charts is not supported yet"));
               }

               CreateNode(seriesPath,true);
               SetXmlNode(seriesPath, value);
               
               XmlNode cache = TopNode.SelectSingleNode("c:val/c:numRef/c:numCache", _ns);
               if (cache != null)
               {
                   cache.ParentNode.RemoveChild(cache);
               }

               XmlNode lit = TopNode.SelectSingleNode("c:val/c:numLit", _ns);
               if (lit != null)
               {
                   lit.ParentNode.RemoveChild(lit);
               }
           }

       }
       const string xSeriesPath = "c:cat/c:numRef/c:f";
       /// <summary>
       /// Set an address for the horisontal labels
       /// </summary>
       public string XSeries
       {
           get
           {
               return GetXmlNode(xSeriesPath);
           }
           set
           {
               CreateNode(xSeriesPath, true);
               SetXmlNode(xSeriesPath, value);

               XmlNode cache = TopNode.SelectSingleNode("c:cat/c:numRef/c:numCache", _ns);
               if (cache != null)
               {
                   cache.ParentNode.RemoveChild(cache);
               }

               XmlNode lit = TopNode.SelectSingleNode("c:cat/c:numLit", _ns);
               if (lit != null)
               {
                   lit.ParentNode.RemoveChild(lit);
               }
           }
       }
   }
}
