/* 
 * You may amend and distribute as you like, but don't remove this header!
 * 
 * EPPlus provides server-side generation of Excel 2007 spreadsheets.
 *
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
 * Jan Källman		                Initial Release		        2009-12-30
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Datalabel on chart level. 
    /// This class is inherited by ExcelChartSerieDataLabel
    /// </summary>
    public class ExcelChartDataLabel : XmlHelper
    {
       internal ExcelChartDataLabel(XmlNamespaceManager ns, XmlNode node)
           : base(ns,node)
       {
           XmlNode topNode = node.SelectSingleNode("c:dLbls", NameSpaceManager);
           if (topNode == null)
           {
               topNode = node.OwnerDocument.CreateElement("c", "dLbls", ExcelPackage.schemaChart);
               //node.InsertAfter(_topNode, node.SelectSingleNode("c:order", NameSpaceManager));
               InserAfter(node, "c:marker,c:tx,c:order,c:ser", topNode);
               SchemaNodeOrder = new string[] { "showVal", "showCatName", "showSerName", "showPercent", "separator", "showLeaderLines","spPr", "txPr" };
               topNode.InnerXml = "<c:showVal val=\"0\" />";
           }
           TopNode = topNode;
       }
       #region "Public properties"
       const string showValPath = "c:showVal/@val";
       public bool ShowValue
       {
           get
           {
               return GetXmlNodeBool(showValPath);
           }
           set
           {
               SetXmlNode(showValPath, value ? "1" : "0");
           }
       }
       const string showCatPath = "c:showCatName/@val";
       public bool ShowCategory
       {
           get
           {
               return GetXmlNodeBool(showCatPath);
           }
           set
           {
               SetXmlNode(showCatPath, value ? "1" : "0");
           }
       }
       const string showSerPath = "c:showSerName/@val";
       public bool ShowSeriesName
       {
           get
           {
               return GetXmlNodeBool(showSerPath);
           }
           set
           {
               SetXmlNode(showSerPath, value ? "1" : "0");
           }
       }
       const string showPerentPath = "c:showPercent/@val";
       public bool ShowPercent
       {
           get
           {
               return GetXmlNodeBool(showPerentPath);
           }
           set
           {
               SetXmlNode(showPerentPath, value ? "1" : "0");
           }
       }
       const string showLeaderLinesPath = "c:showLeaderLines/@val";
       public bool ShowLeaderLines
       {
           get
           {
               return GetXmlNodeBool(showLeaderLinesPath);
           }
           set
           {
               SetXmlNode(showLeaderLinesPath, value ? "1" : "0");
           }
       }
       const string separatorPath = "c:separator";
       public string Separator
       {
           get
           {
               return GetXmlNode(separatorPath);
           }
           set
           {
               SetXmlNode(separatorPath, value);
           }
       }

       ExcelDrawingFill _fill = null;
       public ExcelDrawingFill Fill
       {
           get
           {
               if (_fill == null)
               {
                   _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, "c:spPr");
               }
               return _fill;
           }
       }
       ExcelDrawingBorder _border = null;
       public ExcelDrawingBorder Border
       {
           get
           {
               if (_border == null)
               {
                   _border = new ExcelDrawingBorder(NameSpaceManager, TopNode, "c:spPr/a:ln");
               }
               return _border;
           }
       }
       ExcelTextFont _font = null;
       public ExcelTextFont Font
       {
           get
           {
               if (_font == null)
               {
                   if (TopNode.SelectSingleNode("c:txPr", NameSpaceManager) == null)
                   {
                       CreateNode("c:txPr/a:bodyPr");
                       CreateNode("c:txPr/a:lstStyle");
                   }
                   _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", new string[] { "showVal", "showCatName", "showSerName", "showPercent", "separator", "showLeaderLines", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" });
               }
               return _font;
           }
       }
       
       #endregion
       #region "Position Enum Traslation"
       protected string GetPosText(eLabelPosition pos)
       {
           switch (pos)
           {
               case eLabelPosition.Bottom:
                   return "b";
               case eLabelPosition.Center:
                   return "ctr";
               case eLabelPosition.InBase:
                   return "inBase";
               case eLabelPosition.InEnd:
                   return "inEnd";
               case eLabelPosition.Left:
                   return "l";
               case eLabelPosition.Right:
                   return "r";
               case eLabelPosition.Top:
                   return "t";
               case eLabelPosition.OutEnd:
                   return "outEnd";
               default:
                   return "bestFit";
           }
       }

       protected eLabelPosition GetPosEnum(string pos)
       {
           switch (pos)
           {
               case "b":
                   return eLabelPosition.Bottom;
               case "ctr":
                   return eLabelPosition.Center;
               case "inBase":
                   return eLabelPosition.InBase;
               case "inEnd":
                   return eLabelPosition.InEnd;
               case "l":
                   return eLabelPosition.Left;
               case "r":
                   return eLabelPosition.Right;
               case "t":
                   return eLabelPosition.Top;
               case "outEnd":
                   return eLabelPosition.OutEnd;
               default:
                   return eLabelPosition.BestFit;
           }
       }
    #endregion
    }
}
