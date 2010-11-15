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
 * Jan Källman		                Initial Release		        2009-10-01
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Drawing.Chart
{
    public enum eLegendPosition
    {
        Top,
        Left,
        Right,
        Bottom,
        TopRight
    }
    /// <summary>
    /// Chart ledger
    /// </summary>
    public class ExcelChartLegend : XmlHelper
    {
        ExcelChart _chart;
        internal ExcelChartLegend(XmlNamespaceManager ns, XmlNode node, ExcelChart chart)
           : base(ns,node)
       {
           _chart=chart;
           SchemaNodeOrder = new string[] { "legendPos", "layout","overlay", "txPr", "bodyPr", "lstStyle", "spPr" };
       }
        const string POSITION_PATH = "c:legendPos/@val";
        public eLegendPosition Position 
        {
            get
            {
                switch(GetXmlNodeString(POSITION_PATH).ToLower())
                {
                    case "t":
                        return eLegendPosition.Top;
                    case "b":
                        return eLegendPosition.Bottom;
                    case "l":
                        return eLegendPosition.Left;
                    case "tr":
                        return eLegendPosition.TopRight;
                    default:
                        return eLegendPosition.Right;
                }
            }
            set
            {
                if (TopNode == null) throw(new Exception("Can't set position. Chart has no legend"));
                switch (value)
                {
                    case eLegendPosition.Top:
                        SetXmlNodeString(POSITION_PATH, "t");
                        break;
                    case eLegendPosition.Bottom:
                        SetXmlNodeString(POSITION_PATH, "b");
                        break;
                    case eLegendPosition.Left:
                        SetXmlNodeString(POSITION_PATH, "l");
                        break;
                    case eLegendPosition.TopRight:
                        SetXmlNodeString(POSITION_PATH, "tr");
                        break;
                    default:
                        SetXmlNodeString(POSITION_PATH, "r");
                        break;
                }
            }
        }
        const string OVERLAY_PATH = "c:overlay/@val";
        public bool Overlay
        {
            get
            {
                return GetXmlNodeBool(OVERLAY_PATH,false);
            }
            set
            {
                if (TopNode == null) throw (new Exception("Can't set overlay. Chart has no legend"));
                SetXmlNodeBool(OVERLAY_PATH, value, false);
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
                    if (TopNode.SelectSingleNode("c:txPr",NameSpaceManager) == null)
                    {
                        CreateNode("c:txPr/a:bodyPr");
                        CreateNode("c:txPr/a:lstStyle");
                    }
                    _font = new ExcelTextFont(NameSpaceManager, TopNode, "c:txPr/a:p/a:pPr/a:defRPr", new string[] { "legendPos", "layout", "pPr", "defRPr", "solidFill", "uFill", "latin", "cs", "r", "rPr", "t" });
                }
                return _font;
            }
        }
        public void Remove()
        {
            if (TopNode == null) return;
            TopNode.ParentNode.RemoveChild(TopNode);
            TopNode = null;
        }
        public void Add()
        {
            if(TopNode!=null) return;

            XmlHelper xml = new XmlHelper(NameSpaceManager, _chart.ChartXml);
            xml.SchemaNodeOrder=_chart.SchemaNodeOrder;

            xml.CreateNode("c:chartSpace/c:chart/c:legend");
            TopNode = _chart.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", NameSpaceManager);
            TopNode.InnerXml="<c:legendPos val=\"r\" /><c:layout />";                        
        }
    }
}
