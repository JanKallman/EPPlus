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
 * Jan Källman		                Initial Release		        2009-12-22
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Drawing;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Fill for drawing objects
    /// </summary>
    public sealed class ExcelDrawingFill : XmlHelper
    {
        //ExcelShape _shp;                
        string _fillPath;
        XmlNode _fillNode;
        internal ExcelDrawingFill(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string fillPath) : 
            base(nameSpaceManager, topNode)
        {
          //  _shp=shp;
            _fillPath = fillPath;
            _fillNode = topNode.SelectSingleNode(_fillPath, NameSpaceManager);
            SchemaNodeOrder = new string[] { "tickLblPos", "spPr", "txPr","dLblPos", "crossAx", "printSettings", "showVal", "prstGeom", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill", "ln", "prstDash" };
            //Setfill node
            if (_fillNode != null)
            {
                _fillTypeNode = topNode.SelectSingleNode("solidFill");
                if (_fillTypeNode == null) _fillTypeNode = topNode.SelectSingleNode("noFill");
                if (_fillTypeNode == null) _fillTypeNode = topNode.SelectSingleNode("blipFill");
                if (_fillTypeNode == null) _fillTypeNode = topNode.SelectSingleNode("gradFill");
                if (_fillTypeNode == null) _fillTypeNode = topNode.SelectSingleNode("pattFill");
            }
        }
        eFillStyle _style;
        XmlNode _fillTypeNode = null;
        public eFillStyle Style
        {
            get
            {
                if (_fillTypeNode == null)
                {
                    return eFillStyle.SolidFill;
                }
                else
                {
                    _style=GetStyleEnum(_fillTypeNode.Name);
                }
                return _style;
            }
            set
            {
                if (value == eFillStyle.NoFill || value == eFillStyle.SolidFill)
                {
                    _style = value;
                    CreateFillTopNode(value);
                }
                else
                {
                    throw new NotImplementedException("Fillstyle not implemented");
                }
            }
        }

        private void CreateFillTopNode(eFillStyle value)
        {
            if (_fillTypeNode != null)
            {
                TopNode.RemoveChild(_fillTypeNode);
            }
            CreateNode(_fillPath + "/a:" + GetStyleText(value), false);
            _fillNode=TopNode.SelectSingleNode(_fillPath + "/a:" + GetStyleText(value), NameSpaceManager);
        }

        private eFillStyle GetStyleEnum(string name)
        {
            switch(name)
            {
                case "noFill":
                    return eFillStyle.NoFill;
                case "blipFill":
                    return eFillStyle.BlipFill;
                case "gradFill":
                    return eFillStyle.GradientFill;
                case "grpFill":
                    return eFillStyle.GroupFill;
                case "pattFill":
                    return eFillStyle.PatternFill;
                default:
                    return eFillStyle.SolidFill;
            }
        }

        private string GetStyleText(eFillStyle style)
        {
            switch (style)
            {
                case eFillStyle.BlipFill:
                    return "blipFill";
                case eFillStyle.GradientFill:
                    return "gradFill";
                case eFillStyle.GroupFill:
                    return "grpFill";
                case eFillStyle.NoFill:
                    return "noFill";                
                case eFillStyle.PatternFill:
                    return "pattFill";
                default:
                    return "solidFill";
            }
        }

        const string ColorPath = "/a:solidFill/a:srgbClr/@val";
        public Color Color
        {
            get
            {
                string col = GetXmlNodeString(_fillPath + ColorPath);
                if (col == "")
                {
                    return Color.FromArgb(79, 129, 189);
                }
                else
                {
                    return Color.FromArgb(int.Parse(col,System.Globalization.NumberStyles.AllowHexSpecifier));
                }
            }
            set
            {
                if (_fillTypeNode == null)
                {
                    _style = eFillStyle.SolidFill;
                }
                else if (_style != eFillStyle.SolidFill)
                {
                    throw new Exception("FillStyle must be set to SolidFill");
                }
                CreateNode(_fillPath, false);
                SetXmlNodeString(_fillPath + ColorPath, value.ToArgb().ToString("X").Substring(2, 6));
            }
        }
        const string alphaPath = "/a:solidFill/a:srgbClr/a:alpha/@val";
        /// <summary>
        /// Transparancy in percent
        /// </summary>
        public int Transparancy
        {
            get
            {
                return 100 - (GetXmlNodeInt(_fillPath + alphaPath) / 1000);
            }
            set
            {
                if (_fillTypeNode == null)
                {
                    _style = eFillStyle.SolidFill;
                    Color = Color.FromArgb(79, 129, 189);   //Set a Default color
                }
                else if (_style != eFillStyle.SolidFill)
                {
                    throw new Exception("FillStyle must be set to SolidFill");
                }
                //CreateNode(_fillPath, false);
                SetXmlNodeString(_fillPath + alphaPath, ((100 - value) * 1000).ToString());
            }
        }
    }
}
