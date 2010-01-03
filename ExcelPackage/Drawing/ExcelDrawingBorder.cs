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
 * Jan Källman		                Initial Release		        2009-12-22
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Drawing;

namespace OfficeOpenXml.Drawing
{
    public enum eLineCap
    {
        Flat,   //flat
        Round,  //rnd
        Square  //Sq
    }
    public enum eLineStyle
    {
        Dash,
        DashDot,
        Dot,
        LargeDash,
        LargeDashDot,
        LargeDashDotDot,
        Solid,
        SystemDash,
        SystemDashDot,
        SystemDashDotDot,
        SystemDot
    }
    /// <summary>
    /// Border for drawings
    /// </summary>    
    public class ExcelDrawingBorder : XmlHelper
    {
        string _linePath;
        public ExcelDrawingBorder(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath) : 
            base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder=new string[] {"c:chart"};
            _linePath = linePath;
            _lineStylePath = string.Format(_lineStylePath, linePath);
            _lineCapPath = string.Format(_lineCapPath, linePath);
        }
        #region "Public properties"
        ExcelDrawingFill _fill = null;
        public ExcelDrawingFill Fill
        {
            get
            {
                if (_fill == null)
                {
                    _fill = new ExcelDrawingFill(NameSpaceManager, TopNode, _linePath);
                }
                return _fill;
            }
        }
        string _lineStylePath = "{0}/a:prstDash/@val";
        public eLineStyle LineStyle
        {
            get
            {
                return TranslateLineStyle(GetXmlNode(_lineStylePath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNode(_lineStylePath, TranslateLineStyleText(value));
            }
        }
        string _lineCapPath = "{0}/@cap";
        public eLineCap LineCap
        {
            get
            {
                return TranslateLineCap(GetXmlNode(_lineCapPath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNode(_lineCapPath, TranslateLineCapText(value));
            }
        }
        #endregion
        #region "Translate Enum functions"
        private string TranslateLineStyleText(eLineStyle value)
        {
            string text=value.ToString();
            switch (value)
            {
                case eLineStyle.Dash:
                case eLineStyle.Dot:
                case eLineStyle.DashDot:
                case eLineStyle.Solid:
                    return text.Substring(0,1).ToLower() + text.Substring(1,text.Length-1); //First to Lower case.
                case eLineStyle.LargeDash:
                case eLineStyle.LargeDashDot:
                case eLineStyle.LargeDashDotDot:
                    return "lg" + text.Substring(5, text.Length - 5);
                case eLineStyle.SystemDash:
                case eLineStyle.SystemDashDot:
                case eLineStyle.SystemDashDotDot:
                case eLineStyle.SystemDot:
                    return "sys" + text.Substring(6, text.Length - 6);
                default:
                    throw(new Exception("Invalid Linestyle"));
            }
        }
        private eLineStyle TranslateLineStyle(string text)
        {
            switch (text)
            {
                case "dash":
                case "dot":
                case "dashDot":
                case "solid":
                    return (eLineStyle)Enum.Parse(typeof(eLineStyle), text, true);
                case "lgDash":
                case "lgDashDot":
                case "lgDashDotDot":
                    return (eLineStyle)Enum.Parse(typeof(eLineStyle), "Large" + text.Substring(2, text.Length - 2));
                case "sysDash":
                case "sysDashDot":
                case "sysDashDotDot":
                case "sysDot":
                    return (eLineStyle)Enum.Parse(typeof(eLineStyle), "System" + text.Substring(3, text.Length - 3));
                default:
                    throw (new Exception("Invalid Linestyle"));
            }
        }
        private string TranslateLineCapText(eLineCap value)
        {
            switch (value)
            {
                case eLineCap.Round:
                    return "rnd";
                case eLineCap.Square:
                    return "sq";
                default:
                    return "flat";
            }
        }
        private eLineCap TranslateLineCap(string text)
        {
            switch (text)
            {
                case "rnd":
                    return eLineCap.Round;
                case "sq":
                    return eLineCap.Square;
                default:
                    return eLineCap.Flat;
            }
        }
        #endregion

        
        //public ExcelDrawingFont Font
        //{
        //    get
        //    { 
            
        //    }
        //}
    }
}
