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
 * Jan Källman		                Initial Release		        2009-12-22
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
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
        LongDash,
        LongDashDot,
        LongDashDotDot,
        Solid,
        SystemDash,
        SystemDashDot,
        SystemDashDotDot,
        SystemDot
    }
    /// <summary>
    /// Border for drawings
    /// </summary>    
    public sealed class ExcelDrawingBorder : XmlHelper
    {
        string _linePath;
        internal ExcelDrawingBorder(XmlNamespaceManager nameSpaceManager, XmlNode topNode, string linePath) : 
            base(nameSpaceManager, topNode)
        {
            SchemaNodeOrder = new string[] { "chart","tickLblPos", "spPr", "txPr","crossAx", "printSettings", "showVal", "showCatName", "showSerName", "showPercent", "separator", "showLeaderLines", "noFill", "solidFill", "blipFill", "gradFill", "noFill", "pattFill", "prstDash" };
            _linePath = linePath;   
            _lineStylePath = string.Format(_lineStylePath, linePath);
            _lineCapPath = string.Format(_lineCapPath, linePath);
            _lineWidth = string.Format(_lineWidth, linePath);
        }
        #region "Public properties"
        ExcelDrawingFill _fill = null;
        /// <summary>
        /// Fill
        /// </summary>
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
        /// <summary>
        /// Linestyle
        /// </summary>
        public eLineStyle LineStyle
        {
            get
            {
                return TranslateLineStyle(GetXmlNodeString(_lineStylePath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNodeString(_lineStylePath, TranslateLineStyleText(value));
            }
        }
        string _lineCapPath = "{0}/@cap";
        /// <summary>
        /// Linecap
        /// </summary>
        public eLineCap LineCap
        {
            get
            {
                return TranslateLineCap(GetXmlNodeString(_lineCapPath));
            }
            set
            {
                CreateNode(_linePath, false);
                SetXmlNodeString(_lineCapPath, TranslateLineCapText(value));
            }
        }
        string _lineWidth = "{0}/@w";
        /// <summary>
        /// Width in pixels
        /// </summary>
        public int Width
        {
            get
            {
                return GetXmlNodeInt(_lineWidth) / 12700;
            }
            set
            {
                SetXmlNodeString(_lineWidth, (value * 12700).ToString());
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
                case eLineStyle.LongDash:
                case eLineStyle.LongDashDot:
                case eLineStyle.LongDashDotDot:
                    return "lg" + text.Substring(4, text.Length - 4);
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
                    return (eLineStyle)Enum.Parse(typeof(eLineStyle), "Long" + text.Substring(2, text.Length - 2));
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
