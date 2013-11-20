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
 * Jan Källman		Initial Release		        2010-06-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Globalization;
using System.Drawing;


namespace OfficeOpenXml.Drawing.Vml
{
    /// <summary>
    /// Drawing object used for header and footer pictures
    /// </summary>
    public class ExcelVmlDrawingPicture : ExcelVmlDrawingBase
    {
        ExcelWorksheet _worksheet;
        internal ExcelVmlDrawingPicture(XmlNode topNode, XmlNamespaceManager ns, ExcelWorksheet ws) :
            base(topNode, ns)
        {
            _worksheet = ws;
        }
        /// <summary>
        /// Position ID
        /// </summary>
        public string Position
        {
            get
            {
                return GetXmlNodeString("@id");
            }
        }
        /// <summary>
        /// The width in points
        /// </summary>
        public double Width
        {
            get
            {
                return GetStyleProp("width");
            }
            set
            {
                SetStyleProp("width",value.ToString(CultureInfo.InvariantCulture) + "pt");
            }
        }
        /// <summary>
        /// The height in points
        /// </summary>
        public double Height
        {
            get
            {
                return GetStyleProp("height");
            }
            set
            {
                SetStyleProp("height", value.ToString(CultureInfo.InvariantCulture) + "pt");
            }
        }
        /// <summary>
        /// Margin Left in points
        /// </summary>
        public double Left
        {
            get
            {
                return GetStyleProp("left");
            }
            set
            {
                SetStyleProp("left", value.ToString(CultureInfo.InvariantCulture));
            }
        }
        /// <summary>
        /// Margin top in points
        /// </summary>
        public double Top
        {
            get
            {
                return GetStyleProp("top");
            }
            set
            {
                SetStyleProp("top", value.ToString(CultureInfo.InvariantCulture));
            }
        }
        /// <summary>
        /// The Title of the image
        /// </summary>
        public string Title
        {
            get
            {
                return GetXmlNodeString("v:imagedata/@o:title");
            }
            set
            {
                SetXmlNodeString("v:imagedata/@o:title",value);
            }
        }
        /// <summary>
        /// The image
        /// </summary>
        public Image Image
        {
            get
            {
                var pck = _worksheet._package.Package;
                if (pck.PartExists(ImageUri))
                {
                    var part = pck.GetPart(ImageUri);
                    return Image.FromStream(part.GetStream());
                }
                else
                {
                    return null;
                }
            }
        }
        internal Uri ImageUri
        {
            get;
            set;
        }
        internal string RelId
        {
            get
            {
                return GetXmlNodeString("v:imagedata/@o:relid");
            }
            set
            {
                SetXmlNodeString("v:imagedata/@o:relid",value);
            }
        }        
        /// <summary>
        /// Determines whether an image will be displayed in black and white
        /// </summary>
        public bool BiLevel
        {
            get
            {
                return GetXmlNodeString("v:imagedata/@bilevel")=="t";
            }
            set
            {
                if (value)
                {
                    SetXmlNodeString("v:imagedata/@bilevel", "t");
                }
                else
                {
                    DeleteNode("v:imagedata/@bilevel");
                }
            }
        }
        /// <summary>
        /// Determines whether a picture will be displayed in grayscale mode
        /// </summary>
        public bool GrayScale
        {
            get
            {
                return GetXmlNodeString("v:imagedata/@grayscale")=="t";
            }
            set
            {
                if (value)
                {
                    SetXmlNodeString("v:imagedata/@grayscale", "t");
                }
                else
                {
                    DeleteNode("v:imagedata/@grayscale");
                }
            }
        }
        /// <summary>
        /// Defines the intensity of all colors in an image
        /// Default value is 1
        /// </summary>
        public double Gain
        {
            get
            {
                string v = GetXmlNodeString("v:imagedata/@gain");
                return GetFracDT(v,1);
            }
            set
            {
                if (value < 0)
                {
                    throw (new ArgumentOutOfRangeException("Value must be positive"));
                }
                if (value == 1)
                {
                    DeleteNode("v:imagedata/@gamma");
                }
                else
                {
                    SetXmlNodeString("v:imagedata/@gain", value.ToString("#.0#", CultureInfo.InvariantCulture));
                }
            }
        }
        /// <summary>
        /// Defines the amount of contrast for an image
        /// Default value is 0;
        /// </summary>
        public double Gamma
        {
            get
            {
                string v = GetXmlNodeString("v:imagedata/@gamma");
                return GetFracDT(v,0);
            }
            set
            {
                if (value == 0) //Default
                {
                    DeleteNode("v:imagedata/@gamma");
                }
                else
                {
                    SetXmlNodeString("v:imagedata/@gamma", value.ToString("#.0#", CultureInfo.InvariantCulture));
                }
            }
        }
        /// <summary>
        /// Defines the intensity of black in an image
        /// Default value is 0
        /// </summary>
        public double BlackLevel
        {
            get
            {
                string v = GetXmlNodeString("v:imagedata/@blacklevel");
                return GetFracDT(v, 0);
            }
            set
            {
                if (value == 0)
                {
                    DeleteNode("v:imagedata/@blacklevel");
                }
                else
                {
                    SetXmlNodeString("v:imagedata/@blacklevel", value.ToString("#.0#", CultureInfo.InvariantCulture));
                }
            }
        }

        #region Private Methods
        private double GetFracDT(string v, double def)
        {
            double d;
            if (v.EndsWith("f"))
            {
                v = v.Substring(0, v.Length - 1);
                if (double.TryParse(v, out d))
                {
                    d /= 65535;
                }
                else
                {
                    d = def;
                }
            }
            else
            {
                if (!double.TryParse(v, out d))
                {
                    d = def;
                }
            }
            return d;
        }
        private void SetStyleProp(string propertyName, string value)
        {
            string style = GetXmlNodeString("@style");
            string newStyle = "";
            bool found = false;
            foreach (string prop in style.Split(';'))
            {
                string[] split = prop.Split(':');
                if (split[0] == propertyName)
                {
                    newStyle += propertyName + ":" + value + ";";
                    found = true;
                }
                else
                {
                    newStyle += prop + ";";
                }
            }
            if (!found)
            {
                newStyle += propertyName + ":" + value + ";";
            }
            SetXmlNodeString("@style", newStyle.Substring(0, newStyle.Length - 1));
        }
        private double GetStyleProp(string propertyName)
        {
            string style = GetXmlNodeString("@style");
            foreach (string prop in style.Split(';'))
            {
                string[] split = prop.Split(':');
                if (split[0] == propertyName && split.Length > 1)
                {
                    string value = split[1].EndsWith("pt") ? split[1].Substring(0, split[1].Length - 2) : split[1];
                    double ret;
                    if (double.TryParse(value, out ret))
                    {
                        return ret;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            return 0;
        }
        #endregion
    }
}
