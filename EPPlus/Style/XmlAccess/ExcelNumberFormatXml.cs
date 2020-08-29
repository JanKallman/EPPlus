﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Globalization;
using System.Text.RegularExpressions;
namespace OfficeOpenXml.Style.XmlAccess
{
    /// <summary>
    /// Xml access class for number formats
    /// </summary>
    public sealed class ExcelNumberFormatXml : StyleXmlHelper
    {
        internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager) : base(nameSpaceManager)
        {
            
        }        
        internal ExcelNumberFormatXml(XmlNamespaceManager nameSpaceManager, bool buildIn): base(nameSpaceManager)
        {
            BuildIn = buildIn;
        }
        internal ExcelNumberFormatXml(XmlNamespaceManager nsm, XmlNode topNode) :
            base(nsm, topNode)
        {
            _numFmtId = GetXmlNodeInt("@numFmtId");
            _format = GetXmlNodeString("@formatCode");
        }
        public bool BuildIn { get; private set; }
        int _numFmtId;
//        const string idPath = "@numFmtId";
        /// <summary>
        /// Id for number format
        /// 
        /// Build in ID's
        /// 
        /// 0   General 
        /// 1   0 
        /// 2   0.00 
        /// 3   #,##0 
        /// 4   #,##0.00 
        /// 9   0% 
        /// 10  0.00% 
        /// 11  0.00E+00 
        /// 12  # ?/? 
        /// 13  # ??/?? 
        /// 14  mm-dd-yy 
        /// 15  d-mmm-yy 
        /// 16  d-mmm 
        /// 17  mmm-yy 
        /// 18  h:mm AM/PM 
        /// 19  h:mm:ss AM/PM 
        /// 20  h:mm 
        /// 21  h:mm:ss 
        /// 22  m/d/yy h:mm 
        /// 37  #,##0 ;(#,##0) 
        /// 38  #,##0 ;[Red](#,##0) 
        /// 39  #,##0.00;(#,##0.00) 
        /// 40  #,##0.00;[Red](#,##0.00) 
        /// 45  mm:ss 
        /// 46  [h]:mm:ss 
        /// 47  mmss.0 
        /// 48  ##0.0E+0 
        /// 49  @
        /// </summary>            
        public int NumFmtId
        {
            get
            {
                return _numFmtId;
            }
            set
            {
                _numFmtId = value;
            }
        }
        internal override string Id
        {
            get
            {
                return _format;
            }
        }
        const string fmtPath = "@formatCode";
        string _format = string.Empty;
        public string Format
        {
            get
            {
                return _format;
            }
            set
            {
                _numFmtId = ExcelNumberFormat.GetFromBuildIdFromFormat(value);
                _format = value;
            }
        }
        internal string GetNewID(int NumFmtId, string Format)
        {
            
            if (NumFmtId < 0)
            {
                NumFmtId = ExcelNumberFormat.GetFromBuildIdFromFormat(Format);                
            }
            return NumFmtId.ToString();
        }

        internal static void AddBuildIn(XmlNamespaceManager NameSpaceManager, ExcelStyleCollection<ExcelNumberFormatXml> NumberFormats)
        {
            NumberFormats.Add("General",new ExcelNumberFormatXml(NameSpaceManager,true){NumFmtId=0,Format="General"});
            NumberFormats.Add("0", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 1, Format = "0" });
            NumberFormats.Add("0.00", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 2, Format = "0.00" });
            NumberFormats.Add("#,##0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 3, Format = "#,##0" });
            NumberFormats.Add("#,##0.00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 4, Format = "#,##0.00" });
            NumberFormats.Add("0%", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 9, Format = "0%" });
            NumberFormats.Add("0.00%", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 10, Format = "0.00%" });
            NumberFormats.Add("0.00E+00", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 11, Format = "0.00E+00" });
            NumberFormats.Add("# ?/?", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 12, Format = "# ?/?" });
            NumberFormats.Add("# ??/??", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 13, Format = "# ??/??" });
            NumberFormats.Add("mm-dd-yy", new ExcelNumberFormatXml(NameSpaceManager,true) { NumFmtId = 14, Format = "mm-dd-yy" });
            NumberFormats.Add("d-mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 15, Format = "d-mmm-yy" });
            NumberFormats.Add("d-mmm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 16, Format = "d-mmm" });
            NumberFormats.Add("mmm-yy", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 17, Format = "mmm-yy" });
            NumberFormats.Add("h:mm AM/PM", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 18, Format = "h:mm AM/PM" });
            NumberFormats.Add("h:mm:ss AM/PM", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 19, Format = "h:mm:ss AM/PM" });
            NumberFormats.Add("h:mm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 20, Format = "h:mm" });
            NumberFormats.Add("h:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 21, Format = "h:mm:ss" });
            NumberFormats.Add("m/d/yy h:mm", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 22, Format = "m/d/yy h:mm" });
            NumberFormats.Add("#,##0 ;(#,##0)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 37, Format = "#,##0 ;(#,##0)" });
            NumberFormats.Add("#,##0 ;[Red](#,##0)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 38, Format = "#,##0 ;[Red](#,##0)" });
            NumberFormats.Add("#,##0.00;(#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 39, Format = "#,##0.00;(#,##0.00)" });
            NumberFormats.Add("#,##0.00;[Red](#,##0.00)", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 40, Format = "#,##0.00;[Red](#,##0.00)" });
            NumberFormats.Add("mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 45, Format = "mm:ss" });
            NumberFormats.Add("[h]:mm:ss", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 46, Format = "[h]:mm:ss" });
            NumberFormats.Add("mmss.0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 47, Format = "mmss.0" });
            NumberFormats.Add("##0.0", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 48, Format = "##0.0" });
            NumberFormats.Add("@", new ExcelNumberFormatXml(NameSpaceManager, true) { NumFmtId = 49, Format = "@" });

            NumberFormats.NextId = 164; //Start for custom formats.
        }

        internal override XmlNode CreateXmlNode(XmlNode topNode)
        {
            TopNode = topNode;
            SetXmlNodeString("@numFmtId", NumFmtId.ToString());
            SetXmlNodeString("@formatCode", Format);
            return TopNode;
        }

        internal enum eFormatType
        {
            Unknown = 0,
            Number = 1,
            DateTime = 2,
        }
        ExcelFormatTranslator _translator = null;
        internal ExcelFormatTranslator FormatTranslator
        {
            get
            {
                if (_translator == null)
                {
                    _translator = new ExcelFormatTranslator(Format, NumFmtId);
                }
                return _translator;
            }
        }
        #region Excel --> .Net Format
        internal class ExcelFormatTranslator
        {
            internal enum eSystemDateFormat
            {
                None,
                SystemLongDate,
                SystemLongTime,
                Conditional,
                SystemShortDate,
            }
            internal ExcelFormatTranslator(string format, int numFmtID)
            {
                if (numFmtID == 14)
                {
                    NetFormat = NetFormatForWidth = "";
                    NetTextFormat = NetTextFormatForWidth = "";
                    SpecialDateFormat = eSystemDateFormat.SystemShortDate;
                    DataType = eFormatType.DateTime;
                }
                else if (format.Equals("general",StringComparison.OrdinalIgnoreCase))
                {
                    NetFormat = NetFormatForWidth = "0.#####";
                    NetTextFormat = NetTextFormatForWidth = "";
                    DataType = eFormatType.Number;
                }
                else
                {
                    ToNetFormat(format, false);
                    ToNetFormat(format, true);
                }                
            }
            internal string NetTextFormat { get; private set; }
            internal string NetFormat { get; private set; }
            CultureInfo _ci = null;
            internal CultureInfo Culture
            {
                get
                {
                    if (_ci == null)
                    {
                        return CultureInfo.CurrentCulture;
                    }
                    return _ci;
                }
                private set
                {
                    _ci = value;
                }
            }
            internal eFormatType DataType { get; private set; }
            internal string NetTextFormatForWidth { get; private set; }
            internal string NetFormatForWidth { get; private set; }
            internal string FractionFormat { get; private set; }
            internal eSystemDateFormat SpecialDateFormat { get; private set; }
            private void ToNetFormat(string ExcelFormat, bool forColWidth)
            {
                DataType = eFormatType.Unknown;
                int secCount = 0;
                bool isText = false;
                bool isBracket = false;
                string bracketText = "";
                bool prevBslsh = false;
                bool useMinute = false;
                bool prevUnderScore = false;
                bool ignoreNext = false;
                int fractionPos = -1;
                bool containsAmPm = ExcelFormat.Contains("AM/PM");
                List<int> lstDec=new List<int>();
                StringBuilder sb = new StringBuilder();
                Culture = null;
                var format = "";
                var text = "";
                char clc;

                if (containsAmPm)
                {
                    ExcelFormat = Regex.Replace(ExcelFormat, "AM/PM", "");
                    DataType = eFormatType.DateTime;
                }

                for (int pos = 0; pos < ExcelFormat.Length; pos++)
                {
                    char c = ExcelFormat[pos];
                    if (c == '"')
                    {
                        isText = !isText;
                    }
                    else
                    {
                        if (ignoreNext)
                        {
                            ignoreNext = false;
                            continue;
                        }
                        else if (isText && !isBracket)
                        {
                            sb.Append(c);
                        }
                        else if (isBracket)
                        {
                            if (c == ']')
                            {
                                isBracket = false;
                                if (bracketText[0] == '$')  //Local Info
                                {
                                    string[] li = Regex.Split(bracketText, "-");
                                    if (li[0].Length > 1)
                                    {
                                        sb.Append("\"" + li[0].Substring(1, li[0].Length - 1) + "\"");     //Currency symbol
                                    }
                                    if (li.Length > 1)
                                    {
                                        if (li[1].Equals("f800", StringComparison.OrdinalIgnoreCase))
                                        {
                                            SpecialDateFormat=eSystemDateFormat.SystemLongDate;
                                        }
                                        else if (li[1].Equals("f400", StringComparison.OrdinalIgnoreCase))
                                        {
                                            SpecialDateFormat = eSystemDateFormat.SystemLongTime;
                                        }
                                        else
                                        {
                                            var num = int.Parse(li[1], NumberStyles.HexNumber);
                                            try
                                            {
                                                Culture = CultureInfo.GetCultureInfo(num & 0xFFFF);
                                            }
                                            catch
                                            {
                                                Culture = null;
                                            }
                                        }
                                    }
                                }
                                else if(bracketText.StartsWith("<") ||
                                        bracketText.StartsWith(">") ||
                                        bracketText.StartsWith("=")) //Conditional
                                {
                                    SpecialDateFormat = eSystemDateFormat.Conditional;
                                }
                                else 
                                {
                                    sb.Append(bracketText);
                                    SpecialDateFormat = eSystemDateFormat.Conditional;
                                }
                            }
                            else
                            {
                                bracketText += c;
                            }
                        }
                        else if (prevUnderScore)
                        {
                            if (forColWidth)
                            {
                                sb.AppendFormat("\"{0}\"", c);
                            }
                            prevUnderScore = false;
                        }
                        else
                        {
                            if (c == ';') //We use first part (for positive only at this stage)
                            {
                                secCount++;
                                if (DataType == eFormatType.DateTime || secCount == 3)
                                {
                                    //Add qoutes
                                    if (DataType == eFormatType.DateTime) SetDecimal(lstDec, sb); //Remove?
                                    lstDec = new List<int>();
                                    format = sb.ToString();
                                    sb = new StringBuilder();
                                }
                                else
                                {
                                    sb.Append(c);
                                }
                            }
                            else
                            {
                                clc = c.ToString().ToLower(CultureInfo.InvariantCulture)[0];  //Lowercase character
                                //Set the datetype
                                if (DataType == eFormatType.Unknown)
                                {
                                    if (c == '0' || c == '#' || c == '.')
                                    {
                                        DataType = eFormatType.Number;
                                    }
                                    else if (clc == 'y' || clc == 'm' || clc == 'd' || clc == 'h' || clc == 'm' || clc == 's')
                                    {
                                        DataType = eFormatType.DateTime;
                                    }
                                }

                                if (prevBslsh)
                                {
                                    if (c == '.' || c == ',')
                                    {
                                        sb.Append('\\');
                                    }                                    
                                    sb.Append(c);
                                    prevBslsh = false;
                                }
                                else if (c == '[')
                                {
                                    bracketText = "";
                                    isBracket = true;
                                }
                                else if (c == '\\')
                                {
                                    prevBslsh = true;
                                }
                                else if (c == '0' ||
                                    c == '#' ||
                                    c == '.' ||
                                    c == ',' ||
                                    c == '%' ||
                                    clc == 'd' ||
                                    clc == 's')
                                {
                                    sb.Append(c);
                                    if(c=='.')
                                    {
                                        lstDec.Add(sb.Length - 1);
                                    }
                                }
                                else if (clc == 'h')
                                {
                                    if (containsAmPm)
                                    {
                                        sb.Append('h'); ;
                                    }
                                    else
                                    {
                                        sb.Append('H');
                                    }
                                    useMinute = true;
                                }
                                else if (clc == 'm')
                                {
                                    if (useMinute)
                                    {
                                        sb.Append('m');
                                    }
                                    else
                                    {
                                        sb.Append('M');
                                    }
                                }
                                else if (c == '_') //Skip next but use for alignment
                                {
                                    prevUnderScore = true;
                                }
                                else if (c == '?')
                                {
                                    sb.Append(' ');
                                }
                                else if (c == '/')
                                {
                                    if (DataType == eFormatType.Number)
                                    {
                                        fractionPos = sb.Length;
                                        int startPos = pos - 1;
                                        while (startPos >= 0 &&
                                                (ExcelFormat[startPos] == '?' ||
                                                ExcelFormat[startPos] == '#' ||
                                                ExcelFormat[startPos] == '0'))
                                        {
                                            startPos--;
                                        }

                                        if (startPos > 0)  //RemovePart
                                            sb.Remove(sb.Length-(pos-startPos-1),(pos-startPos-1)) ;

                                        int endPos = pos + 1;
                                        while (endPos < ExcelFormat.Length &&
                                                (ExcelFormat[endPos] == '?' ||
                                                ExcelFormat[endPos] == '#' ||
                                                (ExcelFormat[endPos] >= '0' && ExcelFormat[endPos]<= '9')))
                                        {
                                            endPos++;
                                        }
                                        pos = endPos;
                                        if (FractionFormat != "")
                                        {
                                            FractionFormat = ExcelFormat.Substring(startPos+1, endPos - startPos-1);
                                        }
                                        sb.Append('?'); //Will be replaced later on by the fraction
                                    }
                                    else
                                    {
                                        sb.Append('/');
                                    }
                                }
                                else if (c == '*')
                                {
                                    //repeat char--> ignore
                                    ignoreNext = true;
                                }
                                else if (c == '@')
                                {
                                    sb.Append("{0}");
                                }
                                else
                                {
                                    sb.Append(c);
                                }
                            }
                        }
                    }
                }

                //Add qoutes
                if (DataType == eFormatType.DateTime) SetDecimal(lstDec, sb); //Remove?

                // AM/PM format
                if (containsAmPm)
                {
                    format += "tt";
                }


                if (format == "")
                    format = sb.ToString();
                else
                    text = sb.ToString();

                if (forColWidth)
                {
                    NetFormatForWidth = format;
                    NetTextFormatForWidth = text;
                }
                else
                {
                    NetFormat = format;
                    NetTextFormat = text;
                }
                if (Culture == null)
                {
                    Culture = CultureInfo.CurrentCulture;
                }
            }

            private static void SetDecimal(List<int> lstDec, StringBuilder sb)
            {
                if (lstDec.Count > 1)
                {
                    for (int i = lstDec.Count - 1; i >= 0; i--)
                    {
                        sb.Insert(lstDec[i] + 1, '\'');
                        sb.Insert(lstDec[i], '\'');
                    }
                }
            }

            internal string FormatFraction(double d)
            {
                int numerator, denomerator;

                int intPart = (int)d;

                string[] fmt = FractionFormat.Split('/');

                int fixedDenominator;
                if (!int.TryParse(fmt[1], out fixedDenominator))
                {
                    fixedDenominator = 0;
                }
                
                if (d == 0 || double.IsNaN(d))
                {
                    if (fmt[0].Trim() == "" && fmt[1].Trim() == "")
                    {
                        return new string(' ', FractionFormat.Length);
                    }
                    else
                    {
                        return 0.ToString(fmt[0]) + "/" + 1.ToString(fmt[0]);
                    }
                }

                int maxDigits = fmt[1].Length;
                string sign = d < 0 ? "-" : "";
                if (fixedDenominator == 0)
                {
                    List<double> numerators = new List<double>() { 1, 0 };
                    List<double> denominators = new List<double>() { 0, 1 };

                    if (maxDigits < 1 && maxDigits > 12)
                    {
                        throw (new ArgumentException("Number of digits out of range (1-12)"));
                    }

                    int maxNum = 0;
                    for (int i = 0; i < maxDigits; i++)
                    {
                        maxNum += 9 * (int)(Math.Pow((double)10, (double)i));
                    }

                    double divRes = 1 / ((double)Math.Abs(d) - intPart);
                    double result, prevResult = double.NaN;
                    int listPos = 2, index = 1;
                    while (true)
                    {
                        index++;
                        double intDivRes = Math.Floor(divRes);
                        numerators.Add((intDivRes * numerators[index - 1] + numerators[index - 2]));
                        if (numerators[index] > maxNum)
                        {
                            break;
                        }

                        denominators.Add((intDivRes * denominators[index - 1] + denominators[index - 2]));

                        result = numerators[index] / denominators[index];
                        if (denominators[index] > maxNum)
                        {
                            break;
                        }
                        listPos = index;

                        if (result == prevResult) break;

                        if (result == d) break;

                        prevResult = result;

                        divRes = 1 / (divRes - intDivRes);  //Rest
                    }
                    
                    numerator = (int)numerators[listPos];
                    denomerator = (int)denominators[listPos];
                }
                else
                {
                    numerator = (int)Math.Round((d - intPart) / (1D / fixedDenominator), 0);
                    denomerator = fixedDenominator;
                }
                if (numerator == denomerator || numerator==0)
                {
                    if(numerator == denomerator) intPart++;
                    return sign + intPart.ToString(NetFormat).Replace("?", new string(' ', FractionFormat.Length));
                }
                else if (intPart == 0)
                {
                    return sign + FmtInt(numerator, fmt[0]) + "/" + FmtInt(denomerator, fmt[1]);
                }
                else
                {
                    return sign + intPart.ToString(NetFormat).Replace("?", FmtInt(numerator, fmt[0]) + "/" + FmtInt(denomerator, fmt[1]));
                }
            }

            private string FmtInt(double value, string format)
            {
                string v = value.ToString("#");
                string pad = "";
                if (v.Length < format.Length)
                {
                    for (int i = format.Length - v.Length-1; i >= 0; i--)
                    {
                        if (format[i] == '?')
                        {
                            pad += " ";
                        }
                        else if (format[i] == ' ')
                        {
                            pad += "0";
                        }
                    }
                }
                return pad + v;
            }
        }
#endregion
    }
}
