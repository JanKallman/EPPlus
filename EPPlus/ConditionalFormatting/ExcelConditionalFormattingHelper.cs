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
 * Eyal Seagull        Added       		  2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.Utils;
using System.Drawing;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// Conditional formatting helper
  /// </summary>
  internal static class ExcelConditionalFormattingHelper
  {
    /// <summary>
    /// Check and fix an address (string address)
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public static string CheckAndFixRangeAddress(
      string address)
    {
      if (address.Contains(','))
      {
        throw new FormatException(
          ExcelConditionalFormattingConstants.Errors.CommaSeparatedAddresses);
      }

      address = address.ToUpper();

      if (Regex.IsMatch(address, @"[A-Z]+:[A-Z]+"))
      {
        address = AddressUtility.ParseEntireColumnSelections(address);
      }

      return address;
    }

    /// <summary>
    /// Convert a color code to Color Object
    /// </summary>
    /// <param name="colorCode">Color Code (Ex. "#FFB43C53" or "FFB43C53")</param>
    /// <returns></returns>
    public static Color ConvertFromColorCode(
      string colorCode)
    {
      try
      {
        return Color.FromArgb(Int32.Parse(colorCode.Replace("#", ""), NumberStyles.HexNumber));
      }
      catch
      {
        // Assume white is the default color (instead of giving an error)
        return Color.White;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static string GetAttributeString(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return (value == null) ? string.Empty : value;
      }
      catch
      {
        return string.Empty;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static int GetAttributeInt(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return int.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
      }
      catch
      {
        return int.MinValue;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static int? GetAttributeIntNullable(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return int.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
      }
      catch
      {
        return null;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static bool GetAttributeBool(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return (value == "1" || value == "-1" || value.ToUpper() == "TRUE");
      }
      catch
      {
        return false;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static bool? GetAttributeBoolNullable(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return (value == "1" || value == "-1" || value.ToUpper() == "TRUE");
      }
      catch
      {
        return null;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static double GetAttributeDouble(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return double.Parse(value, NumberStyles.Number, CultureInfo.InvariantCulture);
      }
      catch
      {
        return double.NaN;
      }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="node"></param>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static decimal GetAttributeDecimal(
      XmlNode node,
      string attribute)
    {
      try
      {
        var value = node.Attributes[attribute].Value;
        return decimal.Parse(value, NumberStyles.Any, CultureInfo.InvariantCulture);
      }
      catch
      {
        return decimal.MinValue;
      }
    }

    /// <summary>
    /// Encode to XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
    /// </summary>
    /// <param name="s"></param>
    /// <returns></returns>
    public static string EncodeXML(
      this string s)
    {
      return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
    }

    /// <summary>
    /// Decode from XML (special characteres: &apos; &quot; &gt; &lt; &amp;)
    /// </summary>
    /// <param name="s"></param>
    /// <returns></returns>
    public static string DecodeXML(
      this string s)
    {
      return s.Replace("'", "&apos;").Replace("\"", "&quot;").Replace(">", "&gt;").Replace("<", "&lt;").Replace("&", "&amp;");
    }
  }
}