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
 * Author          Change						                  Date
 * ******************************************************************************
 * Eyal Seagull    Conditional Formatting Adaption    2012-04-17
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
  /// Functions related to the <see cref="ExcelConditionalFormattingOperatorType"/>
	/// </summary>
  internal static class ExcelConditionalFormattingOperatorType
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="type"></param>
		/// <returns></returns>
		internal static string GetAttributeByType(
			eExcelConditionalFormattingOperatorType type)
		{
			switch (type)
			{
        case eExcelConditionalFormattingOperatorType.BeginsWith:
          return ExcelConditionalFormattingConstants.Operators.BeginsWith;

        case eExcelConditionalFormattingOperatorType.Between:
          return ExcelConditionalFormattingConstants.Operators.Between;

        case eExcelConditionalFormattingOperatorType.ContainsText:
          return ExcelConditionalFormattingConstants.Operators.ContainsText;

        case eExcelConditionalFormattingOperatorType.EndsWith:
          return ExcelConditionalFormattingConstants.Operators.EndsWith;

        case eExcelConditionalFormattingOperatorType.Equal:
          return ExcelConditionalFormattingConstants.Operators.Equal;

        case eExcelConditionalFormattingOperatorType.GreaterThan:
          return ExcelConditionalFormattingConstants.Operators.GreaterThan;

        case eExcelConditionalFormattingOperatorType.GreaterThanOrEqual:
          return ExcelConditionalFormattingConstants.Operators.GreaterThanOrEqual;

        case eExcelConditionalFormattingOperatorType.LessThan:
          return ExcelConditionalFormattingConstants.Operators.LessThan;

        case eExcelConditionalFormattingOperatorType.LessThanOrEqual:
          return ExcelConditionalFormattingConstants.Operators.LessThanOrEqual;

        case eExcelConditionalFormattingOperatorType.NotBetween:
          return ExcelConditionalFormattingConstants.Operators.NotBetween;

        case eExcelConditionalFormattingOperatorType.NotContains:
          return ExcelConditionalFormattingConstants.Operators.NotContains;

        case eExcelConditionalFormattingOperatorType.NotEqual:
          return ExcelConditionalFormattingConstants.Operators.NotEqual;
			}

			return string.Empty;
		}

    /// <summary>
    /// 
    /// </summary>
    /// param name="attribute"
    /// <returns></returns>
    internal static eExcelConditionalFormattingOperatorType GetTypeByAttribute(
      string attribute)
    {
      switch (attribute)
      {
        case ExcelConditionalFormattingConstants.Operators.BeginsWith:
          return eExcelConditionalFormattingOperatorType.BeginsWith;

        case ExcelConditionalFormattingConstants.Operators.Between:
          return eExcelConditionalFormattingOperatorType.Between;

        case ExcelConditionalFormattingConstants.Operators.ContainsText:
          return eExcelConditionalFormattingOperatorType.ContainsText;

        case ExcelConditionalFormattingConstants.Operators.EndsWith:
          return eExcelConditionalFormattingOperatorType.EndsWith;

        case ExcelConditionalFormattingConstants.Operators.Equal:
          return eExcelConditionalFormattingOperatorType.Equal;

        case ExcelConditionalFormattingConstants.Operators.GreaterThan:
          return eExcelConditionalFormattingOperatorType.GreaterThan;

        case ExcelConditionalFormattingConstants.Operators.GreaterThanOrEqual:
          return eExcelConditionalFormattingOperatorType.GreaterThanOrEqual;

        case ExcelConditionalFormattingConstants.Operators.LessThan:
          return eExcelConditionalFormattingOperatorType.LessThan;

        case ExcelConditionalFormattingConstants.Operators.LessThanOrEqual:
          return eExcelConditionalFormattingOperatorType.LessThanOrEqual;

        case ExcelConditionalFormattingConstants.Operators.NotBetween:
          return eExcelConditionalFormattingOperatorType.NotBetween;

        case ExcelConditionalFormattingConstants.Operators.NotContains:
          return eExcelConditionalFormattingOperatorType.NotContains;

        case ExcelConditionalFormattingConstants.Operators.NotEqual:
          return eExcelConditionalFormattingOperatorType.NotEqual;
      }

      throw new Exception(
        ExcelConditionalFormattingConstants.Errors.UnexistentOperatorTypeAttribute);
    }
  }
}