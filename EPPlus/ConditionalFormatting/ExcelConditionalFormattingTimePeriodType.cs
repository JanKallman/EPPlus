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
  /// Functions related to the <see cref="ExcelConditionalFormattingTimePeriodType"/>
	/// </summary>
  internal static class ExcelConditionalFormattingTimePeriodType
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="type"></param>
		/// <returns></returns>
		public static string GetAttributeByType(
			eExcelConditionalFormattingTimePeriodType type)
		{
			switch (type)
			{
        case eExcelConditionalFormattingTimePeriodType.Last7Days:
          return ExcelConditionalFormattingConstants.TimePeriods.Last7Days;

        case eExcelConditionalFormattingTimePeriodType.LastMonth:
          return ExcelConditionalFormattingConstants.TimePeriods.LastMonth;

        case eExcelConditionalFormattingTimePeriodType.LastWeek:
          return ExcelConditionalFormattingConstants.TimePeriods.LastWeek;

        case eExcelConditionalFormattingTimePeriodType.NextMonth:
          return ExcelConditionalFormattingConstants.TimePeriods.NextMonth;

        case eExcelConditionalFormattingTimePeriodType.NextWeek:
          return ExcelConditionalFormattingConstants.TimePeriods.NextWeek;

        case eExcelConditionalFormattingTimePeriodType.ThisMonth:
          return ExcelConditionalFormattingConstants.TimePeriods.ThisMonth;

        case eExcelConditionalFormattingTimePeriodType.ThisWeek:
          return ExcelConditionalFormattingConstants.TimePeriods.ThisWeek;

        case eExcelConditionalFormattingTimePeriodType.Today:
          return ExcelConditionalFormattingConstants.TimePeriods.Today;

        case eExcelConditionalFormattingTimePeriodType.Tomorrow:
          return ExcelConditionalFormattingConstants.TimePeriods.Tomorrow;

        case eExcelConditionalFormattingTimePeriodType.Yesterday:
          return ExcelConditionalFormattingConstants.TimePeriods.Yesterday;
			}

			return string.Empty;
		}

    /// <summary>
    /// 
    /// </summary>
    /// <param name="attribute"></param>
    /// <returns></returns>
    public static eExcelConditionalFormattingTimePeriodType GetTypeByAttribute(
      string attribute)
    {
      switch (attribute)
      {
        case ExcelConditionalFormattingConstants.TimePeriods.Last7Days:
          return eExcelConditionalFormattingTimePeriodType.Last7Days;

        case ExcelConditionalFormattingConstants.TimePeriods.LastMonth:
          return eExcelConditionalFormattingTimePeriodType.LastMonth;

        case ExcelConditionalFormattingConstants.TimePeriods.LastWeek:
          return eExcelConditionalFormattingTimePeriodType.LastWeek;

        case ExcelConditionalFormattingConstants.TimePeriods.NextMonth:
          return eExcelConditionalFormattingTimePeriodType.NextMonth;

        case ExcelConditionalFormattingConstants.TimePeriods.NextWeek:
          return eExcelConditionalFormattingTimePeriodType.NextWeek;

        case ExcelConditionalFormattingConstants.TimePeriods.ThisMonth:
          return eExcelConditionalFormattingTimePeriodType.ThisMonth;

        case ExcelConditionalFormattingConstants.TimePeriods.ThisWeek:
          return eExcelConditionalFormattingTimePeriodType.ThisWeek;

        case ExcelConditionalFormattingConstants.TimePeriods.Today:
          return eExcelConditionalFormattingTimePeriodType.Today;

        case ExcelConditionalFormattingConstants.TimePeriods.Tomorrow:
          return eExcelConditionalFormattingTimePeriodType.Tomorrow;

        case ExcelConditionalFormattingConstants.TimePeriods.Yesterday:
          return eExcelConditionalFormattingTimePeriodType.Yesterday;
      }

      throw new Exception(
        ExcelConditionalFormattingConstants.Errors.UnexistentTimePeriodTypeAttribute);
    }
  }
}