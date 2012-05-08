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
 *  * Author							Change						      Date
 * ******************************************************************************
 * Eyal Seagull		    Conditional Formatting      2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// Factory class for ExcelConditionalFormatting.
	/// </summary>
	internal static class ExcelConditionalFormattingRuleFactory
	{
		public static ExcelConditionalFormattingRule Create(
			eExcelConditionalFormattingRuleType type,
      ExcelAddress address,
      int priority,
			ExcelWorksheet worksheet,
			XmlNode itemElementNode)
		{
			Require.Argument(type);
      Require.Argument(address).IsNotNull("address");
      Require.Argument(priority).IsInRange(1, int.MaxValue, "priority");
			Require.Argument(worksheet).IsNotNull("worksheet");
			
			// According the conditional formatting rule type
			switch (type)
			{
        case eExcelConditionalFormattingRuleType.AboveAverage:
          return new ExcelConditionalFormattingAboveAverage(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
          return new ExcelConditionalFormattingAboveOrEqualAverage(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BelowAverage:
          return new ExcelConditionalFormattingBelowAverage(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
          return new ExcelConditionalFormattingBelowOrEqualAverage(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.AboveStdDev:
          return new ExcelConditionalFormattingAboveStdDev(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BelowStdDev:
          return new ExcelConditionalFormattingBelowStdDev(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Bottom:
          return new ExcelConditionalFormattingBottom(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BottomPercent:
          return new ExcelConditionalFormattingBottomPercent(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Top:
          return new ExcelConditionalFormattingTop(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.TopPercent:
          return new ExcelConditionalFormattingTopPercent(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Last7Days:
          return new ExcelConditionalFormattingLast7Days(
            address,
            priority,
            worksheet,
            itemElementNode);


        case eExcelConditionalFormattingRuleType.LastMonth:
          return new ExcelConditionalFormattingLastMonth(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.LastWeek:
          return new ExcelConditionalFormattingLastWeek(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NextMonth:
          return new ExcelConditionalFormattingNextMonth(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NextWeek:
          return new ExcelConditionalFormattingNextWeek(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ThisMonth:
          return new ExcelConditionalFormattingThisMonth(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ThisWeek:
          return new ExcelConditionalFormattingThisWeek(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Today:
          return new ExcelConditionalFormattingToday(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Tomorrow:
          return new ExcelConditionalFormattingTomorrow(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Yesterday:
          return new ExcelConditionalFormattingYesterday(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.BeginsWith:
          return new ExcelConditionalFormattingBeginsWith(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Between:
          return new ExcelConditionalFormattingBetween(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ContainsBlanks:
          return new ExcelConditionalFormattingContainsBlanks(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ContainsErrors:
          return new ExcelConditionalFormattingContainsErrors(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ContainsText:
          return new ExcelConditionalFormattingContainsText(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.DuplicateValues:
          return new ExcelConditionalFormattingDuplicateValues(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.EndsWith:
          return new ExcelConditionalFormattingEndsWith(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Equal:
          return new ExcelConditionalFormattingEqual(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.Expression:
          return new ExcelConditionalFormattingExpression(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.GreaterThan:
          return new ExcelConditionalFormattingGreaterThan(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
          return new ExcelConditionalFormattingGreaterThanOrEqual(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.LessThan:
          return new ExcelConditionalFormattingLessThan(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.LessThanOrEqual:
          return new ExcelConditionalFormattingLessThanOrEqual(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotBetween:
          return new ExcelConditionalFormattingNotBetween(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotContainsBlanks:
          return new ExcelConditionalFormattingNotContainsBlanks(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotContainsErrors:
          return new ExcelConditionalFormattingNotContainsErrors(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotContainsText:
          return new ExcelConditionalFormattingNotContainsText(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.NotEqual:
          return new ExcelConditionalFormattingNotEqual(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.UniqueValues:
          return new ExcelConditionalFormattingUniqueValues(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.ThreeColorScale:
          return new ExcelConditionalFormattingThreeColorScale(
            address,
            priority,
            worksheet,
            itemElementNode);

        case eExcelConditionalFormattingRuleType.TwoColorScale:
          return new ExcelConditionalFormattingTwoColorScale(
            address,
            priority,
						worksheet,
						itemElementNode);

        //TODO: Add DataBar and IconSet
			}

			throw new InvalidOperationException(
        string.Format(
          ExcelConditionalFormattingConstants.Errors.NonSupportedRuleType,
          type.ToString()));
		}
	}
}