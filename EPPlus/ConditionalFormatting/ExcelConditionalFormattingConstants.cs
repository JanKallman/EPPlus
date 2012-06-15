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
 * Eyal Seagull    Conditional Formatting Adaption    2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// The conditional formatting constants
  /// </summary>
  internal static class ExcelConditionalFormattingConstants
  {
    #region Errors
    internal class Errors
    {
      internal const string CommaSeparatedAddresses = @"Multiple addresses may not be commaseparated, use space instead";
      internal const string InvalidCfruleObject = @"The supplied item must inherit OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormattingRule";
      internal const string InvalidConditionalFormattingObject = @"The supplied item must inherit OfficeOpenXml.ConditionalFormatting.ExcelConditionalFormatting";
      internal const string InvalidPriority = @"Invalid priority number. Must be bigger than zero";
      internal const string InvalidRemoveRuleOperation = @"Invalid remove rule operation";
      internal const string MissingCfvoNode = @"Missing 'cfvo' node in Conditional Formatting";
      internal const string MissingCfvoParentNode = @"Missing 'cfvo' parent node in Conditional Formatting";
      internal const string MissingConditionalFormattingNode = @"Missing 'conditionalFormatting' node in Conditional Formatting";
      internal const string MissingItemRuleList = @"Missing item with address '{0}' in Conditional Formatting Rule List";
      internal const string MissingPriorityAttribute = @"Missing 'priority' attribute in Conditional Formatting Rule";
      internal const string MissingRuleType = @"Missing eExcelConditionalFormattingRuleType Type in Conditional Formatting";
      internal const string MissingSqrefAttribute = @"Missing 'sqref' attribute in Conditional Formatting";
      internal const string MissingTypeAttribute = @"Missing 'type' attribute in Conditional Formatting Rule";
      internal const string MissingWorksheetNode = @"Missing 'worksheet' node";
      internal const string NonSupportedRuleType = @"Non supported conditionalFormattingType: {0}";
      internal const string UnexistentCfvoTypeAttribute = @"Unexistent eExcelConditionalFormattingValueObjectType attribute in Conditional Formatting";
      internal const string UnexistentOperatorTypeAttribute = @"Unexistent eExcelConditionalFormattingOperatorType attribute in Conditional Formatting";
      internal const string UnexistentTimePeriodTypeAttribute = @"Unexistent eExcelConditionalFormattingTimePeriodType attribute in Conditional Formatting";
      internal const string UnexpectedRuleTypeAttribute = @"Unexpected eExcelConditionalFormattingRuleType attribute in Conditional Formatting Rule";
      internal const string UnexpectedRuleTypeName = @"Unexpected eExcelConditionalFormattingRuleType TypeName in Conditional Formatting Rule";
      internal const string WrongNumberCfvoColorNodes = @"Wrong number of 'cfvo'/'color' nodes in Conditional Formatting Rule";
    }
    #endregion Errors

    #region Nodes
    internal class Nodes
    {
      internal const string Worksheet = "worksheet";
      internal const string ConditionalFormatting = "conditionalFormatting";
      internal const string CfRule = "cfRule";
      internal const string ColorScale = "colorScale";
      internal const string Cfvo = "cfvo";
      internal const string Color = "color";
      internal const string DataBar = "dataBar";
      internal const string IconSet = "iconSet";
      internal const string Formula = "formula";
    }
    #endregion Nodes

    #region Attributes
    internal class Attributes
    {
      internal const string AboveAverage = "aboveAverage";
      internal const string Bottom = "bottom";
      internal const string DxfId = "dxfId";
      internal const string EqualAverage = "equalAverage";
      internal const string IconSet = "iconSet";
      internal const string Operator = "operator";
      internal const string Percent = "percent";
      internal const string Priority = "priority";
      internal const string Rank = "rank";
      internal const string Reverse = "reverse";
      internal const string Rgb = "rgb";
      internal const string ShowValue = "showValue";
      internal const string Sqref = "sqref";
      internal const string StdDev = "stdDev";
      internal const string StopIfTrue = "stopIfTrue";
      internal const string Text = "text";
      internal const string Theme = "theme";
      internal const string TimePeriod = "timePeriod";
      internal const string Tint = "tint";
      internal const string Type = "type";
      internal const string Val = "val";
    }
    #endregion Attributes

    #region XML Paths
    internal class Paths
    {
      // Main node and attributes
      internal const string Worksheet = "d:" + Nodes.Worksheet;

      // <conditionalFormatting> §18.3.1.18 node
      // can appear more than once in a worksheet
      internal const string ConditionalFormatting = "d:" + Nodes.ConditionalFormatting;

      // <cfRule> §18.3.1.10 node
      // can appear more than once in a <conditionalFormatting>
      internal const string CfRule = "d:" + Nodes.CfRule;

      // <colorScale> §18.3.1.16 node
      internal const string ColorScale = "d:" + Nodes.ColorScale;

      // <cfvo> §18.3.1.11 node
      internal const string Cfvo = "d:" + Nodes.Cfvo;

      // <color> §18.3.1.15 node
      internal const string Color = "d:" + Nodes.Color;

      // <dataBar> §18.3.1.28 node
      internal const string DataBar = "d:" + Nodes.DataBar;

      // <iconSet> §18.3.1.49 node
      internal const string IconSet = "d:" + Nodes.IconSet;

      // <formula> §18.3.1.43 node
      internal const string Formula = "d:" + Nodes.Formula;

      // Attributes (for all the nodes)
      internal const string AboveAverageAttribute = "@" + Attributes.AboveAverage;
      internal const string BottomAttribute = "@" + Attributes.Bottom;
      internal const string DxfIdAttribute = "@" + Attributes.DxfId;
      internal const string EqualAverageAttribute = "@" + Attributes.EqualAverage;
      internal const string IconSetAttribute = "@" + Attributes.IconSet;
      internal const string OperatorAttribute = "@" + Attributes.Operator;
      internal const string PercentAttribute = "@" + Attributes.Percent;
      internal const string PriorityAttribute = "@" + Attributes.Priority;
      internal const string RankAttribute = "@" + Attributes.Rank;
      internal const string ReverseAttribute = "@" + Attributes.Reverse;
      internal const string RgbAttribute = "@" + Attributes.Rgb;
      internal const string ShowValueAttribute = "@" + Attributes.ShowValue;
      internal const string SqrefAttribute = "@" + Attributes.Sqref;
      internal const string StdDevAttribute = "@" + Attributes.StdDev;
      internal const string StopIfTrueAttribute = "@" + Attributes.StopIfTrue;
      internal const string TextAttribute = "@" + Attributes.Text;
      internal const string ThemeAttribute = "@" + Attributes.Theme;
      internal const string TimePeriodAttribute = "@" + Attributes.TimePeriod;
      internal const string TintAttribute = "@" + Attributes.Tint;
      internal const string TypeAttribute = "@" + Attributes.Type;
      internal const string ValAttribute = "@" + Attributes.Val;
    }
    #endregion XML Paths

    #region Rule Type ST_CfType §18.18.12 (with small EPPlus changes)
    internal class RuleType
    {
      internal const string AboveAverage = "aboveAverage";
      internal const string BeginsWith = "beginsWith";
      internal const string CellIs = "cellIs";
      internal const string ColorScale = "colorScale";
      internal const string ContainsBlanks = "containsBlanks";
      internal const string ContainsErrors = "containsErrors";
      internal const string ContainsText = "containsText";
      internal const string DataBar = "dataBar";
      internal const string DuplicateValues = "duplicateValues";
      internal const string EndsWith = "endsWith";
      internal const string Expression = "expression";
      internal const string IconSet = "iconSet";
      internal const string NotContainsBlanks = "notContainsBlanks";
      internal const string NotContainsErrors = "notContainsErrors";
      internal const string NotContainsText = "notContainsText";
      internal const string TimePeriod = "timePeriod";
      internal const string Top10 = "top10";
      internal const string UniqueValues = "UniqueValues";

      // EPPlus Extended Types
      internal const string AboveOrEqualAverage = "aboveOrEqualAverage";
      internal const string AboveStdDev = "aboveStdDev";
      internal const string BelowAverage = "belowAverage";
      internal const string BelowOrEqualAverage = "belowOrEqualAverage";
      internal const string BelowStdDev = "belowStdDev";
      internal const string Between = "between";
      internal const string Bottom = "bottom";
      internal const string BottomPercent = "bottomPercent";
      internal const string Equal = "equal";
      internal const string GreaterThan = "greaterThan";
      internal const string GreaterThanOrEqual = "greaterThanOrEqual";
      internal const string IconSet3 = "iconSet3";
      internal const string IconSet4 = "iconSet4";
      internal const string IconSet5 = "iconSet5";
      internal const string Last7Days = "last7Days";
      internal const string LastMonth = "lastMonth";
      internal const string LastWeek = "lastWeek";
      internal const string LessThan = "lessThan";
      internal const string LessThanOrEqual = "lessThanOrEqual";
      internal const string NextMonth = "nextMonth";
      internal const string NextWeek = "nextWeek";
      internal const string NotBetween = "notBetween";
      internal const string NotEqual = "notEqual";
      internal const string ThisMonth = "thisMonth";
      internal const string ThisWeek = "thisWeek";
      internal const string ThreeColorScale = "threeColorScale";
      internal const string Today = "today";
      internal const string Tomorrow = "tomorrow";
      internal const string Top = "top";
      internal const string TopPercent = "topPercent";
      internal const string TwoColorScale = "twoColorScale";
      internal const string Yesterday = "yesterday";
    }
    #endregion Rule Type ST_CfType §18.18.12 (with small EPPlus changes)

    #region CFVO Type ST_CfvoType §18.18.13
    internal class CfvoType
    {
      internal const string Min = "min";
      internal const string Max = "max";
      internal const string Num = "num";
      internal const string Formula = "formula";
      internal const string Percent = "percent";
      internal const string Percentile = "percentile";
    }
    #endregion CFVO Type ST_CfvoType §18.18.13

    #region Operator Type ST_ConditionalFormattingOperator §18.18.15
    internal class Operators
    {
      internal const string BeginsWith = "beginsWith";
      internal const string Between = "between";
      internal const string ContainsText = "containsText";
      internal const string EndsWith = "endsWith";
      internal const string Equal = "equal";
      internal const string GreaterThan = "greaterThan";
      internal const string GreaterThanOrEqual = "greaterThanOrEqual";
      internal const string LessThan = "lessThan";
      internal const string LessThanOrEqual = "lessThanOrEqual";
      internal const string NotBetween = "notBetween";
      internal const string NotContains = "notContains";
      internal const string NotEqual = "notEqual";
    }
    #endregion Operator Type ST_ConditionalFormattingOperator §18.18.15

    #region Time Period Type ST_TimePeriod §18.18.82
    internal class TimePeriods
    {
      internal const string Last7Days = "last7Days";
      internal const string LastMonth = "lastMonth";
      internal const string LastWeek = "lastWeek";
      internal const string NextMonth = "nextMonth";
      internal const string NextWeek = "nextWeek";
      internal const string ThisMonth = "thisMonth";
      internal const string ThisWeek = "thisWeek";
      internal const string Today = "today";
      internal const string Tomorrow = "tomorrow";
      internal const string Yesterday = "yesterday";
    }
    #endregion Time Period Type ST_TimePeriod §18.18.82

    #region Colors
    internal class Colors
    {
      internal const string CfvoLowValue = @"#FFF8696B";
      internal const string CfvoMiddleValue = @"#FFFFEB84";
      internal const string CfvoHighValue = @"#FF63BE7B";
    }
    #endregion Colors
  }
}