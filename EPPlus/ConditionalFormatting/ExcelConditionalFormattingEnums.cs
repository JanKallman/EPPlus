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
  /// Enum for Conditional Format Type ST_CfType §18.18.12. With some changes.
  /// </summary>
  public enum eExcelConditionalFormattingRuleType
  {
    #region Average
    /// <summary>
    /// This conditional formatting rule highlights cells that are above the average
    /// for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    AboveAverage,

    /// <summary>
    /// This conditional formatting rule highlights cells that are above or equal
    /// the average for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    AboveOrEqualAverage,

    /// <summary>
    /// This conditional formatting rule highlights cells that are below the average
    /// for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    BelowAverage,

    /// <summary>
    /// This conditional formatting rule highlights cells that are below or equal
    /// the average for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    BelowOrEqualAverage,
    #endregion

    #region StdDev
    /// <summary>
    /// This conditional formatting rule highlights cells that are above the standard
    /// deviationa for all values in the range.
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    /// </summary>
    AboveStdDev,

    /// <summary>
    /// This conditional formatting rule highlights cells that are below the standard
    /// deviationa for all values in the range.
    /// </summary>
    /// <remarks>AboveAverage Excel CF Rule Type</remarks>
    BelowStdDev,
    #endregion

    #region TopBottom
    /// <summary>
    /// This conditional formatting rule highlights cells whose values fall in the
    /// bottom N bracket as specified.
    /// </summary>
    /// <remarks>Top10 Excel CF Rule Type</remarks>
    Bottom,

    /// <summary>
    /// This conditional formatting rule highlights cells whose values fall in the
    /// bottom N percent as specified.
    /// </summary>
    /// <remarks>Top10 Excel CF Rule Type</remarks>
    BottomPercent,

    /// <summary>
    /// This conditional formatting rule highlights cells whose values fall in the
    /// top N bracket as specified.
    /// </summary>
    /// <remarks>Top10 Excel CF Rule Type</remarks>
    Top,

    /// <summary>
    /// This conditional formatting rule highlights cells whose values fall in the
    /// top N percent as specified.
    /// </summary>
    /// <remarks>Top10 Excel CF Rule Type</remarks>
    TopPercent,
    #endregion

    #region TimePeriod
    /// <summary>
    /// This conditional formatting rule highlights cells containing dates in the
    /// last 7 days.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    Last7Days,

    /// <summary>
    /// This conditional formatting rule highlights cells containing dates in the
    /// last month.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    LastMonth,

    /// <summary>
    /// This conditional formatting rule highlights cells containing dates in the
    /// last week.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    LastWeek,

    /// <summary>
    /// This conditional formatting rule highlights cells containing dates in the
    /// next month.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    NextMonth,

    /// <summary>
    /// This conditional formatting rule highlights cells containing dates in the
    /// next week.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    NextWeek,

    /// <summary>
    /// This conditional formatting rule highlights cells containing dates in this
    /// month.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    ThisMonth,

    /// <summary>
    /// This conditional formatting rule highlights cells containing dates in this
    /// week.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    ThisWeek,

    /// <summary>
    /// This conditional formatting rule highlights cells containing today dates.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    Today,

    /// <summary>
    /// This conditional formatting rule highlights cells containing tomorrow dates.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    Tomorrow,

    /// <summary>
    /// This conditional formatting rule highlights cells containing yesterday dates.
    /// </summary>
    /// <remarks>TimePeriod Excel CF Rule Type</remarks>
    Yesterday,
    #endregion

    #region CellIs
    /// <summary>
    /// This conditional formatting rule highlights cells in the range that begin with
    /// the given text.
    /// </summary>
    /// <remarks>
    /// Equivalent to using the LEFT() sheet function and comparing values.
    /// </remarks>
    /// <remarks>BeginsWith Excel CF Rule Type</remarks>
    BeginsWith,

    /// <summary>
    /// This conditional formatting rule highlights cells in the range between the
    /// given two formulas.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    Between,

    /// <summary>
    /// This conditional formatting rule highlights cells that are completely blank.
    /// </summary>
    /// <remarks>
    /// Equivalent of using LEN(TRIM()). This means that if the cell contains only
    /// characters that TRIM() would remove, then it is considered blank. An empty cell
    /// is also considered blank.
    /// </remarks>
    /// <remarks>ContainsBlanks Excel CF Rule Type</remarks>
    ContainsBlanks,

    /// <summary>
    /// This conditional formatting rule highlights cells with formula errors.
    /// </summary>
    /// <remarks>
    /// Equivalent to using ISERROR() sheet function to determine if there is
    /// a formula error.
    /// </remarks>
    /// <remarks>ContainsErrors Excel CF Rule Type</remarks>
    ContainsErrors,

    /// <summary>
    /// This conditional formatting rule highlights cells in the range that begin with
    /// the given text.
    /// </summary>
    /// <remarks>
    /// Equivalent to using the LEFT() sheet function and comparing values.
    /// </remarks>
    /// <remarks>ContainsText Excel CF Rule Type</remarks>
    ContainsText,

    /// <summary>
    /// This conditional formatting rule highlights duplicated values.
    /// </summary>
    /// <remarks>DuplicateValues Excel CF Rule Type</remarks>
    DuplicateValues,

    /// <summary>
    /// This conditional formatting rule highlights cells ending with given text.
    /// </summary>
    /// <remarks>
    /// Equivalent to using the RIGHT() sheet function and comparing values.
    /// </remarks>
    /// <remarks>EndsWith Excel CF Rule Type</remarks>
    EndsWith,

    /// <summary>
    /// This conditional formatting rule highlights cells equals to with given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    Equal,

    /// <summary>
    /// This conditional formatting rule contains a formula to evaluate. When the 
    /// formula result is true, the cell is highlighted.
    /// </summary>
    /// <remarks>Expression Excel CF Rule Type</remarks>
    Expression,

    /// <summary>
    /// This conditional formatting rule highlights cells greater than the given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    GreaterThan,

    /// <summary>
    /// This conditional formatting rule highlights cells greater than or equal the
    /// given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    GreaterThanOrEqual,

    /// <summary>
    /// This conditional formatting rule highlights cells less than the given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    LessThan,

    /// <summary>
    /// This conditional formatting rule highlights cells less than or equal the
    /// given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    LessThanOrEqual,

    /// <summary>
    /// This conditional formatting rule highlights cells outside the range in
    /// given two formulas.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    NotBetween,

    /// <summary>
    /// This conditional formatting rule highlights cells that does not contains the
    /// given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    NotContains,

    /// <summary>
    /// This conditional formatting rule highlights cells that are not blank.
    /// </summary>
    /// <remarks>
    /// Equivalent of using LEN(TRIM()). This means that if the cell contains only
    /// characters that TRIM() would remove, then it is considered blank. An empty cell
    /// is also considered blank.
    /// </remarks>
    /// <remarks>NotContainsBlanks Excel CF Rule Type</remarks>
    NotContainsBlanks,

    /// <summary>
    /// This conditional formatting rule highlights cells without formula errors.
    /// </summary>
    /// <remarks>
    /// Equivalent to using ISERROR() sheet function to determine if there is a
    /// formula error.
    /// </remarks>
    /// <remarks>NotContainsErrors Excel CF Rule Type</remarks>
    NotContainsErrors,

    /// <summary>
    /// This conditional formatting rule highlights cells that do not contain
    /// the given text.
    /// </summary>
    /// <remarks>
    /// Equivalent to using the SEARCH() sheet function.
    /// </remarks>
    /// <remarks>NotContainsText Excel CF Rule Type</remarks>
    NotContainsText,

    /// <summary>
    /// This conditional formatting rule highlights cells not equals to with
    /// given formula.
    /// </summary>
    /// <remarks>CellIs Excel CF Rule Type</remarks>
    NotEqual,

    /// <summary>
    /// This conditional formatting rule highlights unique values in the range.
    /// </summary>
    /// <remarks>UniqueValues Excel CF Rule Type</remarks>
    UniqueValues,
    #endregion

    #region ColorScale
    /// <summary>
    /// Three Color Scale (Low, Middle and High Color Scale)
    /// </summary>
    /// <remarks>ColorScale Excel CF Rule Type</remarks>
    ThreeColorScale,

    /// <summary>
    /// Two Color Scale (Low and High Color Scale)
    /// </summary>
    /// <remarks>ColorScale Excel CF Rule Type</remarks>
    TwoColorScale,
    #endregion

    #region IconSet
    /// <summary>
    /// This conditional formatting rule applies a 3 set icons to cells according
    /// to their values.
    /// </summary>
    /// <remarks>IconSet Excel CF Rule Type</remarks>
    ThreeIconSet,

    /// <summary>
    /// This conditional formatting rule applies a 4 set icons to cells according
    /// to their values.
    /// </summary>
    /// <remarks>IconSet Excel CF Rule Type</remarks>
    FourIconSet,

    /// <summary>
    /// This conditional formatting rule applies a 5 set icons to cells according
    /// to their values.
    /// </summary>
    /// <remarks>IconSet Excel CF Rule Type</remarks>
    FiveIconSet,
    #endregion

    #region DataBar
    /// <summary>
    /// This conditional formatting rule displays a gradated data bar in the range of cells.
    /// </summary>
    /// <remarks>DataBar Excel CF Rule Type</remarks>
    DataBar
    #endregion
  }

  /// <summary>
  /// Enum for Conditional Format Value Object Type ST_CfvoType §18.18.13
  /// </summary>
  public enum eExcelConditionalFormattingValueObjectType
  {
    /// <summary>
    /// Formula
    /// </summary>
    Formula,

    /// <summary>
    /// Maximum Value
    /// </summary>
    Max,

    /// <summary>
    /// Minimum Value
    /// </summary>
    Min,

    /// <summary>
    /// Number Value
    /// </summary>
    Num,

    /// <summary>
    /// Percent
    /// </summary>
    Percent,

    /// <summary>
    /// Percentile
    /// </summary>
    Percentile
  }

  /// <summary>
  /// Enum for Conditional Formatting Value Object Position
  /// </summary>
  public enum eExcelConditionalFormattingValueObjectPosition
  {
    /// <summary>
    /// The lower position for both TwoColorScale and ThreeColorScale
    /// </summary>
    Low,

    /// <summary>
    /// The middle position only for ThreeColorScale
    /// </summary>
    Middle,

    /// <summary>
    /// The highest position for both TwoColorScale and ThreeColorScale
    /// </summary>
    High
  }

  /// <summary>
  /// Enum for Conditional Formatting Value Object Node Type
  /// </summary>
  public enum eExcelConditionalFormattingValueObjectNodeType
  {
    /// <summary>
    /// 'cfvo' node
    /// </summary>
    Cfvo,

    /// <summary>
    /// 'color' node
    /// </summary>
    Color
  }

  /// <summary>
  /// Enum for Conditional Formatting Operartor Type ST_ConditionalFormattingOperator §18.18.15
  /// </summary>
  public enum eExcelConditionalFormattingOperatorType
  {
    /// <summary>
    /// Begins With. 'Begins with' operator
    /// </summary>
    BeginsWith,

    /// <summary>
    /// Between. 'Between' operator
    /// </summary>
    Between,

    /// <summary>
    /// Contains. 'Contains' operator
    /// </summary>
    ContainsText,

    /// <summary>
    /// Ends With. 'Ends with' operator
    /// </summary>
    EndsWith,

    /// <summary>
    /// Equal. 'Equal to' operator
    /// </summary>
    Equal,

    /// <summary>
    /// Greater Than. 'Greater than' operator
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Greater Than Or Equal. 'Greater than or equal to' operator
    /// </summary>
    GreaterThanOrEqual,

    /// <summary>
    /// Less Than. 'Less than' operator
    /// </summary>
    LessThan,

    /// <summary>
    /// Less Than Or Equal. 'Less than or equal to' operator
    /// </summary>
    LessThanOrEqual,

    /// <summary>
    /// Not Between. 'Not between' operator
    /// </summary>
    NotBetween,

    /// <summary>
    /// Does Not Contain. 'Does not contain' operator
    /// </summary>
    NotContains,

    /// <summary>
    /// Not Equal. 'Not equal to' operator
    /// </summary>
    NotEqual
  }

  /// <summary>
  /// Enum for Conditional Formatting Time Period Type ST_TimePeriod §18.18.82
  /// </summary>
  public enum eExcelConditionalFormattingTimePeriodType
  {
    /// <summary>
    /// Last 7 Days. A date in the last seven days.
    /// </summary>
    Last7Days,

    /// <summary>
    /// Last Month. A date occuring in the last calendar month.
    /// </summary>
    LastMonth,

    /// <summary>
    /// Last Week. A date occuring last week.
    /// </summary>
    LastWeek,

    /// <summary>
    /// Next Month. A date occuring in the next calendar month.
    /// </summary>
    NextMonth,

    /// <summary>
    /// Next Week. A date occuring next week.
    /// </summary>
    NextWeek,

    /// <summary>
    /// This Month. A date occuring in this calendar month.
    /// </summary>
    ThisMonth,

    /// <summary>
    /// This Week. A date occuring this week.
    /// </summary>
    ThisWeek,

    /// <summary>
    /// Today. Today's date.
    /// </summary>
    Today,

    /// <summary>
    /// Tomorrow. Tomorrow's date.
    /// </summary>
    Tomorrow,

    /// <summary>
    /// Yesterday. Yesterday's date.
    /// </summary>
    Yesterday
  }

  /// <summary>
  /// 18.18.42 ST_IconSetType (Icon Set Type) - Only 3 icons
  /// </summary>
  public enum eExcelconditionalFormatting3IconsSetType
  {
    /// <summary>
    /// (3 Arrows) 3 arrows icon set.
    /// </summary>
    Arrows,

    /// <summary>
    /// (3 Arrows (Gray)) 3 gray arrows icon set.
    /// </summary>
    ArrowsGray,

    /// <summary>
    /// (3 Flags) 3 flags icon set. 
    /// </summary>
    Flags,

    /// <summary>
    /// (3 Signs) 3 signs icon set.
    /// </summary>
    Signs,

    /// <summary>
    /// (3 Symbols Circled) 3 symbols icon set.
    /// </summary>
    Symbols,

    /// <summary>
    /// (3 Symbols) 3 Symbols icon set.
    /// </summary>
    Symbols2,

    /// <summary>
    /// (3 Traffic Lights) 3 traffic lights icon set (#1).
    /// </summary>
    TrafficLights1,

    /// <summary>
    /// (3 Traffic Lights Black) 3 traffic lights icon set with thick black border.
    /// </summary>
    TrafficLights2
  }

  /// <summary>
  /// 18.18.42 ST_IconSetType (Icon Set Type) - Only 4 icons
  /// </summary>
  public enum eExcelconditionalFormatting4IconsSetType
  {
    /// <summary>
    /// (4 Arrows) 4 arrows icon set.
    /// </summary>
    Arrows,

    /// <summary>
    /// (4 Arrows (Gray)) 4 gray arrows icon set.
    /// </summary>
    ArrowsGray,

    /// <summary>
    /// (4 Ratings) 4 ratings icon set.
    /// </summary>
    Rating,

    /// <summary>
    /// (4 Red To Black) 4 'red to black' icon set.
    /// </summary>
    RedToBlack,

    /// <summary>
    /// (4 Traffic Lights) 4 traffic lights icon set.
    /// </summary>
    TrafficLights
  }

  /// <summary>
  /// 18.18.42 ST_IconSetType (Icon Set Type) - Only 5 icons
  /// </summary>
  public enum eExcelconditionalFormatting5IconsSetType
  {
    /// <summary>
    /// (5 Arrows) 5 arrows icon set.
    /// </summary>
    Arrows,

    /// <summary>
    /// (5 Arrows (Gray)) 5 gray arrows icon set.
    /// </summary>
    ArrowsGray,

    /// <summary>
    /// (5 Quarters) 5 quarters icon set.
    /// </summary>
    Quarters,

    /// <summary>
    /// (5 Ratings Icon Set) 5 rating icon set.
    /// </summary>
    Rating
  }
  /// <summary>
  /// 18.18.42 ST_IconSetType (Icon Set Type)
  /// </summary>
  public enum eExcelconditionalFormattingIconsSetType
  {
      /// <summary>
      /// (3 Arrows) 3 arrows icon set.
      /// </summary>
      ThreeArrows,

      /// <summary>
      /// (3 Arrows (Gray)) 3 gray arrows icon set.
      /// </summary>
      ThreeArrowsGray,

      /// <summary>
      /// (3 Flags) 3 flags icon set. 
      /// </summary>
      ThreeFlags,

      /// <summary>
      /// (3 Signs) 3 signs icon set.
      /// </summary>
      ThreeSigns,

      /// <summary>
      /// (3 Symbols Circled) 3 symbols icon set.
      /// </summary>
      ThreeSymbols,

      /// <summary>
      /// (3 Symbols) 3 Symbols icon set.
      /// </summary>
      ThreeSymbols2,

      /// <summary>
      /// (3 Traffic Lights) 3 traffic lights icon set (#1).
      /// </summary>
      ThreeTrafficLights1,

      /// <summary>
      /// (3 Traffic Lights Black) 3 traffic lights icon set with thick black border.
      /// </summary>
      ThreeTrafficLights2,
 
    /// <summary>
    /// (4 Arrows) 4 arrows icon set.
    /// </summary>
    FourArrows,

    /// <summary>
    /// (4 Arrows (Gray)) 4 gray arrows icon set.
    /// </summary>
    FourArrowsGray,

    /// <summary>
    /// (4 Ratings) 4 ratings icon set.
    /// </summary>
    FourRating,

    /// <summary>
    /// (4 Red To Black) 4 'red to black' icon set.
    /// </summary>
    FourRedToBlack,

    /// <summary>
    /// (4 Traffic Lights) 4 traffic lights icon set.
    /// </summary>
    FourTrafficLights,

      /// <summary>
    /// (5 Arrows) 5 arrows icon set.
    /// </summary>
    FiveArrows,

    /// <summary>
    /// (5 Arrows (Gray)) 5 gray arrows icon set.
    /// </summary>
    FiveArrowsGray,

    /// <summary>
    /// (5 Quarters) 5 quarters icon set.
    /// </summary>
    FiveQuarters,

    /// <summary>
    /// (5 Ratings Icon Set) 5 rating icon set.
    /// </summary>
    FiveRating
}
}