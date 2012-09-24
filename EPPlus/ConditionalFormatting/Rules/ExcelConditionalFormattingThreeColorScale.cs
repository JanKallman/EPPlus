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
using System.Drawing;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;

namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// ExcelConditionalFormattingThreeColorScale
  /// </summary>
  public class ExcelConditionalFormattingThreeColorScale
    : ExcelConditionalFormattingRule,
    IExcelConditionalFormattingThreeColorScale
  {
    /****************************************************************************************/

    #region Private Properties
    /// <summary>
    /// Private Low Value
    /// </summary>
    private ExcelConditionalFormattingColorScaleValue _lowValue;

    /// <summary>
    /// Private Middle Value
    /// </summary>
    private ExcelConditionalFormattingColorScaleValue _middleValue;

    /// <summary>
    /// Private High Value
    /// </summary>
    private ExcelConditionalFormattingColorScaleValue _highValue;
    #endregion Private Properties

    /****************************************************************************************/

    #region Constructors
    /// <summary>
    /// 
    /// </summary>
    /// <param name="type"></param>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    /// <param name="namespaceManager"></param>
    internal ExcelConditionalFormattingThreeColorScale(
      ExcelAddress address,
      int priority,
      ExcelWorksheet worksheet,
      XmlNode itemElementNode,
      XmlNamespaceManager namespaceManager)
      : base(
        eExcelConditionalFormattingRuleType.ThreeColorScale,
        address,
        priority,
        worksheet,
        itemElementNode,
        (namespaceManager == null) ? worksheet.NameSpaceManager : namespaceManager)
    {
      // Create the <colorScale> node inside the <cfRule> node
      var colorScaleNode = CreateComplexNode(
        Node,
        ExcelConditionalFormattingConstants.Paths.ColorScale);

      // LowValue default
      LowValue = new ExcelConditionalFormattingColorScaleValue(
        eExcelConditionalFormattingValueObjectPosition.Low,
        eExcelConditionalFormattingValueObjectType.Min,
        ColorTranslator.FromHtml(ExcelConditionalFormattingConstants.Colors.CfvoLowValue),
        eExcelConditionalFormattingRuleType.ThreeColorScale,
        address,
        priority,
        worksheet,
        NameSpaceManager);

      // MiddleValue default
      MiddleValue = new ExcelConditionalFormattingColorScaleValue(
        eExcelConditionalFormattingValueObjectPosition.Middle,
        eExcelConditionalFormattingValueObjectType.Percent,
        ColorTranslator.FromHtml(ExcelConditionalFormattingConstants.Colors.CfvoMiddleValue),
        50,
        string.Empty,
        eExcelConditionalFormattingRuleType.ThreeColorScale,
        address,
        priority,
        worksheet,
        NameSpaceManager);

      // HighValue default
      HighValue = new ExcelConditionalFormattingColorScaleValue(
        eExcelConditionalFormattingValueObjectPosition.High,
        eExcelConditionalFormattingValueObjectType.Max,
        ColorTranslator.FromHtml(ExcelConditionalFormattingConstants.Colors.CfvoHighValue),
        eExcelConditionalFormattingRuleType.ThreeColorScale,
        address,
        priority,
        worksheet,
        NameSpaceManager);
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode"></param>
    internal ExcelConditionalFormattingThreeColorScale(
      ExcelAddress address,
      int priority,
      ExcelWorksheet worksheet,
      XmlNode itemElementNode)
      : this(
        address,
        priority,
        worksheet,
        itemElementNode,
        null)
    {
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="priority"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    internal ExcelConditionalFormattingThreeColorScale(
      ExcelAddress address,
      int priority,
      ExcelWorksheet worksheet)
      : this(
        address,
        priority,
        worksheet,
        null,
        null)
    {
    }
    #endregion Constructors

    /****************************************************************************************/

    #region Public Properties
    /// <summary>
    /// Low Value for Three Color Scale Object Value
    /// </summary>
    public ExcelConditionalFormattingColorScaleValue LowValue
    {
      get { return _lowValue; }
      set { _lowValue = value; }
    }

    /// <summary>
    /// Middle Value for Three Color Scale Object Value
    /// </summary>
    public ExcelConditionalFormattingColorScaleValue MiddleValue
    {
      get { return _middleValue; }
      set { _middleValue = value; }
    }

    /// <summary>
    /// High Value for Three Color Scale Object Value
    /// </summary>
    public ExcelConditionalFormattingColorScaleValue HighValue
    {
      get { return _highValue; }
      set { _highValue = value; }
    }
    #endregion Public Properties

    /****************************************************************************************/
  }
}