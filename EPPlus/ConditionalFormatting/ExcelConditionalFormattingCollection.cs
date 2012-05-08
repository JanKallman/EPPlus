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
 * Author					Change						                Date
 * ******************************************************************************
 * Eyal Seagull		Conditional Formatting            2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using OfficeOpenXml.Utils;
using System.Xml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.ConditionalFormatting
{
  /// <summary>
  /// Collection of <see cref="ExcelConditionalFormattingRule"/>.
  /// This class is providing the API for EPPlus conditional formatting.
  /// </summary>
  /// <remarks>
  /// <para>
  /// The public methods of this class (Add[...]ConditionalFormatting) will create a ConditionalFormatting/CfRule entry in the worksheet. When this
  /// Conditional Formatting has been created changes to the properties will affect the workbook immediately.
  /// </para>
  /// <para>
  /// Each type of Conditional Formatting Rule has diferente set of properties.
  /// </para>
  /// <code>
  /// // Add a Three Color Scale conditional formatting
  /// var cf = worksheet.ConditionalFormatting.AddThreeColorScale(new ExcelAddress("A1:C10"));
  /// // Set the conditional formatting properties
  /// cf.LowValue.Type = ExcelConditionalFormattingValueObjectType.Min;
  /// cf.LowValue.Color = Color.White;
  /// cf.MiddleValue.Type = ExcelConditionalFormattingValueObjectType.Percent;
  /// cf.MiddleValue.Value = 50;
  /// cf.MiddleValue.Color = Color.Blue;
  /// cf.HighValue.Type = ExcelConditionalFormattingValueObjectType.Max;
  /// cf.HighValue.Color = Color.Black;
  /// </code>
  /// </remarks>
  public class ExcelConditionalFormattingCollection
    : XmlHelper,
    IEnumerable<IExcelConditionalFormattingRule>
  {
    /****************************************************************************************/

    #region Private Properties
    private List<IExcelConditionalFormattingRule> _rules = new List<IExcelConditionalFormattingRule>();
    private ExcelWorksheet _worksheet = null;
    #endregion Private Properties

    /****************************************************************************************/

    #region Constructors
    /// <summary>
    /// Initialize the <see cref="ExcelConditionalFormattingCollection"/>
    /// </summary>
    /// <param name="worksheet"></param>
    internal ExcelConditionalFormattingCollection(
      ExcelWorksheet worksheet)
      : base(
        worksheet.NameSpaceManager,
        worksheet.WorksheetXml.DocumentElement)
    {
      Require.Argument(worksheet).IsNotNull("worksheet");

      _worksheet = worksheet;
      SchemaNodeOrder = _worksheet.SchemaNodeOrder;

      // Look for all the <conditionalFormatting>
      var conditionalFormattingNodes = TopNode.SelectNodes(
        "//" + ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
        _worksheet.NameSpaceManager);

      // Check if we found at least 1 node
      if ((conditionalFormattingNodes != null)
        && (conditionalFormattingNodes.Count > 0))
      {
        // Foreach <conditionalFormatting>
        foreach (XmlNode conditionalFormattingNode in conditionalFormattingNodes)
        {
          // Check if @sqref attribute exists
          if (conditionalFormattingNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref] == null)
          {
            throw new Exception(
              ExcelConditionalFormattingConstants.Errors.MissingSqrefAttribute);
          }

          // Get the @sqref attribute
          ExcelAddress address = new ExcelAddress(
            conditionalFormattingNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Sqref].Value);

          // Check for all the <cfRules> nodes and load them
          var cfRuleNodes = conditionalFormattingNode.SelectNodes(
            ExcelConditionalFormattingConstants.Paths.CfRule,
            _worksheet.NameSpaceManager);

          // Foreach <cfRule> inside the current <conditionalFormatting>
          foreach (XmlNode cfRuleNode in cfRuleNodes)
          {
            // Check if @type attribute exists
            if (cfRuleNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Type] == null)
            {
              throw new Exception(
                ExcelConditionalFormattingConstants.Errors.MissingTypeAttribute);
            }

            // Check if @priority attribute exists
            if (cfRuleNode.Attributes[ExcelConditionalFormattingConstants.Attributes.Priority] == null)
            {
              throw new Exception(
                ExcelConditionalFormattingConstants.Errors.MissingPriorityAttribute);
            }

            // Get the <cfRule> main attributes
            string typeAttribute = ExcelConditionalFormattingHelper.GetAttributeString(
              cfRuleNode,
              ExcelConditionalFormattingConstants.Attributes.Type);

            int priority = ExcelConditionalFormattingHelper.GetAttributeInt(
              cfRuleNode,
              ExcelConditionalFormattingConstants.Attributes.Priority);

            // Transform the @type attribute to EPPlus Rule Type (slighty diferente)
            var type = ExcelConditionalFormattingRuleType.GetTypeByAttrbiute(
              typeAttribute,
              cfRuleNode,
              _worksheet.NameSpaceManager);

            // Create the Rule according to the correct type, address and priority
            var cfRule = ExcelConditionalFormattingRuleFactory.Create(
              type,
              address,
              priority,
              _worksheet,
              cfRuleNode);

            // Add the new rule to the list
            _rules.Add(cfRule);
          }
        }
      }
    }
    #endregion Constructors

    /****************************************************************************************/

    #region Methods
    /// <summary>
    /// 
    /// </summary>
    private void EnsureRootElementExists()
    {
      // Find the <worksheet> node
      if (_worksheet.WorksheetXml.DocumentElement == null)
      {
        throw new Exception(
          ExcelConditionalFormattingConstants.Errors.MissingWorksheetNode);
      }
    }

    /// <summary>
    /// GetRootNode
    /// </summary>
    /// <returns></returns>
    private XmlNode GetRootNode()
    {
      EnsureRootElementExists();
      return _worksheet.WorksheetXml.DocumentElement;
    }

    /// <summary>
    /// Validates address - not empty (collisions are allowded)
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    private ExcelAddress ValidateAddress(
      ExcelAddress address)
    {
      Require.Argument(address).IsNotNull("address");

      //TODO: Are there any other validation we need to do?
      return address;
    }

    /// <summary>
    /// Get the next priority sequencial number
    /// </summary>
    /// <returns></returns>
    private int GetNextPriority()
    {
      // Consider zero as the last priority when we have no CF rules
      int lastPriority = 0;

      // Search for the last priority
      foreach (var cfRule in _rules)
      {
        if (cfRule.Priority > lastPriority)
        {
          lastPriority = cfRule.Priority;
        }
      }

      // Our next priority is the last plus one
      return lastPriority + 1;
    }
    #endregion Methods

    /****************************************************************************************/

    #region IEnumerable<IExcelConditionalFormatting>
    /// <summary>
    /// Number of validations
    /// </summary>
    public int Count
    {
      get { return _rules.Count; }
    }

    /// <summary>
    /// Index operator, returns by 0-based index
    /// </summary>
    /// <param name="index"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingRule this[int index]
    {
      get { return _rules[index]; }
      set { _rules[index] = value; }
    }

    /// <summary>
    /// Get the 'cfRule' enumerator
    /// </summary>
    /// <returns></returns>
    IEnumerator<IExcelConditionalFormattingRule> IEnumerable<IExcelConditionalFormattingRule>.GetEnumerator()
    {
      return _rules.GetEnumerator();
    }

    /// <summary>
    /// Get the 'cfRule' enumerator
    /// </summary>
    /// <returns></returns>
    IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
      return _rules.GetEnumerator();
    }

    /// <summary>
    /// Removes all 'cfRule' from the collection and from the XML.
    /// <remarks>
    /// This is the same as removing all the 'conditionalFormatting' nodes.
    /// </remarks>
    /// </summary>
    public void RemoveAll()
    {
      // Look for all the <conditionalFormatting> nodes
      var conditionalFormattingNodes = TopNode.SelectNodes(
        "//" + ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
        _worksheet.NameSpaceManager);

      // Remove all the <conditionalFormatting> nodes one by one
      foreach (XmlNode conditionalFormattingNode in conditionalFormattingNodes)
      {
        conditionalFormattingNode.ParentNode.RemoveChild(conditionalFormattingNode);
      }

      // Clear the <cfRule> item list
      _rules.Clear();
    }

    /// <summary>
    /// Remove a Conditional Formatting Rule by its object
    /// </summary>
    /// <param name="item"></param>
    public void Remove(
      IExcelConditionalFormattingRule item)
    {
      Require.Argument(item).IsNotNull("item");

      try
      {
        // Point to the parent node
        var oldParentNode = item.Node.ParentNode;

        // Remove the <cfRule> from the old <conditionalFormatting> parent node
        oldParentNode.RemoveChild(item.Node);

        // Check if the old <conditionalFormatting> parent node has <cfRule> node inside it
        if (!oldParentNode.HasChildNodes)
        {
          // Remove the old parent node
          oldParentNode.ParentNode.RemoveChild(oldParentNode);
        }

        _rules.Remove(item);
      }
      catch
      {
        throw new Exception(
          ExcelConditionalFormattingConstants.Errors.InvalidRemoveRuleOperation);
      }
    }

    /// <summary>
    /// Remove a Conditional Formatting Rule by its 0-based index
    /// </summary>
    /// <param name="index"></param>
    public void RemoveAt(
      int index)
    {
      Require.Argument(index).IsInRange(0, this.Count - 1, "index");

      Remove(this[index]);
    }

    /// <summary>
    /// Remove a Conditional Formatting Rule by its priority
    /// </summary>
    /// <param name="priority"></param>
    public void RemoveByPriority(
      int priority)
    {
      try
      {
        Remove(RulesByPriority(priority));
      }
      catch
      {
      }
    }

    /// <summary>
    /// Get a rule by its priority
    /// </summary>
    /// <param name="priority"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingRule RulesByPriority(
      int priority)
    {
      return _rules.Find(x => x.Priority == priority);
    }
    #endregion IEnumerable<IExcelConditionalFormatting>

    /****************************************************************************************/

    #region Conditional Formatting Rules
    /// <summary>
    /// Add rule (internal)
    /// </summary>
    /// <param name="type"></param>
    /// <param name="address"></param>
    /// <returns></returns>
    internal IExcelConditionalFormattingRule AddRule(
      eExcelConditionalFormattingRuleType type,
      ExcelAddress address)
    {
      Require.Argument(address).IsNotNull("address");

      address = ValidateAddress(address);
      EnsureRootElementExists();

      // Create the Rule according to the correct type, address and priority
      IExcelConditionalFormattingRule cfRule = ExcelConditionalFormattingRuleFactory.Create(
        type,
        address,
        GetNextPriority(),
        _worksheet,
        null);

      // Add the newly created rule to the list
      _rules.Add(cfRule);

      // Return the newly created rule
      return cfRule;
    }

    /// <summary>
    /// Add AboveAverage Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingAverageGroup AddAboveAverage(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingAverageGroup)AddRule(
        eExcelConditionalFormattingRuleType.AboveAverage,
        address);
    }

    /// <summary>
    /// Add AboveOrEqualAverage Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingAverageGroup AddAboveOrEqualAverage(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingAverageGroup)AddRule(
        eExcelConditionalFormattingRuleType.AboveOrEqualAverage,
        address);
    }

    /// <summary>
    /// Add BelowAverage Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingAverageGroup AddBelowAverage(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingAverageGroup)AddRule(
        eExcelConditionalFormattingRuleType.BelowAverage,
        address);
    }

    /// <summary>
    /// Add BelowOrEqualAverage Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingAverageGroup AddBelowOrEqualAverage(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingAverageGroup)AddRule(
        eExcelConditionalFormattingRuleType.BelowOrEqualAverage,
        address);
    }

    /// <summary>
    /// Add AboveStdDev Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingStdDevGroup AddAboveStdDev(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingStdDevGroup)AddRule(
        eExcelConditionalFormattingRuleType.AboveStdDev,
        address);
    }

    /// <summary>
    /// Add BelowStdDev Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingStdDevGroup AddBelowStdDev(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingStdDevGroup)AddRule(
        eExcelConditionalFormattingRuleType.BelowStdDev,
        address);
    }

    /// <summary>
    /// Add Bottom Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTopBottomGroup AddBottom(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTopBottomGroup)AddRule(
        eExcelConditionalFormattingRuleType.Bottom,
        address);
    }

    /// <summary>
    /// Add BottomPercent Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTopBottomGroup AddBottomPercent(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTopBottomGroup)AddRule(
        eExcelConditionalFormattingRuleType.BottomPercent,
        address);
    }

    /// <summary>
    /// Add Top Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTopBottomGroup AddTop(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTopBottomGroup)AddRule(
        eExcelConditionalFormattingRuleType.Top,
        address);
    }

    /// <summary>
    /// Add TopPercent Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTopBottomGroup AddTopPercent(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTopBottomGroup)AddRule(
        eExcelConditionalFormattingRuleType.TopPercent,
        address);
    }

    /// <summary>
    /// Add Last7Days Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddLast7Days(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.Last7Days,
        address);
    }

    /// <summary>
    /// Add LastMonth Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddLastMonth(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.LastMonth,
        address);
    }

    /// <summary>
    /// Add LastWeek Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddLastWeek(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.LastWeek,
        address);
    }

    /// <summary>
    /// Add NextMonth Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddNextMonth(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.NextMonth,
        address);
    }

    /// <summary>
    /// Add NextWeek Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddNextWeek(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.NextWeek,
        address);
    }

    /// <summary>
    /// Add ThisMonth Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddThisMonth(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.ThisMonth,
        address);
    }

    /// <summary>
    /// Add ThisWeek Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddThisWeek(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.ThisWeek,
        address);
    }

    /// <summary>
    /// Add Today Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddToday(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.Today,
        address);
    }

    /// <summary>
    /// Add Tomorrow Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddTomorrow(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.Tomorrow,
        address);
    }

    /// <summary>
    /// Add Yesterday Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTimePeriodGroup AddYesterday(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTimePeriodGroup)AddRule(
        eExcelConditionalFormattingRuleType.Yesterday,
        address);
    }

    /// <summary>
    /// Add BeginsWith Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingBeginsWith AddBeginsWith(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingBeginsWith)AddRule(
        eExcelConditionalFormattingRuleType.BeginsWith,
        address);
    }

    /// <summary>
    /// Add Between Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingBetween AddBetween(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingBetween)AddRule(
        eExcelConditionalFormattingRuleType.Between,
        address);
    }

    /// <summary>
    /// Add ContainsBlanks Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingContainsBlanks AddContainsBlanks(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingContainsBlanks)AddRule(
        eExcelConditionalFormattingRuleType.ContainsBlanks,
        address);
    }

    /// <summary>
    /// Add ContainsErrors Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingContainsErrors AddContainsErrors(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingContainsErrors)AddRule(
        eExcelConditionalFormattingRuleType.ContainsErrors,
        address);
    }

    /// <summary>
    /// Add ContainsText Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingContainsText AddContainsText(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingContainsText)AddRule(
        eExcelConditionalFormattingRuleType.ContainsText,
        address);
    }

    /// <summary>
    /// Add DuplicateValues Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingDuplicateValues AddDuplicateValues(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingDuplicateValues)AddRule(
        eExcelConditionalFormattingRuleType.DuplicateValues,
        address);
    }

    /// <summary>
    /// Add EndsWith Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingEndsWith AddEndsWith(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingEndsWith)AddRule(
        eExcelConditionalFormattingRuleType.EndsWith,
        address);
    }

    /// <summary>
    /// Add Equal Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingEqual AddEqual(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingEqual)AddRule(
        eExcelConditionalFormattingRuleType.Equal,
        address);
    }

    /// <summary>
    /// Add Expression Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingExpression AddExpression(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingExpression)AddRule(
        eExcelConditionalFormattingRuleType.Expression,
        address);
    }

    /// <summary>
    /// Add GreaterThan Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingGreaterThan AddGreaterThan(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingGreaterThan)AddRule(
        eExcelConditionalFormattingRuleType.GreaterThan,
        address);
    }

    /// <summary>
    /// Add GreaterThanOrEqual Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingGreaterThanOrEqual AddGreaterThanOrEqual(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingGreaterThanOrEqual)AddRule(
        eExcelConditionalFormattingRuleType.GreaterThanOrEqual,
        address);
    }

    /// <summary>
    /// Add LessThan Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingLessThan AddLessThan(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingLessThan)AddRule(
        eExcelConditionalFormattingRuleType.LessThan,
        address);
    }

    /// <summary>
    /// Add LessThanOrEqual Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingLessThanOrEqual AddLessThanOrEqual(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingLessThanOrEqual)AddRule(
        eExcelConditionalFormattingRuleType.LessThanOrEqual,
        address);
    }

    /// <summary>
    /// Add NotBetween Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingNotBetween AddNotBetween(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingNotBetween)AddRule(
        eExcelConditionalFormattingRuleType.NotBetween,
        address);
    }

    /// <summary>
    /// Add NotContainsBlanks Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingNotContainsBlanks AddNotContainsBlanks(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingNotContainsBlanks)AddRule(
        eExcelConditionalFormattingRuleType.NotContainsBlanks,
        address);
    }

    /// <summary>
    /// Add NotContainsErrors Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingNotContainsErrors AddNotContainsErrors(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingNotContainsErrors)AddRule(
        eExcelConditionalFormattingRuleType.NotContainsErrors,
        address);
    }

    /// <summary>
    /// Add NotContainsText Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingNotContainsText AddNotContainsText(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingNotContainsText)AddRule(
        eExcelConditionalFormattingRuleType.NotContainsText,
        address);
    }

    /// <summary>
    /// Add NotEqual Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingNotEqual AddNotEqual(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingNotEqual)AddRule(
        eExcelConditionalFormattingRuleType.NotEqual,
        address);
    }

    /// <summary>
    /// Add Unique Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingUniqueValues AddUniqueValues(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingUniqueValues)AddRule(
        eExcelConditionalFormattingRuleType.UniqueValues,
        address);
    }

    /// <summary>
    /// Add ThreeColorScale Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingThreeColorScale AddThreeColorScale(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingThreeColorScale)AddRule(
        eExcelConditionalFormattingRuleType.ThreeColorScale,
        address);
    }

    /// <summary>
    /// Add TwoColorScale Rule
    /// </summary>
    /// <param name="address"></param>
    /// <returns></returns>
    public IExcelConditionalFormattingTwoColorScale AddTwoColorScale(
      ExcelAddress address)
    {
      return (IExcelConditionalFormattingTwoColorScale)AddRule(
        eExcelConditionalFormattingRuleType.TwoColorScale,
        address);
    }

    //TODO: Add the DataBar and IconSet
    #endregion Conditional Formatting Rules

    /****************************************************************************************/
  }
}