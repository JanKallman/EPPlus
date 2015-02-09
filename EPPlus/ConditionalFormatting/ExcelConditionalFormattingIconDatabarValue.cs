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
using OfficeOpenXml.Utils;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Security;

namespace OfficeOpenXml.ConditionalFormatting
{
	/// <summary>
	/// 18.3.1.11 cfvo (Conditional Format Value Object)
	/// Describes the values of the interpolation points in a gradient scale.
	/// </summary>
	public class ExcelConditionalFormattingIconDataBarValue
		: XmlHelper
	{
		/****************************************************************************************/

		#region Private Properties
		private eExcelConditionalFormattingRuleType _ruleType;
		private ExcelWorksheet _worksheet;
		#endregion Private Properties

		/****************************************************************************************/

		#region Constructors
    /// <summary>
    /// Initialize the cfvo (§18.3.1.11) node
    /// </summary>
    /// <param name="type"></param>
    /// <param name="value"></param>
    /// <param name="formula"></param>
    /// <param name="ruleType"></param>
    /// <param name="address"></param>
    /// <param name="priority"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode">The cfvo (§18.3.1.11) node parent. Can be any of the following:
    /// colorScale (§18.3.1.16); dataBar (§18.3.1.28); iconSet (§18.3.1.49)</param>
    /// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingIconDataBarValue(
			eExcelConditionalFormattingValueObjectType type,
			double value,
			string formula,
			eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            int priority,
			ExcelWorksheet worksheet,
			XmlNode itemElementNode,
			XmlNamespaceManager namespaceManager)
			: this(
            ruleType,
            address,
            worksheet,
            itemElementNode,
			namespaceManager)
		{
			Require.Argument(priority).IsInRange(1, int.MaxValue, "priority");

            // Check if the parent does not exists
			if (itemElementNode == null)
			{
				// Get the parent node path by the rule type
				string parentNodePath = ExcelConditionalFormattingValueObjectType.GetParentPathByRuleType(
					ruleType);

				// Check for en error (rule type does not have <cfvo>)
				if (parentNodePath == string.Empty)
				{
					throw new Exception(
						ExcelConditionalFormattingConstants.Errors.MissingCfvoParentNode);
				}

				// Point to the <cfvo> parent node
        itemElementNode = _worksheet.WorksheetXml.SelectSingleNode(
					string.Format(
						"//{0}[{1}='{2}']/{3}[{4}='{5}']/{6}",
					// {0}
						ExcelConditionalFormattingConstants.Paths.ConditionalFormatting,
					// {1}
						ExcelConditionalFormattingConstants.Paths.SqrefAttribute,
					// {2}
						address.Address,
					// {3}
						ExcelConditionalFormattingConstants.Paths.CfRule,
					// {4}
						ExcelConditionalFormattingConstants.Paths.PriorityAttribute,
					// {5}
						priority,
					// {6}
						parentNodePath),
					_worksheet.NameSpaceManager);

				// Check for en error (rule type does not have <cfvo>)
                if (itemElementNode == null)
				{
					throw new Exception(
						ExcelConditionalFormattingConstants.Errors.MissingCfvoParentNode);
				}
			}

            TopNode = itemElementNode;

			// Save the attributes
			RuleType = ruleType;
			Type = type;
			Value = value;
			Formula = formula;
		}
    /// <summary>
    /// Initialize the cfvo (§18.3.1.11) node
    /// </summary>
    /// <param name="ruleType"></param>
    /// <param name="address"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode">The cfvo (§18.3.1.11) node parent. Can be any of the following:
    /// colorScale (§18.3.1.16); dataBar (§18.3.1.28); iconSet (§18.3.1.49)</param>
    /// <param name="namespaceManager"></param>
        internal ExcelConditionalFormattingIconDataBarValue(
            eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            ExcelWorksheet worksheet,
            XmlNode itemElementNode,
            XmlNamespaceManager namespaceManager)
            : base(
                namespaceManager,
                itemElementNode)
        {
            Require.Argument(address).IsNotNull("address");
            Require.Argument(worksheet).IsNotNull("worksheet");

            // Save the worksheet for private methods to use
            _worksheet = worksheet;

            // Schema order list
            SchemaNodeOrder = new string[]
			{
                ExcelConditionalFormattingConstants.Nodes.Cfvo,
			};

            //Check if the parent does not exists
            if (itemElementNode == null)
            {
                // Get the parent node path by the rule type
                string parentNodePath = ExcelConditionalFormattingValueObjectType.GetParentPathByRuleType(
                    ruleType);

                // Check for en error (rule type does not have <cfvo>)
                if (parentNodePath == string.Empty)
                {
                    throw new Exception(
                        ExcelConditionalFormattingConstants.Errors.MissingCfvoParentNode);
                }
            }
            RuleType = ruleType;            
        }
		/// <summary>
		/// Initialize the <see cref="ExcelConditionalFormattingColorScaleValue"/>
		/// </summary>
		/// <param name="type"></param>
		/// <param name="value"></param>
		/// <param name="formula"></param>
		/// <param name="ruleType"></param>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		/// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingIconDataBarValue(
			eExcelConditionalFormattingValueObjectType type,
			double value,
			string formula,
			eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            int priority,
			ExcelWorksheet worksheet,
			XmlNamespaceManager namespaceManager)
			: this(
				type,
				value,
				formula,
				ruleType,
                address,
                priority,
				worksheet,
				null,
				namespaceManager)
		{
            
		}
		/// <summary>
		/// Initialize the <see cref="ExcelConditionalFormattingColorScaleValue"/>
		/// </summary>
		/// <param name="type"></param>
		/// <param name="color"></param>
		/// <param name="ruleType"></param>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		/// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingIconDataBarValue(
			eExcelConditionalFormattingValueObjectType type,
			Color color,
			eExcelConditionalFormattingRuleType ruleType,
            ExcelAddress address,
            int priority,
			ExcelWorksheet worksheet,
			XmlNamespaceManager namespaceManager)
			: this(
				type,
				0,
				null,
				ruleType,
                address,
                priority,
				worksheet,
				null,
				namespaceManager)
		{
		}
		#endregion Constructors

		/****************************************************************************************/

		#region Methods
        #endregion

        /****************************************************************************************/

		#region Exposed Properties

		/// <summary>
		/// 
		/// </summary>
		internal eExcelConditionalFormattingRuleType RuleType
		{
			get { return _ruleType; }
			set { _ruleType = value; }
		}

		/// <summary>
		/// 
		/// </summary>
		public eExcelConditionalFormattingValueObjectType Type
		{
			get
			{
				var typeAttribute = GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.TypeAttribute);

				return ExcelConditionalFormattingValueObjectType.GetTypeByAttrbiute(typeAttribute);
			}
			set
			{
                if ((_ruleType==eExcelConditionalFormattingRuleType.ThreeIconSet || _ruleType==eExcelConditionalFormattingRuleType.FourIconSet || _ruleType==eExcelConditionalFormattingRuleType.FiveIconSet) &&
                    (value == eExcelConditionalFormattingValueObjectType.Min || value == eExcelConditionalFormattingValueObjectType.Max))
                {
                    throw(new ArgumentException("Value type can't be Min or Max for icon sets"));
                }
                SetXmlNodeString(ExcelConditionalFormattingConstants.Paths.TypeAttribute, value.ToString().ToLower(CultureInfo.InvariantCulture));                
			}
		}

		/// <summary>
		/// Get/Set the 'cfvo' node @val attribute
		/// </summary>
		public Double Value
		{
			get
			{
                if ((Type == eExcelConditionalFormattingValueObjectType.Num)
                    || (Type == eExcelConditionalFormattingValueObjectType.Percent)
                    || (Type == eExcelConditionalFormattingValueObjectType.Percentile))
                {
                    return GetXmlNodeDouble(ExcelConditionalFormattingConstants.Paths.ValAttribute);
                }
                else
                {
                    return 0;
                }
            }
			set
			{
				string valueToStore = string.Empty;

				// Only some types use the @val attribute
				if ((Type == eExcelConditionalFormattingValueObjectType.Num)
					|| (Type == eExcelConditionalFormattingValueObjectType.Percent)
					|| (Type == eExcelConditionalFormattingValueObjectType.Percentile))
				{
					valueToStore = value.ToString(CultureInfo.InvariantCulture);
				}

                SetXmlNodeString(ExcelConditionalFormattingConstants.Paths.ValAttribute, valueToStore);
			}
		}

		/// <summary>
		/// Get/Set the Formula of the Object Value (uses the same attribute as the Value)
		/// </summary>
		public string Formula
		{
			get
			{
				// Return empty if the Object Value type is not Formula
				if (Type != eExcelConditionalFormattingValueObjectType.Formula)
				{
					return string.Empty;
				}

				// Excel stores the formula in the @val attribute
				return GetXmlNodeString(ExcelConditionalFormattingConstants.Paths.ValAttribute);
			}
			set
			{
				// Only store the formula if the Object Value type is Formula
				if (Type == eExcelConditionalFormattingValueObjectType.Formula)
				{
                    SetXmlNodeString(ExcelConditionalFormattingConstants.Paths.ValAttribute, value);
				}
			}
		}
		#endregion Exposed Properties

		/****************************************************************************************/
	}
}