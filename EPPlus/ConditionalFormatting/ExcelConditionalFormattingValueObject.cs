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
	public class ExcelConditionalFormattingValueObject
		: XmlHelper
	{
		/****************************************************************************************/

		#region Private Properties
		private eExcelConditionalFormattingValueObjectPosition _position;
		private eExcelConditionalFormattingRuleType _ruleType;
		private ExcelWorksheet _worksheet;
		#endregion Private Properties

		/****************************************************************************************/

		#region Constructors
    /// <summary>
    /// Initialize the cfvo (§18.3.1.11) node
    /// </summary>
    /// <param name="position"></param>
    /// <param name="type"></param>
    /// <param name="color"></param>
    /// <param name="value"></param>
    /// <param name="formula"></param>
    /// <param name="ruleType"></param>
    /// <param name="address"></param>
    /// <param name="priority"></param>
    /// <param name="worksheet"></param>
    /// <param name="itemElementNode">The cfvo (§18.3.1.11) node parent. Can be any of the following:
    /// colorScale (§18.3.1.16); dataBar (§18.3.1.28); iconSet (§18.3.1.49)</param>
    /// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingValueObject(
			eExcelConditionalFormattingValueObjectPosition position,
			eExcelConditionalFormattingValueObjectType type,
			Color color,
			double value,
			string formula,
			eExcelConditionalFormattingRuleType ruleType,
      ExcelAddress address,
      int priority,
			ExcelWorksheet worksheet,
			XmlNode itemElementNode,
			XmlNamespaceManager namespaceManager)
			: base(
				namespaceManager,
				itemElementNode)
		{
			Require.Argument(priority).IsInRange(1, int.MaxValue, "priority");
			Require.Argument(address).IsNotNull("address");
			Require.Argument(worksheet).IsNotNull("worksheet");

			// Save the worksheet for private methods to use
			_worksheet = worksheet;

			// Schema order list
			SchemaNodeOrder = new string[]
			{
        ExcelConditionalFormattingConstants.Nodes.Cfvo,
        ExcelConditionalFormattingConstants.Nodes.Color
			};

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

			// Point to the <cfvo> parent node (<colorScale>, <dataBar> or <iconSet>)
      // This is different than normal, as TopNode does not point to the node itself but to
      // its PARENT. Later, in the CreateNodeByOrdem method the TopNode will be updated.
      TopNode = itemElementNode;

			// Save the attributes
			Position = position;
			RuleType = ruleType;
			Type = type;
			Color = color;
			Value = value;
			Formula = formula;
		}

		/// <summary>
		/// Initialize the <see cref="ExcelConditionalFormattingValueObject"/>
		/// </summary>
		/// <param name="position"></param>
		/// <param name="type"></param>
		/// <param name="color"></param>
		/// <param name="value"></param>
		/// <param name="formula"></param>
		/// <param name="ruleType"></param>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		/// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingValueObject(
			eExcelConditionalFormattingValueObjectPosition position,
			eExcelConditionalFormattingValueObjectType type,
			Color color,
			double value,
			string formula,
			eExcelConditionalFormattingRuleType ruleType,
      ExcelAddress address,
      int priority,
			ExcelWorksheet worksheet,
			XmlNamespaceManager namespaceManager)
			: this(
				position,
				type,
				color,
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
		/// Initialize the <see cref="ExcelConditionalFormattingValueObject"/>
		/// </summary>
		/// <param name="position"></param>
		/// <param name="type"></param>
		/// <param name="color"></param>
		/// <param name="ruleType"></param>
		/// <param name="priority"></param>
		/// <param name="address"></param>
		/// <param name="worksheet"></param>
		/// <param name="namespaceManager"></param>
		internal ExcelConditionalFormattingValueObject(
			eExcelConditionalFormattingValueObjectPosition position,
			eExcelConditionalFormattingValueObjectType type,
			Color color,
			eExcelConditionalFormattingRuleType ruleType,
      ExcelAddress address,
      int priority,
			ExcelWorksheet worksheet,
			XmlNamespaceManager namespaceManager)
			: this(
				position,
				type,
				color,
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
		/// <summary>
		/// Get the node order (1, 2 ou 3) according to the Position (Low, Middle and High)
		/// and the Rule Type (TwoColorScale ou ThreeColorScale).
		/// </summary>
		/// <returns></returns>
		private int GetNodeOrder()
		{
			return ExcelConditionalFormattingValueObjectType.GetOrderByPosition(
				Position,
				RuleType);
		}

		/// <summary>
		/// Create the 'cfvo'/'color' nodes in the right order. They should appear like this:
		///		"cfvo"   --> Low Value (value object)
		///		"cfvo"   --> Middle Value (value object)
		///		"cfvo"   --> High Value (value object)
		///		"color"  --> Low Value (color)
		///		"color"  --> Middle Value (color)
		///		"color"  --> High Value (color)
		/// </summary>
		/// <param name="nodeType"></param>
		/// <param name="attributePath"></param>
		/// <param name="attributeValue"></param>
		private void CreateNodeByOrdem(
			eExcelConditionalFormattingValueObjectNodeType nodeType,
			string attributePath,
			string attributeValue)
		{
      // Save the current TopNode
      XmlNode currentTopNode = TopNode;

      string nodePath = ExcelConditionalFormattingValueObjectType.GetNodePathByNodeType(nodeType);
			int nodeOrder = GetNodeOrder();
			eNodeInsertOrder nodeInsertOrder = eNodeInsertOrder.SchemaOrder;
			XmlNode referenceNode = null;

			if (nodeOrder > 1)
			{
				// Find the node just before the one we need to include
				referenceNode = TopNode.SelectSingleNode(
					string.Format(
						"{0}[position()={1}]",
					// {0}
						nodePath,
					// {1}
						nodeOrder - 1),
					_worksheet.NameSpaceManager);

				// Only if the prepend node exists than insert after
				if (referenceNode != null)
				{
					nodeInsertOrder = eNodeInsertOrder.After;
				}
			}

			// Create the node in the right order
			var node = CreateComplexNode(
				TopNode,
				string.Format(
					"{0}[position()={1}]",
				// {0}
					nodePath,
				// {1}
					nodeOrder),
				nodeInsertOrder,
				referenceNode);

      // Point to the new node as the temporary TopNode (we need it for the XmlHelper functions)
      TopNode = node;

      // Add/Remove the attribute (if the attributeValue is empty then it will be removed)
      SetXmlNodeString(
        node,
        attributePath,
        attributeValue,
        true);

      // Point back to the <cfvo>/<color> parent node
      TopNode = currentTopNode;
		}
		#endregion Methos

		/****************************************************************************************/

		#region Exposed Properties
		/// <summary>
		/// 
		/// </summary>
		internal eExcelConditionalFormattingValueObjectPosition Position
		{
			get { return _position; }
			set { _position = value; }
		}

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
				var typeAttribute = GetXmlNodeString(
					string.Format(
						"{0}[position()={1}]/{2}",
					// {0}
						ExcelConditionalFormattingConstants.Paths.Cfvo,
					// {1}
						GetNodeOrder(),
					// {2}
						ExcelConditionalFormattingConstants.Paths.TypeAttribute));

				return ExcelConditionalFormattingValueObjectType.GetTypeByAttrbiute(typeAttribute);
			}
			set
			{
				CreateNodeByOrdem(
					eExcelConditionalFormattingValueObjectNodeType.Cfvo,
					ExcelConditionalFormattingConstants.Paths.TypeAttribute,
					ExcelConditionalFormattingValueObjectType.GetAttributeByType(value));

				bool removeValAttribute = false;

				// Make sure unnecessary attributes are removed (occures when we change
				// the value object type)
				switch (Type)
				{
					case eExcelConditionalFormattingValueObjectType.Min:
					case eExcelConditionalFormattingValueObjectType.Max:
						removeValAttribute = true;
						break;
				}

				// Check if we need to remove the @val attribute
				if (removeValAttribute)
				{
				  string nodePath = ExcelConditionalFormattingValueObjectType.GetNodePathByNodeType(
						eExcelConditionalFormattingValueObjectNodeType.Cfvo);
				  int nodeOrder = GetNodeOrder();

					// Remove the attribute (removed when the value = '')
				  CreateComplexNode(
				    TopNode,
				    string.Format(
				      "{0}[position()={1}]/{2}=''",
				    // {0}
				      nodePath,
				    // {1}
				      nodeOrder,
				    // {2}
				      ExcelConditionalFormattingConstants.Paths.ValAttribute));
				}
			}
		}

		/// <summary>
		/// 
		/// </summary>
		public Color Color
		{
			get
			{
				// Color Code like "FF5B34F2"
				var colorCode = GetXmlNodeString(
					string.Format(
						"{0}[position()={1}]/{2}",
					// {0}
						ExcelConditionalFormattingConstants.Paths.Color,
					// {1}
						GetNodeOrder(),
					// {2}
						ExcelConditionalFormattingConstants.Paths.RgbAttribute));

				return ExcelConditionalFormattingHelper.ConvertFromColorCode(colorCode);
			}
			set
			{
				// Use the color code to store (Ex. "FF5B35F2")
				CreateNodeByOrdem(
					eExcelConditionalFormattingValueObjectNodeType.Color,
					ExcelConditionalFormattingConstants.Paths.RgbAttribute,
					value.Name.ToUpper());
			}
		}

		/// <summary>
		/// Get/Set the 'cfvo' node @val attribute
		/// </summary>
		public Double Value
		{
			get
			{
				return GetXmlNodeDouble(
					string.Format(
						"{0}[position()={1}]/{2}",
					// {0}
						ExcelConditionalFormattingConstants.Paths.Cfvo,
					// {1}
						GetNodeOrder(),
					// {2}
						ExcelConditionalFormattingConstants.Paths.ValAttribute));
			}
			set
			{
				string valueToStore = string.Empty;

				// Only some types use the @val attribute
				if ((Type == eExcelConditionalFormattingValueObjectType.Num)
					|| (Type == eExcelConditionalFormattingValueObjectType.Percent)
					|| (Type == eExcelConditionalFormattingValueObjectType.Percentile))
				{
					valueToStore = value.ToString();
				}

				CreateNodeByOrdem(
					eExcelConditionalFormattingValueObjectNodeType.Cfvo,
					ExcelConditionalFormattingConstants.Paths.ValAttribute,
					valueToStore);
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
				return GetXmlNodeString(
					string.Format(
						"{0}[position()={1}]/{2}",
					// {0}
						ExcelConditionalFormattingConstants.Paths.Cfvo,
					// {1}
						GetNodeOrder(),
					// {2}
						ExcelConditionalFormattingConstants.Paths.ValAttribute));
			}
			set
			{
				// Only store the formula if the Object Value type is Formula
				if (Type == eExcelConditionalFormattingValueObjectType.Formula)
				{
					CreateNodeByOrdem(
						eExcelConditionalFormattingValueObjectNodeType.Cfvo,
						ExcelConditionalFormattingConstants.Paths.ValAttribute,
						(value == null) ? string.Empty : value.ToString());
				}
			}
		}
		#endregion Exposed Properties

		/****************************************************************************************/
	}
}