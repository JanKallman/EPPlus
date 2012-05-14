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
 * Author          Change						              Date
 * ******************************************************************************
 * Eyal Seagull    Conditional Formatting         2012-04-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Contracts
{
	/// <summary>
	/// Interface for conditional formatting rule
	/// </summary>
	public interface IExcelConditionalFormattingRule
	{
    /// <summary>
    /// The 'cfRule' XML node
    /// </summary>
    XmlNode Node { get; }

    /// <summary>
    /// Type of conditional formatting rule. ST_CfType §18.18.12.
    /// </summary>
    eExcelConditionalFormattingRuleType Type { get; }

    /// <summary>
    /// <para>Range over which these conditional formatting rules apply.</para>
    /// <para>The possible values for this attribute are defined by the
    /// ST_Sqref simple type (§18.18.76).</para>
    /// </summary>
    ExcelAddress Address { get; set; }

		/// <summary>
		/// The priority of this conditional formatting rule. This value is used to determine
		/// which format should be evaluated and rendered. Lower numeric values are higher
		/// priority than higher numeric values, where 1 is the highest priority.
		/// </summary>
    int Priority { get; set; }

    /// <summary>
    /// If this flag is 1, no rules with lower priority shall be applied over this rule,
    /// when this rule evaluates to true.
    /// </summary>
    bool StopIfTrue { get; set; }

    /// <summary>
    /// <para>This is an index to a dxf element in the Styles Part indicating which cell
    /// formatting to apply when the conditional formatting rule criteria is met.</para>
    /// <para>The possible values for this attribute are defined by the ST_DxfId simple type
    /// (§18.18.25).</para>
    /// </summary>
    int DxfId { get; set; }
  }
}