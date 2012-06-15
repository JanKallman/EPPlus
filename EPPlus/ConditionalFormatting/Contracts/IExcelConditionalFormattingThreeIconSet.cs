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

using OfficeOpenXml.ConditionalFormatting;

namespace OfficeOpenXml.ConditionalFormatting.Contracts
{
	/// <summary>
	/// IExcelConditionalFormattingThreeIconSet
	/// </summary>
	public interface IExcelConditionalFormattingThreeIconSet<T, I>
    : IExcelConditionalFormattingIconSetGroup<T>
	{
		#region Public Properties
    /// <summary>
    /// Icon1 (part of the 3, 4 ou 5 Icon Set)
    /// </summary>
    I Icon1 { get; set; }

    /// <summary>
    /// Icon2 (part of the 3, 4 ou 5 Icon Set)
    /// </summary>
    I Icon2 { get; set; }

    /// <summary>
    /// Icon3 (part of the 3, 4 ou 5 Icon Set)
    /// </summary>
    I Icon3 { get; set; }
    #endregion Public Properties
	}
}