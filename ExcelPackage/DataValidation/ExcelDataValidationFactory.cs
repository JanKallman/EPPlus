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
 * Mats Alm   		                Added       		        2011-01-01
 * Jan Källman		                License changed GPL-->LGPL  2011-12-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// Factory class for ExcelDataValidation.
    /// </summary>
    internal static class ExcelDataValidationFactory
    {
        /// <summary>
        /// Creates an instance of <see cref="ExcelDataValidation"/> out of the given parameters.
        /// </summary>
        /// <param name="type"></param>
        /// <param name="worksheet"></param>
        /// <param name="address"></param>
        /// <param name="itemElementNode"></param>
        /// <returns></returns>
        public static ExcelDataValidation Create(ExcelDataValidationType type, ExcelWorksheet worksheet, string address, XmlNode itemElementNode)
        {
            Require.Argument(type).IsNotNull("validationType");
            switch (type.Type)
            {
                case eDataValidationType.TextLength:
                case eDataValidationType.Whole:
                    return new ExcelDataValidationInt(worksheet, address, type, itemElementNode);
                case eDataValidationType.Decimal:
                    return new ExcelDataValidationDecimal(worksheet, address, type, itemElementNode);
                case eDataValidationType.List:
                    return new ExcelDataValidationList(worksheet, address, type, itemElementNode);
                case eDataValidationType.DateTime:
                    return new ExcelDataValidationDateTime(worksheet, address, type, itemElementNode);
                case eDataValidationType.Time:
                    return new ExcelDataValidationTime(worksheet, address, type, itemElementNode);
                case eDataValidationType.Custom:
                    return new ExcelDataValidationCustom(worksheet, address, type, itemElementNode);
                default:
                    throw new InvalidOperationException("Non supported validationtype: " + type.Type.ToString());
            }
        }
    }
}
