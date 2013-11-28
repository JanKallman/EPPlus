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
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    public class IndexToAddressTranslator
    {
        public IndexToAddressTranslator(ExcelDataProvider excelDataProvider)
            : this(excelDataProvider, ExcelReferenceType.AbsoluteRowAndColumn)
        {

        }

        public IndexToAddressTranslator(ExcelDataProvider excelDataProvider, ExcelReferenceType referenceType)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            _excelDataProvider = excelDataProvider;
            _excelReferenceType = referenceType;
        }

        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ExcelReferenceType _excelReferenceType;

        protected internal static string GetColumnLetter(int iColumnNumber, bool fixedCol)
        {

            if (iColumnNumber < 1)
            {
                //throw new Exception("Column number is out of range");
                return "#REF!";
            }

            string sCol = "";
            do
            {
                sCol = ((char)('A' + ((iColumnNumber - 1) % 26))) + sCol;
                iColumnNumber = (iColumnNumber - ((iColumnNumber - 1) % 26)) / 26;
            }
            while (iColumnNumber > 0);
            return fixedCol ? "$" + sCol : sCol;
        }

        public string ToAddress(int col, int row)
        {
            var fixedCol = _excelReferenceType == ExcelReferenceType.AbsoluteRowAndColumn ||
                           _excelReferenceType == ExcelReferenceType.RelativeRowAbsolutColumn;
            var colString = GetColumnLetter(col, fixedCol);
            return colString + GetRowNumber(row);
        }

        private string GetRowNumber(int rowNo)
        {
            var retVal = rowNo < (_excelDataProvider.ExcelMaxRows) ? rowNo.ToString() : string.Empty;
            if (!string.IsNullOrEmpty(retVal))
            {
                switch (_excelReferenceType)
                {
                    case ExcelReferenceType.AbsoluteRowAndColumn:
                    case ExcelReferenceType.AbsoluteRowRelativeColumn:
                        return "$" + retVal;
                    default:
                        return retVal;
                }
            }
            return retVal;
        }
    }
}
