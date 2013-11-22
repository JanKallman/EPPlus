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

        const int MaxAlphabetIndex = 25;
        const int NLettersInAlphabet = 26;
        private readonly ExcelDataProvider _excelDataProvider;
        private readonly ExcelReferenceType _excelReferenceType;

        public string ToAddress(int col, int row)
        {
            if (col <= MaxAlphabetIndex)
            {
                return string.Concat(GetColumn(IntToChar(col)), GetRowNumber(row));
            }
            else if (col <= (Math.Pow(NLettersInAlphabet, 2) + NLettersInAlphabet))
            {
                var firstChar = col / NLettersInAlphabet;
                var secondChar = col - (NLettersInAlphabet * firstChar);
                return string.Concat(GetColumn(IntToChar(firstChar), IntToChar(secondChar)), GetRowNumber(row));
            }
            else if(col < (Math.Pow(NLettersInAlphabet, 3) + NLettersInAlphabet))
            {
                var x = NLettersInAlphabet * NLettersInAlphabet;
                var rest = col - x;
                var firstChar = col / x;
                var secondChar = rest / NLettersInAlphabet;
                var thirdChar = rest % NLettersInAlphabet;
                return string.Concat(GetColumn(IntToChar(firstChar), IntToChar(secondChar), IntToChar(thirdChar)), GetRowNumber(row));
            }
            throw new InvalidOperationException("ExcelFormulaParser does not the supplied number of columns " + col);
        }

        private string GetColumn(params char[] chars)
        {
            var retVal = new StringBuilder().Append(chars).ToString();
            switch (_excelReferenceType)
            {
                case ExcelReferenceType.AbsoluteRowAndColumn:
                case ExcelReferenceType.RelativeRowAbsolutColumn:
                    return "$" + retVal;
                default:
                    return retVal;
            }
        }

        private char IntToChar(int i)
        {
            return (char)(i + 64);
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
