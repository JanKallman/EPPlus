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
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
    /// <summary>
    /// Handles translations from Spreadsheet addresses to 0-based numeric index.
    /// </summary>
    public class AddressTranslator
    {
        public enum RangeCalculationBehaviour
        {
            FirstPart,
            LastPart
        }

        private readonly ExcelDataProvider _excelDataProvider;

        public AddressTranslator(ExcelDataProvider excelDataProvider)
        {
            Require.That(excelDataProvider).Named("excelDataProvider").IsNotNull();
            _excelDataProvider = excelDataProvider;
        }

        /// <summary>
        /// Translates an address in format "A1" to col- and rowindex.
        /// 
        /// If the supplied address is a range, the address of the first part will be calculated.
        /// </summary>
        /// <param name="address"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        public virtual void ToColAndRow(string address, out int col, out int row)
        {
            ToColAndRow(address, out col, out row, RangeCalculationBehaviour.FirstPart);
        }

        /// <summary>
        /// Translates an address in format "A1" to col- and rowindex.
        /// </summary>
        /// <param name="address"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <param name="behaviour"></param>
        public virtual void ToColAndRow(string address, out int col, out int row, RangeCalculationBehaviour behaviour)
        {
            address = address.ToUpper();
            var alphaPart = GetAlphaPart(address);
            col = 0;
            var nLettersInAlphabet = 26;
            for (int x = 0; x < alphaPart.Length; x++)
            {
                var pos = alphaPart.Length - x - 1;
                var currentNumericValue = GetNumericAlphaValue(alphaPart[x]);
                col += (nLettersInAlphabet * pos * currentNumericValue);
                if (pos == 0)
                {
                    col += currentNumericValue;
                }
            }
            //col--;
            //row = GetIntPart(address) - 1 ?? GetRowIndexByBehaviour(behaviour);
            row = GetIntPart(address) ?? GetRowIndexByBehaviour(behaviour);

        }

        private int GetRowIndexByBehaviour(RangeCalculationBehaviour behaviour)
        {
            if (behaviour == RangeCalculationBehaviour.FirstPart)
            {
                return 1;
            }
            return _excelDataProvider.ExcelMaxRows;
        }

        private int GetNumericAlphaValue(char c)
        {
            return (int)c - 64;
        }

        private string GetAlphaPart(string address)
        {
            return Regex.Match(address, "[A-Z]+").Value;
        }

        private int? GetIntPart(string address)
        {
            if (Regex.IsMatch(address, "[0-9]+"))
            {
                return int.Parse(Regex.Match(address, "[0-9]+").Value);
            }
            return null;
        }
    }
}
