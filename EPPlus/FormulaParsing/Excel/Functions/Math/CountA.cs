/* Copyright (C) 2011  Jan Källman
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
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class CountA : HiddenValuesHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 1);
            var nItems = 0d;
            Calculate(arguments, context,  ref nItems);
            return CreateResult(nItems, DataType.Integer);
        }

        private void Calculate(IEnumerable<FunctionArgument> items, ParsingContext context, ref double nItems)
        {
            foreach (var item in items)
            {
                if (item.Value is ExcelDataProvider.IRangeInfo)
                {
                    foreach (var c in (ExcelDataProvider.IRangeInfo)item.Value)
                    {
                        if (ShouldIgnore(c, context) == false && ShouldCount(c.Value, item.ExcelStateFlagIsSet(ExcelCellState.HiddenCell)))
                        {
                            nItems++;
                        }
                    }
                }
                else if (item.Value is IEnumerable<FunctionArgument>)
                {
                    Calculate((IEnumerable<FunctionArgument>)item.Value, context, ref nItems);
                }
                else if (ShouldCount(item.Value, item.ExcelStateFlagIsSet(ExcelCellState.HiddenCell)))
                {
                    nItems++;
                }
                
            }
        }
        private bool ShouldCount(object value, bool isHiddenCell)
        {
            if (isHiddenCell)
            {
                return false;
            }
            if (value == null) return false;
            return (!string.IsNullOrEmpty(value.ToString()));
        }
    }
}
