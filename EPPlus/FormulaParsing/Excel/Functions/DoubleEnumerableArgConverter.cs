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
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public class DoubleEnumerableArgConverter : CollectionFlattener<double>
    {
        public virtual IEnumerable<double> ConvertArgs(bool ignoreHidden, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
                {
                    if (arg.IsExcelRange)
                    {
                        foreach (var cell in arg.ValueAsRangeInfo)
                        {
                            if(!ignoreErrors && cell.IsExcelError) throw new ExcelErrorValueException(ExcelErrorValue.Parse(cell.Value.ToString()));
                            if (!CellStateHelper.ShouldIgnore(ignoreHidden, cell, context) && ConvertUtil.IsNumeric(cell.Value))
                            {
                                argList.Add(cell.ValueDouble);
                            }       
                        }
                    }
                    else
                    {
                        if(!ignoreErrors && arg.ValueIsExcelError) throw new ExcelErrorValueException(arg.ValueAsExcelErrorValue);
                        if (ConvertUtil.IsNumeric(arg.Value) && !CellStateHelper.ShouldIgnore(ignoreHidden, arg, context))
                        {
                            argList.Add(ConvertUtil.GetValueDouble(arg.Value));
                        }
                    }
                });
        }

        public virtual IEnumerable<double> ConvertArgsIncludingOtherTypes(IEnumerable<FunctionArgument> arguments)
        {
            return base.FuncArgsToFlatEnumerable(arguments, (arg, argList) =>
            {
                //var cellInfo = arg.Value as EpplusExcelDataProvider.CellInfo;
                //var value = cellInfo != null ? cellInfo.Value : arg.Value;
                if (arg.Value is ExcelDataProvider.IRangeInfo)
                {
                    foreach (var cell in (ExcelDataProvider.IRangeInfo)arg.Value)
                    {
                        argList.Add(cell.ValueDoubleLogical);
                    }
                }
                else
                {
                    if (arg.Value is double || arg.Value is int || arg.Value is bool)
                    {
                        argList.Add(Convert.ToDouble(arg.Value));
                    }
                    else if (arg.Value is string)
                    {
                        argList.Add(0d);
                    }
                }
            });
        }
    }
}
