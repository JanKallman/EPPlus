﻿/* Copyright (C) 2011  Jan Källman
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
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
    public class IsError : ErrorHandlingFunction
    {
        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            if (arguments == null || arguments.Count() == 0)
            {
                return CreateResult(false, DataType.Boolean);
            }
            foreach (var argument in arguments)
            {
                if (argument.Value is ExcelDataProvider.IRangeInfo)
                {
                    var r = (ExcelDataProvider.IRangeInfo)argument.Value;
                    if (ExcelErrorValue.Values.IsErrorValue(r.GetValue(r.Address._fromRow, r.Address._fromCol)))
                    {
                        return CreateResult(true, DataType.Boolean);
                    }
                }
                else
                {
                    if (ExcelErrorValue.Values.IsErrorValue(argument.Value))
                    {
                        return CreateResult(true, DataType.Boolean);
                    }
                }                
            }
            return CreateResult(false, DataType.Boolean);
        }

        public override CompileResult HandleError(string errorCode)
        {
            return CreateResult(true, DataType.Boolean);
        }
    }
}
