﻿/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
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
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class ExpressionConverter : IExpressionConverter
    {
        public StringExpression ToStringExpression(Expression expression)
        {
            var result = expression.Compile();
            var newExp = new StringExpression(result.Result.ToString());
            newExp.Operator = expression.Operator;
            return newExp;
        }

        public Expression FromCompileResult(CompileResult compileResult)
        {
            switch (compileResult.DataType)
            {
                case DataType.Integer:
                    return compileResult.Result is string
                        ? new IntegerExpression(compileResult.Result.ToString())
                        : new IntegerExpression(Convert.ToDouble(compileResult.Result));
                case DataType.String:
                    return new StringExpression(compileResult.Result.ToString());
                case DataType.Decimal:
                    return compileResult.Result is string
                               ? new DecimalExpression(compileResult.Result.ToString())
                               : new DecimalExpression(((double) compileResult.Result));
                case DataType.Boolean:
                    return compileResult.Result is string
                               ? new BooleanExpression(compileResult.Result.ToString())
                               : new BooleanExpression((bool) compileResult.Result);
                //case DataType.Enumerable:
                //    return 
                case DataType.ExcelError:
                    //throw (new OfficeOpenXml.FormulaParsing.Exceptions.ExcelErrorValueException((ExcelErrorValue)compileResult.Result)); //Added JK
                    return compileResult.Result is string
                        ? new ExcelErrorExpression(compileResult.Result.ToString(),
                            ExcelErrorValue.Parse(compileResult.Result.ToString()))
                        : new ExcelErrorExpression((ExcelErrorValue) compileResult.Result);
                case DataType.Empty:
                   return new IntegerExpression(0); //Added JK
                case DataType.Time:
                case DataType.Date:
                    return new DecimalExpression((double)compileResult.Result);

            }
            return null;
        }

        private static IExpressionConverter _instance;
        public static IExpressionConverter Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ExpressionConverter();
                }
                return _instance;
            }
        }
    }
}
