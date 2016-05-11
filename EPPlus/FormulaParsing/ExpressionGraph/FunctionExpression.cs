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
using System.Globalization;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    /// <summary>
    /// Expression that handles execution of a function.
    /// </summary>
    public class FunctionExpression : AtomicExpression
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="expression">should be the of the function</param>
        /// <param name="parsingContext"></param>
        /// <param name="isNegated">True if the numeric result of the function should be negated.</param>
        public FunctionExpression(string expression, ParsingContext parsingContext, bool isNegated)
            : base(expression)
        {
            _parsingContext = parsingContext;
            _functionCompilerFactory = new FunctionCompilerFactory(parsingContext.Configuration.FunctionRepository);
            _isNegated = isNegated;
            base.AddChild(new FunctionArgumentExpression(this));
        }

        private readonly ParsingContext _parsingContext;
        private readonly FunctionCompilerFactory _functionCompilerFactory;
        private readonly bool _isNegated;


        public override CompileResult Compile()
        {
            try
            {
                var function = _parsingContext.Configuration.FunctionRepository.GetFunction(ExpressionString);
                if (function == null)
                {
                    if (_parsingContext.Debug)
                    {
                        _parsingContext.Configuration.Logger.Log(_parsingContext, string.Format("'{0}' is not a supported function", ExpressionString));
                    }
                    return new CompileResult(ExcelErrorValue.Create(eErrorType.Name), DataType.ExcelError);
                }
                if (_parsingContext.Debug)
                {
                    _parsingContext.Configuration.Logger.LogFunction(ExpressionString);
                }
                var compiler = _functionCompilerFactory.Create(function);
                var result = compiler.Compile(HasChildren ? Children : Enumerable.Empty<Expression>(), _parsingContext);
                if (_isNegated)
                {
                    if (!result.IsNumeric)
                    {
                        if (_parsingContext.Debug)
                        {
                            var msg = string.Format("Trying to negate a non-numeric value ({0}) in function '{1}'",
                                result.Result, ExpressionString);
                            _parsingContext.Configuration.Logger.Log(_parsingContext, msg);
                        }
                        return new CompileResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
                    }
                    return new CompileResult(result.ResultNumeric * -1, result.DataType);
                }
                return result;
            }
            catch (ExcelErrorValueException e)
            {
                if (_parsingContext.Debug)
                {
                    _parsingContext.Configuration.Logger.Log(_parsingContext, e);
                }
                return new CompileResult(e.ErrorValue, DataType.ExcelError);
            }
            
        }

        public override Expression PrepareForNextChild()
        {
            return base.AddChild(new FunctionArgumentExpression(this));
        }

        public override bool HasChildren
        {
            get
            {
                return (Children.Any() && Children.First().Children.Any());
            }
        }

        public override Expression AddChild(Expression child)
        {
            Children.Last().AddChild(child);
            return child;
        }
    }
}
