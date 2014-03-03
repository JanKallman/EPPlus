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

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class FunctionExpression : AtomicExpression
    {
        public FunctionExpression(string expression, ParsingContext parsingContext)
            : base(expression)
        {
            _parsingContext = parsingContext;
            base.AddChild(new FunctionArgumentExpression(this));
        }

        private readonly ParsingContext _parsingContext;
        private readonly FunctionCompilerFactory _functionCompilerFactory = new FunctionCompilerFactory();


        public override CompileResult Compile()
        {
            try
            {
                var function = _parsingContext.Configuration.FunctionRepository.GetFunction(ExpressionString);
                var compiler = _functionCompilerFactory.Create(function);
                return compiler.Compile(HasChildren ? Children : Enumerable.Empty<Expression>(), _parsingContext);
            }
            catch (ExcelErrorValueException e)
            {
                return new CompileResult(e.ErrorValue, DataType.ExcelError);
            }
            catch (InvalidOperationException ioe)
            {
                throw;
            }
            
        }

        public override Expression PrepareForNextChild()
        {
            return base.AddChild(new FunctionArgumentExpression(this));
        }

        private bool _childAdded;

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
