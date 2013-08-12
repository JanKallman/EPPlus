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
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using System.Collections;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    public abstract class FunctionCompiler
    {
        protected ExcelFunction Function
        {
            get;
            private set;
        }

        public FunctionCompiler(ExcelFunction function)
        {
            Require.That(function).Named("function").IsNotNull();
            Function = function;
        }

        protected void BuildFunctionArguments(object result, List<FunctionArgument> args)
        {
            var objects = result as IEnumerable<object>;
            if (objects != null)
            {
                var argList = new List<FunctionArgument>();
                foreach (var arg in objects)
                {
                    BuildFunctionArguments(arg, argList);
                }
                args.Add(new FunctionArgument(argList));
            }
            else
            {
                args.Add(new FunctionArgument(result));
            }
        }

        public abstract CompileResult Compile(IEnumerable<Expression> children, ParsingContext context);
    }
}
