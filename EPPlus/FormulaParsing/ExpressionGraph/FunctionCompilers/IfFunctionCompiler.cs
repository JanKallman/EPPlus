/*******************************************************************************
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
 * Mats Alm   		                Added       		        2014-01-27
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers
{
    /// <summary>
    /// Why do the If function require a compiler of its own you might ask;)
    /// 
    /// It is because it only needs to evaluate one of the two last expressions. This
    /// compiler handles this - it ignores the irrelevant expression.
    /// </summary>
    public class IfFunctionCompiler : FunctionCompiler
    {
        public IfFunctionCompiler(ExcelFunction function, ParsingContext context)
            : base(function, context)
        {
            Require.That(function).Named("function").IsNotNull();
            if (!(function is If)) throw new ArgumentException("function must be of type If");
        }

        public override CompileResult Compile(IEnumerable<Expression> children)
        {
            // 2 is allowed, Excel returns FALSE if false is the outcome of the expression
            if(children.Count() < 2) throw new ExcelErrorValueException(eErrorType.Value);
            var args = new List<FunctionArgument>();
            Function.BeforeInvoke(Context);
            var firstChild = children.ElementAt(0);
            var v = firstChild.Compile().Result;

            /****  Handle names and ranges ****/
            if (v is ExcelDataProvider.INameInfo)
            {
                v = ((ExcelDataProvider.INameInfo)v).Value;
            }
            
            if (v is ExcelDataProvider.IRangeInfo)
            {
                var r=((ExcelDataProvider.IRangeInfo)v);
                if(r.GetNCells()>1)
                {
                    throw(new ArgumentException("Logical can't be more than one cell"));
                }
                v = r.GetOffset(0, 0);
            }
            bool boolVal;
            if(v is bool)
            {
                boolVal = (bool)v;
            }
            else if(!Utils.ConvertUtil.TryParseBooleanString(v, out boolVal))
            {
                if(OfficeOpenXml.Utils.ConvertUtil.IsNumeric(v))
                {
                    boolVal = OfficeOpenXml.Utils.ConvertUtil.GetValueDouble(v)!=0;
                }
                else
                {
                    throw (new ArgumentException("Invalid logical test"));
                }
            }
            /****  End Handle names and ranges ****/
            
            args.Add(new FunctionArgument(boolVal));
            if (boolVal)
            {
                var val = children.ElementAt(1).Compile().Result;
                args.Add(new FunctionArgument(val));
                args.Add(new FunctionArgument(null));
            }
            else
            {
                object val;
                var child = children.ElementAtOrDefault(2);
                if (child == null)
                {
                    // if no false expression given, Excel returns false
                    val = false;
                }
                else
                {
                    val = child.Compile().Result;
                }
                args.Add(new FunctionArgument(null));
                args.Add(new FunctionArgument(val));
            }
            return Function.Execute(args, Context);
        }
    }
}
