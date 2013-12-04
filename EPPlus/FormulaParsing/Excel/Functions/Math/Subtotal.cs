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
    public class Subtotal : ExcelFunction
    {
        private Dictionary<int, HiddenValuesHandlingFunction> _functions = new Dictionary<int, HiddenValuesHandlingFunction>();
        
        public Subtotal()
        {
            Initialize();
        }

        private void Initialize()
        {
            _functions[1] = new Average();
            _functions[2] = new Count();
            _functions[3] = new CountA();
            _functions[4] = new Max();
            _functions[5] = new Min();
            _functions[6] = new Product();
            _functions[7] = new Stdev();
            _functions[8] = new StdevP();
            _functions[9] = new Sum();
            _functions[10] = new Var();
            _functions[11] = new VarP();

            AddHiddenValueHandlingFunction(new Average(), 101);
            AddHiddenValueHandlingFunction(new Count(), 102);
            AddHiddenValueHandlingFunction(new CountA(), 103);
            AddHiddenValueHandlingFunction(new Max(), 104);
            AddHiddenValueHandlingFunction(new Min(), 105);
            AddHiddenValueHandlingFunction(new Product(), 106);
            AddHiddenValueHandlingFunction(new Stdev(), 107);
            AddHiddenValueHandlingFunction(new StdevP(), 108);
            AddHiddenValueHandlingFunction(new Sum(), 109);
            AddHiddenValueHandlingFunction(new Var(), 110);
            AddHiddenValueHandlingFunction(new VarP(), 111);
        }

        private void AddHiddenValueHandlingFunction(HiddenValuesHandlingFunction func, int funcNum)
        {
            func.IgnoreHiddenValues = true;
            _functions[funcNum] = func;
        }

        public override void BeforeInvoke(ParsingContext context)
        {
            base.BeforeInvoke(context);
            context.Scopes.Current.IsSubtotal = true;
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var funcNum = ArgToInt(arguments, 0);
            if (context.Scopes.Current.Parent != null && context.Scopes.Current.Parent.IsSubtotal)
            {
                return CreateResult(0d, DataType.Decimal);
            }
            var actualArgs = arguments.Skip(1);
            ExcelFunction function = null;
            function = GetFunctionByCalcType(funcNum);
            var compileResult = function.Execute(actualArgs, context);
            compileResult.IsResultOfSubtotal = true;
            return compileResult;
        }

        private ExcelFunction GetFunctionByCalcType(int funcNum)
        {
            if (!_functions.ContainsKey(funcNum))
            {
                throw new ArgumentException("Invalid funcNum " + funcNum + ", valid ranges are 1-11 and 101-111");
            }
            return _functions[funcNum];
        }
    }
}
