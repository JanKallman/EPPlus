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
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
    public class AverageIf : HiddenValuesHandlingFunction
    {
        private readonly ExpressionEvaluator _expressionEvaluator;

        public AverageIf()
            : this(new ExpressionEvaluator())
        {

        }

        public AverageIf(ExpressionEvaluator evaluator)
        {
            Require.That(evaluator).Named("evaluator").IsNotNull();
            _expressionEvaluator = evaluator;
        }

        private bool Evaluate(object obj, string expression)
        {
            double? candidate = default(double?);
            if (IsNumeric(obj))
            {
                candidate = ConvertUtil.GetValueDouble(obj);
            }
            if (candidate.HasValue)
            {
                return _expressionEvaluator.Evaluate(candidate.Value, expression);
            }
            return _expressionEvaluator.Evaluate(obj, expression);
        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);
            var firstArg = arguments.ElementAt(0);
            var args = firstArg.Value as IEnumerable<FunctionArgument>;
            if (args == null && firstArg.IsExcelRange)
            {
                args = new List<FunctionArgument>(){ firstArg };
            }
            var criteria = arguments.ElementAt(1).ValueFirst != null ? ArgToString(arguments, 1) : null;
            var retVal = 0d;
            if (arguments.Count() > 2)
            {
                var secondArg = arguments.ElementAt(2);
                var lookupRange = secondArg.Value as IEnumerable<FunctionArgument>;
                if (lookupRange == null && secondArg.IsExcelRange)
                {
                    lookupRange = new List<FunctionArgument>() {secondArg};
                }
                retVal = CalculateWithLookupRange(args, criteria, lookupRange, context);
            }
            else
            {
                retVal = CalculateSingleRange(args, criteria, context);
            }
            return CreateResult(retVal, DataType.Decimal);
        }

        private double CalculateWithLookupRange(IEnumerable<FunctionArgument> range, string criteria, IEnumerable<FunctionArgument> sumRange, ParsingContext context)
        {
            var retVal = 0d;
            var nMatches = 0;
            var flattenedRange = ArgsToObjectEnumerable(false, range, context);
            var flattenedSumRange = ArgsToDoubleEnumerable(sumRange, context);
            for (var x = 0; x < flattenedRange.Count(); x++)
            {
                var candidate = flattenedSumRange.ElementAt(x);
                if (criteria != null && Evaluate(flattenedRange.ElementAt(x), criteria))
                {
                    nMatches++;
                    retVal += candidate;
                }
            }
            return Divide(retVal, nMatches);
        }

        private double CalculateSingleRange(IEnumerable<FunctionArgument> args, string expression, ParsingContext context)
        {
            var retVal = 0d;
            var nMatches = 0;
            var flattendedRange = ArgsToDoubleEnumerable(args, context);
            var candidates = flattendedRange as double[] ?? flattendedRange.ToArray();
            foreach (var candidate in candidates)
            {
                if (expression != null && Evaluate(candidate, expression))
                {
                    retVal += candidate;
                    nMatches++;
                }
            }
            return Divide(retVal, nMatches);
        }
    }
}
