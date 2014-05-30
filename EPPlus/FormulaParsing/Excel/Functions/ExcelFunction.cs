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
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using System.Collections;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    /// <summary>
    /// Base class for Excel function implementations.
    /// </summary>
    public abstract class ExcelFunction
    {
        public ExcelFunction()
            : this(new ArgumentCollectionUtil(), new ArgumentParsers(), new CompileResultValidators())
        {

        }

        public ExcelFunction(
            ArgumentCollectionUtil argumentCollectionUtil, 
            ArgumentParsers argumentParsers,
            CompileResultValidators compileResultValidators)
        {
            _argumentCollectionUtil = argumentCollectionUtil;
            _argumentParsers = argumentParsers;
            _compileResultValidators = compileResultValidators;
        }

        private readonly ArgumentCollectionUtil _argumentCollectionUtil;
        private readonly ArgumentParsers _argumentParsers;
        private readonly CompileResultValidators _compileResultValidators;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arguments">Arguments to the function, each argument can contain primitive types, lists or <see cref="ExcelDataProvider.IRangeInfo">Excel ranges</see></param>
        /// <param name="context">The <see cref="ParsingContext"/> contains various data that can be useful in functions.</param>
        /// <returns>A <see cref="CompileResult"/> containing the calculated value</returns>
        public abstract CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context);

        /// <summary>
        /// If overridden, this method is called before Execute is called.
        /// </summary>
        /// <param name="context"></param>
        public virtual void BeforeInvoke(ParsingContext context) { }

        public virtual bool IsLookupFuction 
        { 
            get 
            { 
                return false; 
            } 
        }

        public virtual bool IsErrorHandlingFunction
        {
            get
            {
                return false;
            }
        }
        
        /// <summary>
        /// Used for some Lookupfunctions to indicate that function arguments should
        /// not be compiled before the function is called.
        /// </summary>
        public bool SkipArgumentEvaluation { get; set; }
        protected object GetFirstValue(IEnumerable<FunctionArgument> val)
        {
            var arg = ((IEnumerable<FunctionArgument>)val).FirstOrDefault();
            if(arg.Value is ExcelDataProvider.IRangeInfo)
            {
                //var r=((ExcelDataProvider.IRangeInfo)arg);
                var r = arg.ValueAsRangeInfo;
                return r.GetValue(r.Address._fromRow, r.Address._fromCol);
            }
            else
            {
                return arg==null?null:arg.Value;
            }
        }
        /// <summary>
        /// This functions validates that the supplied <paramref name="arguments"/> contains at least
        /// (the value of) <paramref name="minLength"/> elements. If one of the arguments is an
        /// <see cref="ExcelDataProvider.IRangeInfo">Excel range</see> the number of cells in
        /// that range will be counted as well.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="minLength"></param>
        /// <param name="errorTypeToThrow">The <see cref="eErrorType"/> of the <see cref="ExcelErrorValueException"/> that will be thrown if <paramref name="minLength"/> is not met.</param>
        protected void ValidateArguments(IEnumerable<FunctionArgument> arguments, int minLength,
                                         eErrorType errorTypeToThrow)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            ThrowExcelErrorValueExceptionIf(() =>
                {
                    var nArgs = 0;
                    if (arguments.Any())
                    {
                        foreach (var arg in arguments)
                        {
                            nArgs++;
                            if (nArgs >= minLength) return false;
                            if (arg.IsExcelRange)
                            {
                                nArgs += arg.ValueAsRangeInfo.GetNCells();
                                if (nArgs >= minLength) return false;
                            }
                        }
                    }
                    return true;
                }, errorTypeToThrow);
        }

        /// <summary>
        /// This functions validates that the supplied <paramref name="arguments"/> contains at least
        /// (the value of) <paramref name="minLength"/> elements. If one of the arguments is an
        /// <see cref="ExcelDataProvider.IRangeInfo">Excel range</see> the number of cells in
        /// that range will be counted as well.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="minLength"></param>
        /// <exception cref="ArgumentException"></exception>
        protected void ValidateArguments(IEnumerable<FunctionArgument> arguments, int minLength)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            ThrowArgumentExceptionIf(() =>
                {
                    var nArgs = 0;
                    if (arguments.Any())
                    {
                        foreach (var arg in arguments)
                        {
                            nArgs++;
                            if (nArgs >= minLength) return false;
                            if (arg.IsExcelRange)
                            {
                                nArgs += arg.ValueAsRangeInfo.GetNCells();
                                if (nArgs >= minLength) return false;
                            }
                        }
                    }
                    return true;
                }, "Expecting at least {0} arguments", minLength.ToString());
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// <paramref name="index"/> as an integer.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <returns>Value of the argument as an integer.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected int ArgToInt(IEnumerable<FunctionArgument> arguments, int index)
        {
            var val = arguments.ElementAt(index).ValueFirst;
            return (int)_argumentParsers.GetParser(DataType.Integer).Parse(val);
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// <paramref name="index"/> as a string.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <returns>Value of the argument as a string.</returns>
        protected string ArgToString(IEnumerable<FunctionArgument> arguments, int index)
        {
            var obj = arguments.ElementAt(index).ValueFirst;
            return obj != null ? obj.ToString() : string.Empty;
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// <paramref name="index"/> as a bool
        /// </summary>
        /// <param name="obj"></param>
        /// <returns>Value of the argument as a double.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected double ArgToDecimal(object obj)
        {
            return (double)_argumentParsers.GetParser(DataType.Decimal).Parse(obj);
        }

        /// <summary>
        /// Returns the value of the argument att the position of the 0-based
        /// <paramref name="index"/> as a <see cref="System.Double"/>.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <returns>Value of the argument as an integer.</returns>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected double ArgToDecimal(IEnumerable<FunctionArgument> arguments, int index)
        {
            return ArgToDecimal(arguments.ElementAt(index).Value);
        }

        protected bool IsNumericString(object value)
        {
            if (value == null || string.IsNullOrEmpty(value.ToString())) return false;
            return Regex.IsMatch(value.ToString(), @"^[\d]+(\,[\d])?");
        }

        /// <summary>
        /// If the argument is a boolean value its value will be returned.
        /// If the argument is an integer value, true will be returned if its
        /// value is not 0, otherwise false.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        protected bool ArgToBool(IEnumerable<FunctionArgument> arguments, int index)
        {
            var obj = arguments.ElementAt(index).Value ?? string.Empty;
            return (bool)_argumentParsers.GetParser(DataType.Boolean).Parse(obj);
        }

        /// <summary>
        /// Throws an <see cref="ArgumentException"/> if <paramref name="condition"/> evaluates to true.
        /// </summary>
        /// <param name="condition"></param>
        /// <param name="message"></param>
        /// <exception cref="ArgumentException"></exception>
        protected void ThrowArgumentExceptionIf(Func<bool> condition, string message)
        {
            if (condition())
            {
                throw new ArgumentException(message);
            }
        }

        /// <summary>
        /// Throws an <see cref="ArgumentException"/> if <paramref name="condition"/> evaluates to true.
        /// </summary>
        /// <param name="condition"></param>
        /// <param name="message"></param>
        /// <param name="formats">Formats to the message string.</param>
        protected void ThrowArgumentExceptionIf(Func<bool> condition, string message, params object[] formats)
        {
            message = string.Format(message, formats);
            ThrowArgumentExceptionIf(condition, message);
        }

        /// <summary>
        /// Throws an <see cref="ExcelErrorValueException"/> with the given <paramref name="errorType"/> set.
        /// </summary>
        /// <param name="errorType"></param>
        protected void ThrowExcelErrorValueException(eErrorType errorType)
        {
            throw new ExcelErrorValueException("An excel function error occurred", ExcelErrorValue.Create(errorType));
        }

        /// <summary>
        /// Throws an <see cref="ArgumentException"/> if <paramref name="condition"/> evaluates to true.
        /// </summary>
        /// <param name="condition"></param>
        /// <param name="errorType"></param>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected void ThrowExcelErrorValueExceptionIf(Func<bool> condition, eErrorType errorType)
        {
            if (condition())
            {
                throw new ExcelErrorValueException("An excel function error occurred", ExcelErrorValue.Create(errorType));
            }
        }

        protected bool IsNumeric(object val)
        {
            if (val == null) return false;
            return (val.GetType().IsPrimitive || val is double || val is decimal  || val is System.DateTime || val is TimeSpan);
        }

        //protected virtual bool IsNumber(object obj)
        //{
        //    if (obj == null) return false;
        //    return (obj is int || obj is double || obj is short || obj is decimal || obj is long);
        //}

        /// <summary>
        /// Helper method for comparison of two doubles.
        /// </summary>
        /// <param name="d1"></param>
        /// <param name="d2"></param>
        /// <returns></returns>
        protected bool AreEqual(double d1, double d2)
        {
            return System.Math.Abs(d1 - d2) < double.Epsilon;
        }

        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected virtual IEnumerable<double> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments,
                                                                     ParsingContext context)
        {
            return ArgsToDoubleEnumerable(false, arguments, context);
        }

        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="ignoreErrors">If a cell contains an error, that error will be ignored if this method is set to true</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected virtual IEnumerable<double> ArgsToDoubleEnumerable(bool ignoreHiddenCells, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return _argumentCollectionUtil.ArgsToDoubleEnumerable(ignoreHiddenCells, ignoreErrors, arguments, context);
        }

        /// <summary>
        /// Will return the arguments as an enumerable of doubles.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected virtual IEnumerable<double> ArgsToDoubleEnumerable(bool ignoreHiddenCells, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return ArgsToDoubleEnumerable(ignoreHiddenCells, true, arguments, context);
        }

        /// <summary>
        /// Will return the arguments as an enumerable of objects.
        /// </summary>
        /// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
        /// <param name="arguments"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        protected virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHiddenCells, IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            return _argumentCollectionUtil.ArgsToObjectEnumerable(ignoreHiddenCells, arguments, context);
        }

        /// <summary>
        /// Use this method to create a result to return from Excel functions. 
        /// </summary>
        /// <param name="result"></param>
        /// <param name="dataType"></param>
        /// <returns></returns>
        protected CompileResult CreateResult(object result, DataType dataType)
        {
            var validator = _compileResultValidators.GetValidator(dataType);
            validator.Validate(result);
            return new CompileResult(result, dataType);
        }

        /// <summary>
        /// Use this method to apply a function on a collection of arguments. The <paramref name="result"/>
        /// should be modifyed in the supplied <paramref name="action"/> and will contain the result
        /// after this operation has been performed.
        /// </summary>
        /// <param name="collection"></param>
        /// <param name="result"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        protected virtual double CalculateCollection(IEnumerable<FunctionArgument> collection, double result, Func<FunctionArgument,double,double> action)
        {
            return _argumentCollectionUtil.CalculateCollection(collection, result, action);
        }

        /// <summary>
        /// if the supplied <paramref name="arg">argument</paramref> contains an Excel error
        /// an <see cref="ExcelErrorValueException"/> with that errorcode will be thrown
        /// </summary>
        /// <param name="arg"></param>
        /// <exception cref="ExcelErrorValueException"></exception>
        protected void CheckForAndHandleExcelError(FunctionArgument arg)
        {
            if (arg.ValueIsExcelError)
            {
                throw (new ExcelErrorValueException(arg.ValueAsExcelErrorValue));
            }
        }

        /// <summary>
        /// If the supplied <paramref name="cell"/> contains an Excel error
        /// an <see cref="ExcelErrorValueException"/> with that errorcode will be thrown
        /// </summary>
        /// <param name="cell"></param>
        protected void CheckForAndHandleExcelError(ExcelDataProvider.ICellInfo cell)
        {
            if (cell.IsExcelError)
            {
                throw (new ExcelErrorValueException(ExcelErrorValue.Parse(cell.Value.ToString())));
            }
        }
    }
}
