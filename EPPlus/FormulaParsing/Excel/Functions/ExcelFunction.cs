using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using System.Globalization;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
    public abstract class ExcelFunction
    {
        public ExcelFunction()
            : this(new ArgumentCollectionUtil(), new ArgumentParsers())
        {

        }

        public ExcelFunction(ArgumentCollectionUtil argumentCollectionUtil, ArgumentParsers argumentParsers)
        {
            _argumentCollectionUtil = argumentCollectionUtil;
            _argumentParsers = argumentParsers;
        }

        private readonly ArgumentCollectionUtil _argumentCollectionUtil;
        private readonly ArgumentParsers _argumentParsers;

        public abstract CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context);

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

        protected void ValidateArguments(IEnumerable<FunctionArgument> arguments, int minLength)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            ThrowArgumentExceptionIf(() => arguments.Count() < minLength, "Expecting at least {0} arguments", minLength.ToString());
        }

        protected int ArgToInt(IEnumerable<FunctionArgument> arguments, int index)
        {
            var val = arguments.ElementAt(index).Value;
            return (int)_argumentParsers.GetParser(DataType.Integer).Parse(val);
        }

        protected string ArgToString(IEnumerable<FunctionArgument> arguments, int index)
        {
            var obj = arguments.ElementAt(index).Value;
            return obj != null ? obj.ToString() : string.Empty;
        }

        protected double ArgToDecimal(object obj)
        {
            return (double)_argumentParsers.GetParser(DataType.Decimal).Parse(obj);
        }

        protected double ArgToDecimal(IEnumerable<FunctionArgument> arguments, int index)
        {
            return ArgToDecimal(arguments.ElementAt(index).Value);
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

        protected void ThrowArgumentExceptionIf(Func<bool> condition, string message)
        {
            if (condition())
            {
                throw new ArgumentException(message);
            }
        }

        protected void ThrowArgumentExceptionIf(Func<bool> condition, string message, params string[] formats)
        {
            message = string.Format(message, formats);
            ThrowArgumentExceptionIf(condition, message);
        }

        protected void ThrowExcelFunctionException(ExcelErrorCodes code)
        {
            throw new ExcelFunctionException("An excel function error occurred", code);
        }

        protected void ThrowExcelFunctionExceptionIf(Func<bool> condition, ExcelErrorCodes code)
        {
            if (condition())
            {
                throw new ExcelFunctionException("An excel function error occurred", code);
            }
        }

        protected bool IsNumeric(object val)
        {
            if (val == null) return false;
            return (val.GetType().IsPrimitive || val is double || val is decimal || val is System.DateTime || val is TimeSpan);
        }
        //protected double GetNumeric(object value)
        //{
        //    try
        //    {
        //        if ((value.GetType().IsPrimitive || value is double || value is decimal || value is System.DateTime || value is TimeSpan))
        //        {
        //            if (value is System.DateTime)
        //            {
        //                return ((System.DateTime)value).ToOADate();
        //            }
        //            else if (value is TimeSpan)
        //            {
        //                return new System.DateTime(((TimeSpan)value).Ticks).ToOADate();
        //            }
        //            else if (value is bool)
        //            {
        //                return 0;
        //            }
        //            else
        //            {
        //                //if (v is double && double.IsNaN((double)v))
        //                //{
        //                //    return 0;
        //                //}
        //                //else if (v is double && double.IsInfinity((double)v))
        //                //{
        //                //    return "#NUM!";
        //                //}
        //                //else
        //                //{
        //                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        //            }
        //        }
        //        else
        //        {
        //            return 0;
        //        }
        //    }
        //    catch
        //    {
        //        return 0;
        //    }
        //}

        protected virtual IEnumerable<double> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments)
        {
            return _argumentCollectionUtil.ArgsToDoubleEnumerable(arguments);
        }

        protected CompileResult CreateResult(object result, DataType dataType)
        {
            return new CompileResult(result, dataType);
        }

        protected virtual double CalculateCollection(IEnumerable<FunctionArgument> collection, double result, Func<FunctionArgument,double,double> action)
        {
            return _argumentCollectionUtil.CalculateCollection(collection, result, action);
        }
    }
}
