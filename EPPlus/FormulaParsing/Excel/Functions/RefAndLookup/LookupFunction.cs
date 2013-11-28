using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public abstract class LookupFunction : ExcelFunction
    {
        private readonly ValueMatcher _valueMatcher;

        public LookupFunction()
            : this(new ValueMatcher())
        {

        }

        public LookupFunction(ValueMatcher valueMatcher)
        {
            _valueMatcher = valueMatcher;
        }

        public override bool IsLookupFuction
        {
            get
            {
                return true;
            }
        }

        protected int IsMatch(object o1, object o2)
        {
            return _valueMatcher.IsMatch(o1, o2);
        }

        protected LookupDirection GetLookupDirection(RangeAddress rangeAddress)
        {
            var nRows = rangeAddress.ToRow - rangeAddress.FromRow;
            var nCols = rangeAddress.ToCol - rangeAddress.FromCol;
            return nCols > nRows ? LookupDirection.Horizontal : LookupDirection.Vertical;
        }

        protected CompileResult Lookup(LookupNavigator navigator, LookupArguments lookupArgs)
        {
            object lastValue = null;
            object lastLookupValue = null;
            int? lastMatchResult = null;
            do
            {
                var matchResult = IsMatch(navigator.CurrentValue, lookupArgs.SearchedValue);
                if (matchResult == 0)
                {
                    return CreateResult(navigator.GetLookupValue(), DataType.String);
                }
                if (lookupArgs.RangeLookup)
                {
                    if (lastValue != null && matchResult > 0 && lastMatchResult < 0)
                    {
                        return CreateResult(lastLookupValue, DataType.String);
                    }
                    lastMatchResult = matchResult;
                    lastValue = navigator.CurrentValue;
                    lastLookupValue = navigator.GetLookupValue();
                }
            }
            while (navigator.MoveNext());

            throw new ExcelFunctionException("Lookupfunction failed to lookup value", ExcelErrorCodes.NoValueAvaliable);
        }
    }
}
