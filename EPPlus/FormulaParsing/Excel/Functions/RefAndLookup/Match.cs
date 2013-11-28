using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class Match : LookupFunction
    {
        private enum MatchType
        {
            ClosestAbove = -1,
            ExactMatch = 0,
            ClosestBelow = 1
        }

        public Match()
            : base(new WildCardValueMatcher())
        {

        }

        public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
        {
            ValidateArguments(arguments, 2);

            var searchedValue = arguments.ElementAt(0).Value;
            var address = ArgToString(arguments, 1);
            var rangeAddressFactory = new RangeAddressFactory(context.ExcelDataProvider);
            var rangeAddress = rangeAddressFactory.Create(address);
            var matchType = GetMatchType(arguments);
            var args = new LookupArguments(searchedValue, address, 0, 0, false);
            var lookupDirection = GetLookupDirection(rangeAddress);
            var navigator = new LookupNavigator(lookupDirection, args, context);
            int? lastMatchResult = default(int?);
            do
            {
                var matchResult = IsMatch(navigator.CurrentValue, searchedValue);
                if (matchType == MatchType.ClosestBelow && matchResult >= 0)
                {
                    if (!lastMatchResult.HasValue && matchResult > 0)
                    {
                        // TODO: error handling. This happens only if the first item is
                        // below the searched value.
                    }
                    var index = matchResult == 0 ? navigator.Index + 1 : navigator.Index;
                    return CreateResult(index, DataType.Integer);
                }
                if (matchType == MatchType.ClosestAbove && matchResult <= 0)
                {
                    if (!lastMatchResult.HasValue && matchResult < 0)
                    {
                        // TODO: error handling. This happens only if the first item is
                        // above the searched value
                    }
                    var index = matchResult == 0 ? navigator.Index + 1 : navigator.Index;
                    return CreateResult(index, DataType.Integer);
                }
                if (matchType == MatchType.ExactMatch && matchResult == 0)
                {
                    return CreateResult(navigator.Index + 1, DataType.Integer);
                }
                lastMatchResult = matchResult;
            }
            while (navigator.MoveNext());
            return CreateResult(null, DataType.Integer);
        }

        private MatchType GetMatchType(IEnumerable<FunctionArgument> arguments)
        {
            var matchType = MatchType.ClosestBelow;
            if (arguments.Count() > 2)
            {
                matchType = (MatchType)ArgToInt(arguments, 2);
            }
            return matchType;
        }
    }
}
