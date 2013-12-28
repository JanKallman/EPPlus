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
            var navigator = LookupNavigatorFactory.Create(lookupDirection, args, context);
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
