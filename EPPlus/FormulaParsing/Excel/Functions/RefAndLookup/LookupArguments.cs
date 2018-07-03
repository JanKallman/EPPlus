﻿/* Copyright (C) 2011  Jan Källman
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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class LookupArguments
    {
        public enum LookupArgumentDataType
        {
            ExcelRange,
            DataArray
        }

        public LookupArguments(IEnumerable<FunctionArgument> arguments, ParsingContext context)
            : this(arguments, new ArgumentParsers(), context)
        {

        }

        public LookupArguments(IEnumerable<FunctionArgument> arguments, ArgumentParsers argumentParsers, ParsingContext context)
        {
            _argumentParsers = argumentParsers;
            SearchedValue = arguments.ElementAt(0).Value;
            var arg1 = arguments.ElementAt(1).Value;
            var dataArray = arg1 as IEnumerable<FunctionArgument>;
            if (dataArray != null)
            {
                DataArray = dataArray;
                ArgumentDataType = LookupArgumentDataType.DataArray;
            }
            else
            {
                //if (arg1 is ExcelDataProvider.INameInfo) arg1 = ((ExcelDataProvider.INameInfo) arg1).Value;
                var rangeInfo = arg1 as ExcelDataProvider.IRangeInfo;
                if (rangeInfo != null)
                {
                    RangeAddress = string.IsNullOrEmpty(rangeInfo.Address.WorkSheet) ? rangeInfo.Address.Address : "'" + rangeInfo.Address.WorkSheet + "'!" + rangeInfo.Address.Address;
                    RangeInfo = rangeInfo;
                    ArgumentDataType = LookupArgumentDataType.ExcelRange;
                }
                else
                {
                    RangeAddress = arg1.ToString();
                    ArgumentDataType = LookupArgumentDataType.ExcelRange;
                }  
            }
            var indexVal = arguments.ElementAt(2);

            if (indexVal.DataType == DataType.ExcelAddress)
            {
                var address = new ExcelAddress(indexVal.Value.ToString());
                var indexObj = context.ExcelDataProvider.GetRangeValue(address.WorkSheet, address._fromRow, address._fromCol);
                LookupIndex = (int) _argumentParsers.GetParser(DataType.Integer).Parse(indexObj);
            }
            else
            {
                LookupIndex = (int)_argumentParsers.GetParser(DataType.Integer).Parse(arguments.ElementAt(2).Value);
            }
            
            if (arguments.Count() > 3)
            {
                RangeLookup = (bool)_argumentParsers.GetParser(DataType.Boolean).Parse(arguments.ElementAt(3).Value);
            }
            else
            {
                RangeLookup = true;
            }
        }

        public LookupArguments(object searchedValue, string rangeAddress, int lookupIndex, int lookupOffset, bool rangeLookup, ExcelDataProvider.IRangeInfo rangeInfo)
        {
            SearchedValue = searchedValue;
            RangeAddress = rangeAddress;
            RangeInfo = rangeInfo;
            LookupIndex = lookupIndex;
            LookupOffset = lookupOffset;
            RangeLookup = rangeLookup;
        }

        private readonly ArgumentParsers _argumentParsers;

        public object SearchedValue { get; private set; }

        public string RangeAddress { get; private set; }

        public int LookupIndex { get; private set; }

        public int LookupOffset { get; private set; }

        public bool RangeLookup { get; private set; }

        public IEnumerable<FunctionArgument> DataArray { get; private set; }

        public ExcelDataProvider.IRangeInfo RangeInfo { get; private set; }

        public LookupArgumentDataType ArgumentDataType { get; private set; } 
    }
}
