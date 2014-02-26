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
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class ExcelLookupNavigator : LookupNavigator
    {
        private int _currentRow;
        private int _currentCol;
        private object _currentValue;
        private RangeAddress _rangeAddress;
        private int _index;

        public ExcelLookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
            : base(direction, arguments, parsingContext)
        {
            Initialize();
        }

        private void Initialize()
        {
            _index = 0;
            var factory = new RangeAddressFactory(ParsingContext.ExcelDataProvider);
            if (Arguments.RangeInfo == null)
            {
                _rangeAddress = factory.Create(ParsingContext.Scopes.Current.Address.Worksheet, Arguments.RangeAddress);
            }
            else
            {
                _rangeAddress = factory.Create(Arguments.RangeInfo.Address.WorkSheet, Arguments.RangeInfo.Address.Address);
            }
            _currentCol = _rangeAddress.FromCol;
            _currentRow = _rangeAddress.FromRow;
            SetCurrentValue();
        }

        private void SetCurrentValue()
        {
            _currentValue = ParsingContext.ExcelDataProvider.GetCellValue(_rangeAddress.Worksheet, _currentRow, _currentCol);
        }

        private bool HasNext()
        {
            if (Direction == LookupDirection.Vertical)
            {
                return _currentRow < _rangeAddress.ToRow;
            }
            else
            {
                return _currentCol < _rangeAddress.ToCol;
            }
        }

        public override int Index
        {
            get { return _index; }
        }

        public override bool MoveNext()
        {
            if (!HasNext()) return false;
            if (Direction == LookupDirection.Vertical)
            {
                _currentRow++;
            }
            else
            {
                _currentCol++;
            }
            _index++;
            SetCurrentValue();
            return true;
        }

        public override object CurrentValue
        {
            get { return _currentValue; }
        }

        public override object GetLookupValue()
        {
            var row = _currentRow;
            var col = _currentCol;
            if (Direction == LookupDirection.Vertical)
            {
                col += Arguments.LookupIndex - 1;
                row += Arguments.LookupOffset;
            }
            else
            {
                row += Arguments.LookupIndex - 1;
                col += Arguments.LookupOffset;
            }
            return ParsingContext.ExcelDataProvider.GetCellValue(_rangeAddress.Worksheet, row, col); 
        }
    }
}
