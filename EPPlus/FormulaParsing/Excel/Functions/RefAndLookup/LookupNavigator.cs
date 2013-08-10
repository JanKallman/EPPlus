using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
    public class LookupNavigator
    {
        private readonly LookupDirection _direction;
        private readonly LookupArguments _arguments;
        private readonly ParsingContext _parsingContext;
        private RangeAddress _rangeAddress;
        private int _currentRow;
        private int _currentCol;

        public LookupNavigator(LookupDirection direction, LookupArguments arguments, ParsingContext parsingContext)
        {
            Require.That(arguments).Named("arguments").IsNotNull();
            Require.That(parsingContext).Named("parsingContext").IsNotNull();
            Require.That(parsingContext.ExcelDataProvider).Named("parsingContext.ExcelDataProvider").IsNotNull();
            _direction = direction;
            _arguments = arguments;
            _parsingContext = parsingContext;
            Initialize();
        }

        private void Initialize()
        {
            var factory = new RangeAddressFactory(_parsingContext.ExcelDataProvider);
            _rangeAddress = factory.Create(_arguments.RangeAddress);
            _currentCol = _rangeAddress.FromCol;
            _currentRow = _rangeAddress.FromRow;
            SetCurrentValue();
        }

        private void SetCurrentValue()
        {
            CurrentValue = _parsingContext.ExcelDataProvider.GetCellValue(_rangeAddress.Worksheet, _currentRow, _currentCol);
            //if (cellValue.Value != null)
            //{
            //    CurrentValue = cellValue.Value;
            //}
            ////else if (!string.IsNullOrEmpty(cellValue.Formula))
            ////{
            ////    CurrentValue = _parsingContext.Parser.Parse(cellValue.Formula);
            ////}
            //else
            //{
            //    CurrentValue = null;
            //}
        }

        private bool HasNext()
        {
            if (_direction == LookupDirection.Vertical)
            {
                return _currentRow < _rangeAddress.ToRow;
            }
            else
            {
                return _currentCol < _rangeAddress.ToCol;
            }
        }

        public int Index
        {
            get;
            private set;
        }

        public virtual bool MoveNext()
        {
            if (!HasNext()) return false;
            if (_direction == LookupDirection.Vertical)
            {
                _currentRow++;
            }
            else
            {
                _currentCol++;
            }
            Index++;
            SetCurrentValue();
            return true;
        }

        public object CurrentValue
        {
            get;
            private set;
        }

        public virtual object GetLookupValue()
        {
            var row = _currentRow;
            var col = _currentCol;
            if (_direction == LookupDirection.Vertical)
            {
                col += _arguments.LookupIndex - 1;
                row += _arguments.LookupOffset;
            }
            else
            {
                row += _arguments.LookupIndex - 1;
                col += _arguments.LookupOffset;
            }
            return _parsingContext.ExcelDataProvider.GetCellValue(_rangeAddress.Worksheet, row, col); 
            //var cellValue = _parsingContext.ExcelDataProvider.GetCellValue(row, col);
            //return cellValue != null ? cellValue.Value : null;
        }
    }
}
