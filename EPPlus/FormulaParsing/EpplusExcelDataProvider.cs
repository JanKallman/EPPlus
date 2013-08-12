using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace OfficeOpenXml.FormulaParsing
{
    public class EpplusExcelDataProvider : ExcelDataProvider
    {
        private readonly ExcelPackage _package;
        private ExcelWorksheet _currentWorksheet;
        private RangeAddressFactory _rangeAddressFactory;

        public EpplusExcelDataProvider(ExcelPackage package)
        {
            _package = package;
            _rangeAddressFactory = new RangeAddressFactory(this);
        }

        public override ExcelNamedRangeCollection GetWorksheetNames()
        {
            return _package.Workbook.Worksheets.First().Names;
        }

        public override ExcelNamedRangeCollection GetWorkbookNameValues()
        {
            return _package.Workbook.Names;
        }

        public override IEnumerable<object> GetRangeValues(string worksheetName, string address)
        {
            SetCurrentWorksheet(worksheetName);
            var addr = new ExcelAddress(worksheetName, address);
            var wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
            var ws = _package.Workbook.Worksheets[wsName];
            return (IEnumerable<object>)(new CellsStoreEnumerator<object>(ws._values,addr._fromRow, addr._fromCol, addr._toRow, addr._toCol));
        }

        public override IEnumerable<object> GetRangeValues(string address)
        {
            SetCurrentWorksheet(ExcelAddressInfo.Parse(address));
            var addr = new ExcelAddress(address);
            var wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
            var ws = _package.Workbook.Worksheets[wsName];
            return (IEnumerable<object>)(new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol));
            //return ws.Cells[address];
            //var returnList = new List<ExcelCell>();
            //var addressInfo = ExcelAddressInfo.Parse(address);
            //SetCurrentWorksheet(addressInfo);
            //var range = _currentWorksheet.Cells[addressInfo.AddressOnSheet];
            //foreach (var cell in range)
            //{
            //    returnList.Add(new ExcelCell(cell.Value, cell.Formula, cell.Start.Column, cell.Start.Row));
            //}
            //return returnList;
        }


        public object GetValue(int row, int column)
        {
            return _currentWorksheet._values.GetValue(row, column);
        }

        public bool IsMerged(int row, int column)
        {
            return _currentWorksheet._flags.GetFlagValue(row, column, CellFlags.Merged);
        }

        public bool IsHidden(int row, int column)
        {
            return _currentWorksheet.Column(column).Hidden || _currentWorksheet.Column(column).Width == 0 ||
                   _currentWorksheet.Row(row).Hidden || _currentWorksheet.Row(column).Height == 0;
        }

        //public override ExcelCell GetCellValue(string address)
        //{
        //    var addressInfo = ExcelAddressInfo.Parse(address);
        //    SetCurrentWorksheet(addressInfo);
        //    var cell = _currentWorksheet.Cells[addressInfo.AddressOnSheet].FirstOrDefault();
        //    if (cell != null)
        //    {
        //        return new ExcelCell(cell.Value, cell.Formula, cell.Start.Column, cell.Start.Row);
        //    }
        //    return null;
        //}

        public override object GetCellValue(string sheetName, int row, int col)
        {
            return _package.Workbook.Worksheets[sheetName]._values.GetValue(row, col);
        }

        private void SetCurrentWorksheet(ExcelAddressInfo addressInfo)
        {
            if (addressInfo.WorksheetIsSpecified)
            {
                _currentWorksheet = _package.Workbook.Worksheets[addressInfo.Worksheet];
            }
            else if (_currentWorksheet == null)
            {
                _currentWorksheet = _package.Workbook.Worksheets.First();
            }
        }

        private void SetCurrentWorksheet(string worksheetName)
        {
            if (!string.IsNullOrEmpty(worksheetName))
            {
                _currentWorksheet = _package.Workbook.Worksheets[worksheetName];    
            }
            else
            {
                _currentWorksheet = _package.Workbook.Worksheets.First(); 
            }
            
        }

        //public override void SetCellValue(string address, object value)
        //{
        //    var addressInfo = ExcelAddressInfo.Parse(address);
        //    var ra = _rangeAddressFactory.Create(address);
        //    SetCurrentWorksheet(addressInfo);
        //    //var valueInfo = (ICalcEngineValueInfo)_currentWorksheet;
        //    //valueInfo.SetFormulaValue(ra.FromRow + 1, ra.FromCol + 1, value);
        //    _currentWorksheet.Cells[ra.FromRow + 1, ra.FromCol + 1].Value = value;
        //}

        public override void Dispose()
        {
            _package.Dispose();
        }

        public override int ExcelMaxColumns
        {
            get { return ExcelPackage.MaxColumns; }
        }

        public override int ExcelMaxRows
        {
            get { return ExcelPackage.MaxRows; }
        }

        public override string GetRangeFormula(string worksheetName, int row, int column)
        {
            return _package.Workbook.Worksheets[worksheetName].GetFormula(row, column);
        }

        public override object GetRangeValue(string worksheetName, int row, int column)
        {
            return _package.Workbook.Worksheets[worksheetName].GetValue(row, column);
        }

        public override List<LexicalAnalysis.Token> GetRangeFormulaTokens(string worksheetName, int row, int column)
        {
            return _package.Workbook.Worksheets[worksheetName]._formulaTokens.GetValue(row, column);
        }

        public override bool IsRowHidden(string worksheetName, int row)
        {
            var b = _package.Workbook.Worksheets[worksheetName].Row(row).Height == 0 || 
                    _package.Workbook.Worksheets[worksheetName].Row(row).Hidden;

            return b;
        }
    }
}
    