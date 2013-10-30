using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace OfficeOpenXml.FormulaParsing
{
    public class EpplusExcelDataProvider : ExcelDataProvider
    {
        public class CellInfo : ICellInfo
        {
            internal ExcelWorksheet _ws;
            CellsStoreEnumerator<object> _values=null;
            int _fromRow, _toRow, _fromCol, _toCol;
            int _cellCount = 0;
            public CellInfo(ExcelWorksheet ws,int fromRow,int toRow, int fromCol, int toCol)                
            {
                _ws = ws;
                _fromRow=fromRow;
                _fromCol=fromCol;
                _toRow=toRow;
                _toCol=toCol;
                _values = new CellsStoreEnumerator<object>(ws._values, _fromRow, _toRow, _fromCol, _toCol);
            }
            public string Address
            {
                get { return _values.CellAddress; }
            }

            public int Row
            {
                get { return _values.Row; }
            }

            public int Column
            {
                get { return _values.Column; }
            }

            public string Formula
            {
                get 
                {
                    return _ws.GetFormula(_values.Row, _values.Column);
                }
            }

            public new object Value
            {
                get { return _values.Value; }
            }

            public bool IsHiddenRow
            {
                get 
                { 
                    var row=_ws._values.GetValue(_values.Row, 0) as ExcelRow;
                    if(row != null)
                    {
                        return row.Hidden || row.Height==0;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            public bool IsEmpty
            {
                get 
                {
                    if (_cellCount > 0)
                    {
                        return true;
                    }
                    else if (_values.Next())
                    {
                        _values.Previous();
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }
            }

            public bool IsMulti
            {
                get 
                {
                    if (_cellCount == 0)
                    {
                        if (_values.Next() && _values.Next())
                        {
                            _values.Reset();
                            return true;
                        }
                        else
                        {
                            _values.Reset();
                            return false;
                        }
                    }
                    else if (_cellCount>1)
                    {
                        return true;
                    }
                    return false;
                }
            }
        
            public ICellInfo Current
            {
                get { return this; }
            }

            public void Dispose()
            {
 	            _values=null;
                _ws=null;
            }

            object System.Collections.IEnumerator.Current
            {
	            get 
                { 
                    return this;
                }
            }

            public bool MoveNext()
            {
                _cellCount++;
                return _values.MoveNext();
            }

            public void Reset()
            {
 	            _values.Init();
            }


            public bool NextCell()
            {
                _cellCount++;
                return _values.MoveNext();
            }

            public IEnumerator<ICellInfo> GetEnumerator()
            {
                return this;
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return this;
            }


            public IList<LexicalAnalysis.Token> Tokens
            {
                get 
                {
                    return _ws._formulaTokens.GetValue(_values.Row, _values.Column);
                }
            }
        }
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

        //internal override CellsStoreEnumerator<object> GetRangeValues(string worksheetName, string address)
        //{
        //    SetCurrentWorksheet(worksheetName);
        //    var addr = new ExcelAddress(worksheetName, address);
        //    if (addr.Table != null)
        //    {
        //        addr.SetRCFromTable(_package, null);
        //    }
        //    var wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
        //    var ws = _package.Workbook.Worksheets[wsName];
        //    return new CellsStoreEnumerator<object>(ws._values,addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
        //}
        internal override ICellInfo GetRange(string worksheet, int row, int column, string address)
        {
            var addr = new ExcelAddress(worksheet, address);
            if (addr.Table != null)
            {
                addr.SetRCFromTable(_package, new ExcelAddressBase(row, column, row, column));
            }
            SetCurrentWorksheet(addr.WorkSheet); 
            var wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
            var ws = _package.Workbook.Worksheets[wsName];
            //return new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            return new CellInfo(ws, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
        }

        internal override IEnumerable<object> GetRangeValues(string address)
        {
            SetCurrentWorksheet(ExcelAddressInfo.Parse(address));
            var addr = new ExcelAddress(address);
            var wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
            var ws = _package.Workbook.Worksheets[wsName];
            return (new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol));
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
    