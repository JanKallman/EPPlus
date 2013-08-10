using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing
{
    /// <summary>
    /// This class should be implemented to be able to deliver excel data
    /// to the formula parser.
    /// </summary>
    public abstract class ExcelDataProvider : IDisposable
    {
        /// <summary>
        /// Returns the names of all worksheet names
        /// </summary>
        /// <returns></returns>
        public abstract ExcelNamedRangeCollection GetWorksheetNames();
        /// <summary>
        /// Returns all defined names in a workbook
        /// </summary>
        /// <returns></returns>
        public abstract ExcelNamedRangeCollection GetWorkbookNameValues();
        /// <summary>
        /// Returns values from the required range.
        /// </summary>
        /// <param name="address">An Excel address</param>
        /// <returns>values from the required cells</returns>
        public abstract IEnumerable<object> GetRangeValues(string worksheetName, string address);

        public abstract IEnumerable<object> GetRangeValues(string address);

        public abstract string GetRangeFormula(string worksheetName, int row, int column);


        ///// <summary>
        ///// Returns a single cell value
        ///// </summary>
        ///// <param name="address"></param>
        ///// <returns></returns>
        //public abstract object GetCellValue(int sheetID, string address);

        /// <summary>
        /// Returns a single cell value
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public abstract object GetCellValue(string sheetName, int row, int col);

        ///// <summary>
        ///// Sets the value on the cell
        ///// </summary>
        ///// <param name="address"></param>
        ///// <param name="value"></param>
        //public abstract void SetCellValue(string address, object value);

        /// <summary>
        /// Use this method to free unmanaged resources.
        /// </summary>
        public abstract void Dispose();

        /// <summary>
        /// Max number of columns in a worksheet that the Excel data provider can handle.
        /// </summary>
        public abstract int ExcelMaxColumns { get; }

        /// <summary>
        /// Max number of rows in a worksheet that the Excel data provider can handle
        /// </summary>
        public abstract int ExcelMaxRows { get; }

        public abstract object GetRangeValue(string worksheetName, int row, int column);
    }
}
