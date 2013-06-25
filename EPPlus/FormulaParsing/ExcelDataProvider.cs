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
        /// Returns all formulas on a worksheet
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public abstract IDictionary<string, string> GetWorksheetFormulas(string sheetName);

        /// <summary>
        /// Returns all formulas in a workbook
        /// </summary>
        /// <returns></returns>
        public abstract IDictionary<string, string> GetWorkbookFormulas();

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
        public abstract IEnumerable<ExcelCell> GetRangeValues(string address);

        /// <summary>
        /// Returns a single cell value
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public abstract ExcelCell GetCellValue(string address);

        /// <summary>
        /// Returns a single cell value
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public abstract ExcelCell GetCellValue(int row, int col);

        /// <summary>
        /// Sets the value on the cell
        /// </summary>
        /// <param name="address"></param>
        /// <param name="value"></param>
        public abstract void SetCellValue(string address, object value);

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
    }
}
