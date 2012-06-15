using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    interface ICalcEngine
    {
        Dictionary<string, string> GetFormulas(string address);   //CellID?, Formula
        Dictionary<string, object> GetNameValues(string address);
        Dictionary<string, object> GetWorkbookNameValues();

        object GetValue(int row, int col);
        void SetValue(int row, int col, object value);
    }

}
