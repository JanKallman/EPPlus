using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Calculation
{
    public interface ICalcEngineValueInfo
    {
        object GetValue(int row, int col);

        bool IsHidden(int row, int col);

        void SetFormulaValue(int row, int col, object value);
    }
}
