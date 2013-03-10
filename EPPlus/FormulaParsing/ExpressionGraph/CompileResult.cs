using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class CompileResult
    {
        private static CompileResult _empty = new CompileResult(null, DataType.Empty);
        public static CompileResult Empty
        {
            get { return _empty; }
        }

        public CompileResult(object result, DataType dataType)
        {
            Result = result;
            DataType = dataType;
        }
        public object Result
        {
            get;
            private set;
        }

        public DataType DataType
        {
            get;
            private set;
        }

        public bool IsNumeric
        {
            get { return DataType == DataType.Decimal || DataType == DataType.Integer; }
        }
    }
}
