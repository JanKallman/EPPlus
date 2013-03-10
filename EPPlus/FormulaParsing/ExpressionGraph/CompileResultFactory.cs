using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public class CompileResultFactory
    {
        public virtual CompileResult Create(object obj)
        {
            if (obj == null) return new CompileResult(null, DataType.String);
            if (obj.GetType().Equals(typeof(string)))
            {
                return new CompileResult(obj, DataType.String);
            }
            if (obj.GetType().Equals(typeof(int)))
            {
                return new CompileResult(obj, DataType.Integer);
            }
            if (obj.GetType().Equals(typeof(double)))
            {
                return new CompileResult(obj, DataType.Decimal);
            }
            if (obj.GetType().Equals(typeof(bool)))
            {
                return new CompileResult(obj, DataType.Boolean);
            }
            throw new ArgumentException("Non supported type " + obj.GetType().FullName);
        }
    }
}
