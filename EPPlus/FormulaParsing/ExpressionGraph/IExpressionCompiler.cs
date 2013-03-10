using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
    public interface IExpressionCompiler
    {
        CompileResult Compile(IEnumerable<Expression> expressions);
    }
}
